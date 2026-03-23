import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Vyhodnocení sdílení", layout="wide")

st.title("⚡ Vyhodnocení sdílení elektřiny")

uploaded_files = st.file_uploader("Vyberte .xlsx soubory", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    pairs = {}
    for file in uploaded_files:
        fname = file.name.strip()
        match = re.search(r'(\d{10,})', fname)
        if match:
            om_number = match.group(1)
            if om_number not in pairs:
                pairs[om_number] = {'pred': None, 'po': None}
            prefix_part = fname.split(om_number)[0].upper()
            if 'S' in prefix_part:
                pairs[om_number]['po'] = file
            else:
                pairs[om_number]['pred'] = file

    kompletni_pary = {om: data for om, data in pairs.items() if data['pred'] and data['po']}
    
    if kompletni_pary:
        st.success(f"Nalezeno {len(kompletni_pary)} kompletních dvojic.")
        
        if st.button("🚀 Spustit výpočet a připravit soubor ke stažení"):
            vysledky = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            total = len(kompletni_pary)
            
            for i, (om, files) in enumerate(kompletni_pary.items()):
                status_text.text(f"Zpracovávám OM {om} ({i+1}/{total})")
                try:
                    # Načtení dat
                    df_pred = pd.read_excel(files['pred'], usecols=[3, 4, 5], skiprows=1, names=['Datum', 'Cas', 'Hodnota_pred'])
                    df_po = pd.read_excel(files['po'], usecols=[3, 4, 5], skiprows=1, names=['Datum', 'Cas', 'Hodnota_po'])

                    df_pred = df_pred.dropna(subset=['Datum', 'Cas'])
                    df_po = df_po.dropna(subset=['Datum', 'Cas'])
                    
                    # Optimalizace klíče
                    df_pred['Klic'] = df_pred['Datum'].astype(str) + " " + df_pred['Cas'].astype(str)
                    df_po['Klic'] = df_po['Datum'].astype(str) + " " + df_po['Cas'].astype(str)

                    df_merged = pd.merge(df_pred, df_po[['Klic', 'Hodnota_po']], on='Klic', how='inner')
                    
                    # Výpočet rozdílu
                    df_merged[om] = pd.to_numeric(df_merged['Hodnota_pred'], errors='coerce') - pd.to_numeric(df_merged['Hodnota_po'], errors='coerce')

                    vysledky.append(df_merged[['Datum', 'Cas', om]])
                except Exception as e:
                    st.error(f"Chyba u OM {om}: {e}")
                
                progress_bar.progress((i + 1) / total)

            if vysledky:
                status_text.text("Slučuji všechna data do jednoho souboru... (může chvíli trvat)")
                final_df = vysledky[0]
                for df in vysledky[1:]:
                    final_df = pd.merge(final_df, df, on=['Datum', 'Cas'], how='outer')

                # Seřazení podle času
                final_df = final_df.sort_values(['Datum', 'Cas'])

                st.success("Hotovo! Tabulka je připravena.")
                
                # Export bez zobrazení velké tabulky na webu (prevence chyby removeChild)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 STÁHNOUT VÝSLEDNÝ EXCEL", 
                    data=output.getvalue(), 
                    file_name="vysledek_sdileni_komplet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                status_text.empty()
                progress_bar.empty()

    if not kompletni_pary:
        st.info("Nahrajte soubory pro spárování.")
