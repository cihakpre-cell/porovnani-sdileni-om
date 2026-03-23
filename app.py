import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Vyhodnocení sdílení", layout="wide")

st.title("⚡ Vyhodnocení sdílení elektřiny")
st.write("Nahrajte soubory. Pokud se nespárují, podívejte se do 'Diagnostiky souborů' níže.")

uploaded_files = st.file_uploader("Vyberte .xlsx soubory", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    pairs = {}
    file_log = [] # Pro diagnostiku

    for file in uploaded_files:
        fname = file.name.strip()
        # Hledáme dlouhé číslo OM (alespoň 10 číslic)
        match = re.search(r'(\d{10,})', fname)
        
        if match:
            om_number = match.group(1)
            if om_number not in pairs:
                pairs[om_number] = {'pred': None, 'po': None}
            
            # Zjistíme, co je v názvu PŘED číslem
            prefix_part = fname.split(om_number)[0].upper()
            
            # Pokud je před číslem 'S', jde o soubor PO sdílení
            if 'S' in prefix_part:
                pairs[om_number]['po'] = file
                typ = "PO sdílení (S__)"
            else:
                pairs[om_number]['pred'] = file
                typ = "PŘED sdílením"
            
            file_log.append({"Název souboru": fname, "Identifikované OM": om_number, "Typ": typ})
        else:
            file_log.append({"Název souboru": fname, "Identifikované OM": "Nenalezeno", "Typ": "Chyba"})

    # Zobrazení diagnostiky pro uživatele
    with st.expander("🔍 Diagnostika načtených souborů (klikněte pro kontrolu párování)"):
        st.table(pd.DataFrame(file_log))

    kompletni_pary = {om: data for om, data in pairs.items() if data['pred'] and data['po']}
    nekompletni = [om for om, data in pairs.items() if not (data['pred'] and data['po'])]

    if kompletni_pary:
        st.success(f"Nalezeno {len(kompletni_pary)} kompletních dvojic.")
    
    if nekompletni:
        st.warning(f"Chybí jeden do páru pro OM: {', '.join(nekompletni)}")
        with st.informative_column if 'informative_column' in locals() else st.container():
             st.info("Tip: Zkontrolujte v tabulce výše, zda jsou soubory správně rozřazeny do 'PŘED' a 'PO'.")

    if kompletni_pary and st.button("Zpracovat data"):
        with st.spinner("Počítám rozdíly..."):
            vysledky = []
            for om, files in kompletni_pary.items():
                try:
                    # Čtení dat (sloupce D, E, F jsou indexy 3, 4, 5)
                    df_pred = pd.read_excel(files['pred'], usecols=[3, 4, 5], skiprows=1, names=['Datum', 'Cas', 'Hodnota_pred'])
                    df_po = pd.read_excel(files['po'], usecols=[3, 4, 5], skiprows=1, names=['Datum', 'Cas', 'Hodnota_po'])

                    # Vyčištění a vytvoření klíče
                    df_pred = df_pred.dropna(subset=['Datum', 'Cas'])
                    df_po = df_po.dropna(subset=['Datum', 'Cas'])
                    
                    df_pred['Klic'] = df_pred['Datum'].astype(str) + " " + df_pred['Cas'].astype(str)
                    df_po['Klic'] = df_po['Datum'].astype(str) + " " + df_po['Cas'].astype(str)

                    # Výpočet
                    df_merged = pd.merge(df_pred, df_po[['Klic', 'Hodnota_po']], on='Klic', how='inner')
                    df_merged[om] = pd.to_numeric(df_merged['Hodnota_pred'], errors='coerce') - pd.to_numeric(df_merged['Hodnota_po'], errors='coerce')

                    vysledky.append(df_merged[['Datum', 'Cas', om]])
                except Exception as e:
                    st.error(f"Chyba u OM {om}: {e}")

            if vysledky:
                final_df = vysledky[0]
                for df in vysledky[1:]:
                    final_df = pd.merge(final_df, df, on=['Datum', 'Cas'], how='outer')

                st.dataframe(final_df.head(10))
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False)
                
                st.download_button(label="📥 Stáhnout výsledky", data=output.getvalue(), file_name="vysledek_sdileni.xlsx")
