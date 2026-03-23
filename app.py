import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Vyhodnocení sdílení", layout="centered")

st.title("⚡ Vyhodnocení sdílení")

# Uploader
uploaded_files = st.file_uploader("Nahrajte .xlsx soubory", accept_multiple_files=True, type=['xlsx'], key="file_uploader")

if uploaded_files:
    # 1. Párování souborů
    pairs = {}
    for file in uploaded_files:
        fname = file.name.strip()
        match = re.search(r'(\d{10,})', fname)
        if match:
            om = match.group(1)
            if om not in pairs: pairs[om] = {'pred': None, 'po': None}
            # Pokud je v názvu před číslem 'S', je to PO sdílení
            if 'S' in fname.split(om)[0].upper():
                pairs[om]['po'] = file
            else:
                pairs[om]['pred'] = file

    kompletni = {om: d for om, d in pairs.items() if d['pred'] and d['po']}
    
    if kompletni:
        st.success(f"Připraveno {len(kompletni)} dvojic ke zpracování.")
        
        if st.button("📊 Zpracovat a vygenerovat Excel"):
            vysledky = []
            msg = st.empty() # Prostor pro stavové zprávy
            
            for i, (om, files) in enumerate(kompletni.items()):
                msg.text(f"Zpracovávám: {i+1}/{len(kompletni)} (OM {om})")
                try:
                    # Načítáme pouze sloupce D, E, F (indexy 3, 4, 5)
                    # Používáme dtype string pro datum/čas pro maximální stabilitu při spojování
                    d1 = pd.read_excel(files['pred'], usecols=[3, 4, 5], skiprows=1, names=['D', 'C', 'V1'])
                    d2 = pd.read_excel(files['po'], usecols=[3, 4, 5], skiprows=1, names=['D', 'C', 'V2'])
                    
                    # Rychlé pročištění
                    d1 = d1.dropna(subset=['D', 'C'])
                    d2 = d2.dropna(subset=['D', 'C'])
                    
                    # Unikátní klíč pro spojení
                    d1['K'] = d1['D'].astype(str) + d1['C'].astype(str)
                    d2['K'] = d2['D'].astype(str) + d2['C'].astype(str)
                    
                    # Spojení a výpočet (rozdíl)
                    m = pd.merge(d1, d2[['K', 'V2']], on='K', how='inner')
                    m[om] = pd.to_numeric(m['V1'], errors='coerce') - pd.to_numeric(m['V2'], errors='coerce')
                    
                    vysledky.append(m[['D', 'C', om]])
                except Exception as e:
                    st.error(f"Chyba u {om}: {str(e)}")

            if vysledky:
                msg.text("Slučování dat do finální tabulky...")
                # Sloučení všech OM do jedné tabulky podle Data a Času
                final = vysledky[0]
                for next_df in vysledky[1:]:
                    final = pd.merge(final, next_df, on=['D', 'C'], how='outer')
                
                final = final.rename(columns={'D': 'Datum', 'C': 'Čas'}).sort_values(['Datum', 'Čas'])

                # Generování Excelu do paměti
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final.to_excel(writer, index=False, sheet_name='Nasdíleno')
                
                msg.empty()
                st.balloons()
                st.download_button(
                    label="✅ STÁHNOUT VÝSLEDNÝ EXCEL",
                    data=output.getvalue(),
                    file_name="vysledek_sdileni.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("Nahrajte soubory. Čekám na kompletní dvojice (originál + S__verze).")
