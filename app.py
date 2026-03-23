import streamlit as st
import pandas as pd
import io
import re

# Nastavení stránky
st.set_page_config(page_title="Zpracování OM")

st.header("⚡ Srovnání spotřeby OM")

# Jednoduchý uploader
uploaded_files = st.file_uploader("1. Nahrajte soubory", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    # Párování
    pairs = {}
    for f in uploaded_files:
        om = re.search(r'(\d{10,})', f.name)
        if om:
            id_om = om.group(1)
            if id_om not in pairs: pairs[id_om] = {'pred': None, 'po': None}
            if 'S__' in f.name or f.name.startswith('S'):
                pairs[id_om]['po'] = f
            else:
                pairs[id_om]['pred'] = f

    kompletni = {k: v for k, v in pairs.items() if v['pred'] and v['po']}
    
    if kompletni:
        st.write(f"Nalezeno {len(kompletni)} dvojic.")
        
        # Tlačítko spustí proces, který se vypočítá "v tichosti"
        if st.button("2. Spustit výpočet"):
            all_dfs = []
            
            # Zpracování bez dynamických status barů
            for om_id, files in kompletni.items():
                try:
                    # Načtení - sloupce D, E, F (indexy 3, 4, 5)
                    d1 = pd.read_excel(files['pred'], usecols=[3, 4, 5], skiprows=1, names=['D', 'C', 'V1'])
                    d2 = pd.read_excel(files['po'], usecols=[3, 4, 5], skiprows=1, names=['D', 'C', 'V2'])
                    
                    d1['K'] = d1['D'].astype(str) + d1['C'].astype(str)
                    d2['K'] = d2['D'].astype(str) + d2['C'].astype(str)
                    
                    m = pd.merge(d1.dropna(subset=['D']), d2[['K', 'V2']].dropna(), on='K')
                    m[om_id] = pd.to_numeric(m['V1'], errors='coerce') - pd.to_numeric(m['V2'], errors='coerce')
                    
                    all_dfs.append(m[['D', 'C', om_id]])
                except:
                    continue

            if all_dfs:
                # Slučování
                final = all_dfs[0]
                for n in all_dfs[1:]:
                    final = pd.merge(final, n, on=['D', 'C'], how='outer')
                
                final = final.rename(columns={'D': 'Datum', 'C': 'Cas'}).sort_values(['Datum', 'Cas'])

                # Příprava Excelu
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final.to_excel(writer, index=False)
                
                # Zobrazení tlačítka pro stažení jako jediný nový prvek
                st.download_button("3. STÁHNOUT VÝSLEDEK", output.getvalue(), "vysledek.xlsx")
