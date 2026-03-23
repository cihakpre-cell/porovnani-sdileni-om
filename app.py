import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Vyhodnocení sdílení", layout="wide")

st.title("⚡ Vyhodnocení sdílení elektřiny po čtvrthodinách")
st.write("Nahrajte všechny excelové soubory (před sdílením i po sdílení najednou). Aplikace je sama spáruje podle čísla OM, vypočítá rozdíl a vytvoří souhrnný Excel.")

# Uploader pro více souborů
uploaded_files = st.file_uploader("Vyberte .xlsx soubory (až 66 souborů)", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    # 1. Rozřazení souborů do slovníku podle čísla OM
    pairs = {}
    for file in uploaded_files:
        # Najdeme číslo OM v názvu souboru (jakákoliv delší sekvence čísel)
        match = re.search(r'(\d+)', file.name)
        if match:
            om_number = match.group(1)
            
            if om_number not in pairs:
                pairs[om_number] = {'pred': None, 'po': None}
                
            # Pokud soubor začíná na S__, jde o data PO sdílení
            if file.name.startswith('S__'):
                pairs[om_number]['po'] = file
            else:
                pairs[om_number]['pred'] = file

    # Kontrola, kolik kompletních párů se našlo
    kompletni_pary = {om: data for om, data in pairs.items() if data['pred'] and data['po']}
    nekompletni = [om for om, data in pairs.items() if not (data['pred'] and data['po'])]

    st.info(f"Nalezeno {len(kompletni_pary)} kompletních dvojic souborů.")
    if nekompletni:
        st.warning(f"Pro tato OM chybí jeden do páru: {', '.join(nekompletni)}")

    # 2. Zpracování dat
    if kompletni_pary and st.button("Zpracovat data a vypočítat sdílení"):
        with st.spinner("Zpracovávám soubory..."):
            vysledky = []
            
            for om, files in kompletni_pary.items():
                try:
                    # Načteme sloupce D, E, F (indexy 3, 4, 5). 
                    # skiprows=1 přeskočí první informační řádek, aby se správně chytly hodnoty
                    df_pred = pd.read_excel(files['pred'], usecols="D:F", skiprows=1, names=['Od_data', 'Od_casu', 'Hodnota_pred'])
                    df_po = pd.read_excel(files['po'], usecols="D:F", skiprows=1, names=['Od_data', 'Od_casu', 'Hodnota_po'])

                    # Očištění o případné prázdné řádky na konci
                    df_pred = df_pred.dropna(subset=['Od_data', 'Od_casu'])
                    df_po = df_po.dropna(subset=['Od_data', 'Od_casu'])

                    # Vytvoření pomocného klíče pro spojení
                    df_pred['Klic'] = df_pred['Od_data'].astype(str) + " " + df_pred['Od_casu'].astype(str)
                    df_po['Klic'] = df_po['Od_data'].astype(str) + " " + df_po['Od_casu'].astype(str)

                    # Spojení tabulek pro dané OM
                    df_merged = pd.merge(df_pred, df_po[['Klic', 'Hodnota_po']], on='Klic', how='inner')

                    # Převod na čísla (pro jistotu, kdyby tam byl text) a výpočet nasdíleného množství
                    df_merged['Hodnota_pred'] = pd.to_numeric(df_merged['Hodnota_pred'], errors='coerce').fillna(0)
                    df_merged['Hodnota_po'] = pd.to_numeric(df_merged['Hodnota_po'], errors='coerce').fillna(0)
                    
                    # Výpočet (před sdílením - po sdílení = co se nasdílelo)
                    df_merged[om] = df_merged['Hodnota_pred'] - df_merged['Hodnota_po']

                    # Ponecháme jen datum, čas a vypočtený rozdíl
                    df_final_om = df_merged[['Od_data', 'Od_casu', om]]
                    vysledky.append(df_final_om)
                    
                except Exception as e:
                    st.error(f"Chyba při zpracování OM {om}: {e}")

            # 3. Spojení všech výsledků do jedné velké tabulky
            if vysledky:
                # Začneme první tabulkou
                finalni_tabulka = vysledky[0]
                # A postupně k ní připojíme všechny ostatní podle data a času
                for df in vysledky[1:]:
                    finalni_tabulka = pd.merge(finalni_tabulka, df, on=['Od_data', 'Od_casu'], how='outer')

                st.success("Zpracování dokončeno!")
                
                # Náhled
                st.write("Náhled prvních 10 řádků výsledku:")
                st.dataframe(finalni_tabulka.head(10))

                # 4. Export do Excelu
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    finalni_tabulka.to_excel(writer, index=False, sheet_name='Nasdilen_Mnozstvi')
                
                excel_data = output.getvalue()

                st.download_button(
                    label="📥 Stáhnout výsledný Excel",
                    data=excel_data,
                    file_name="nasdileno_po_ctvrthodinach.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )