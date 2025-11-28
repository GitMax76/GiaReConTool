import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# Titolo della pagina
st.title("📄 Convertitore PDF a Excel")
st.write("Carica il tuo file PDF per convertirlo ed estrarre la tabella.")

# Widget per caricare il file
uploaded_file = st.file_uploader("Scegli il file PDF", type="pdf")

if uploaded_file is not None:
    st.info("File caricato! Elaborazione in corso...")
    
    all_data = []
    
    # pdfplumber può aprire direttamente l'oggetto file caricato
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()

            if table:
                for row in table:
                    clean_row = [cell.strip() if cell else "" for cell in row]
                    
                    if not clean_row or clean_row[0].lower() == "tipo":
                        continue
                    
                    # --- LA TUA LOGICA DI ESTRAZIONE ORIGINALE ---
                    tipo = clean_row[0]
                    mittente = clean_row[1]
                    oggetto_data_raw = clean_row[2]
                    allegati = clean_row[3] if len(clean_row) > 3 else ""
                    esito = clean_row[4] if len(clean_row) > 4 else ""

                    if len(clean_row) > 5: 
                         oggetto = clean_row[2]
                         data = clean_row[3]
                         allegati = clean_row[4]
                         esito = clean_row[5]
                    else:
                        date_pattern = r"(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})"
                        match = re.search(date_pattern, oggetto_data_raw)
                        
                        if match:
                            data = match.group(1)
                            oggetto = oggetto_data_raw.replace(data, "").strip()
                        else:
                            data = ""
                            oggetto = oggetto_data_raw

                    all_data.append({
                        "Tipo": tipo,
                        "Mittente": mittente,
                        "Oggetto": oggetto,
                        "Data Invio": data,
                        "Allegati": allegati,
                        "Esito Controllo Messaggio": esito
                    })

    # Creazione DataFrame
    if all_data:
        df = pd.DataFrame(all_data)
        df = df[df["Tipo"] != ""] # Pulizia righe vuote

        # Mostra un'anteprima a video
        st.write("### Anteprima Dati:")
        st.dataframe(df.head())

        # Preparazione del file Excel in memoria (senza salvarlo su disco)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
            
        # Bottone per scaricare
        st.download_button(
            label="📥 Scarica file Excel",
            data=buffer,
            file_name="risultato_convertito.xlsx",
            mime="application/vnd.ms-excel"
        )
        st.success("Conversione completata!")
    else:
        st.warning("Non sono riuscito a trovare tabelle valide in questo PDF.")
