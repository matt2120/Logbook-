import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.set_page_config(layout="wide")

st.markdown("""
    <style>
    .stMultiSelect > div {
        max-width: 100% !important;
    }
    div[data-baseweb="tag"] {
        max-width: 100% !important;
        white-space: normal !important;
        overflow-wrap: break-word !important;
        word-break: break-word !important;
        font-size: 14px !important;
        padding: 6px 10px !important;
    }
    div[data-baseweb="select"] > div {
        min-height: 55px !important;
    }
    .stMultiSelect label {
        white-space: normal !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("Logbook unità")

uploaded_file = st.file_uploader("Carica il file Excel dei cambi:", type="xlsx")

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    df_cambi = xls.parse("Matrice Cambio")
    df_cambi.columns = df_cambi.columns.str.strip()

    df_unita = None
    for sheet in xls.sheet_names:
        temp_df = xls.parse(sheet)
        temp_df.columns = temp_df.columns.str.strip()
        if 'Nome Identificativo' in temp_df.columns:
            df_unita = temp_df
            break

    if df_unita is None:
        st.error("Errore: Nessun foglio contiene la colonna 'Nome Identificativo'.")
    else:
        giorni_settimana = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"]

        base_lunedi = datetime(2025, 4, 7)
        today = datetime.today()
        giorni_differenza = (today - base_lunedi).days
        numero_settimana = giorni_differenza // 7
        lunedi_corrente = base_lunedi + timedelta(weeks=numero_settimana)
        lunedi_successivo = lunedi_corrente + timedelta(weeks=1)

        settimana_1 = [lunedi_corrente + timedelta(days=i) for i in range(5)]
        settimana_2 = [lunedi_successivo + timedelta(days=i) for i in range(5)]

        calendario = {}

        def sort_key(cambio):
            for tier in ["S1", "S2", "S3", "S5"]:
                if cambio.startswith(tier):
                    return ["S1", "S2", "S3", "S5"].index(tier)
            return 999

        def settimana_layout(settimana, label):
            st.markdown(f"### 📆 {label}")
            cols = st.columns(5)
            sorted_cambi = sorted(df_cambi["Cambio"].unique(), key=sort_key)
            for i, giorno in enumerate(settimana):
                nome_giorno = giorno.strftime("%A").capitalize()
                giorno_label = f"{label} - {nome_giorno}"
                with cols[i]:
                    cambi_giorno = st.multiselect(
                        f"{nome_giorno} ({giorno.strftime('%d/%m')})",
                        options=sorted_cambi,
                        key=f"cambi_{giorno_label}",
                        label_visibility="visible"
                    )
                    calendario[giorno] = cambi_giorno

        st.subheader("🗓️ Calendario Settimanale Interno")
        settimana_layout(settimana_1, "Settimana 1")
        settimana_layout(settimana_2, "Settimana 2")

        if st.button("Genera Logbook"):
            schedule_preparazione = []
            schedule_pulizia = []
            messaggi_speciali = []

            def clean(s):
                if pd.isna(s):
                    return None
                return str(s).replace(")", ") ").strip().upper()

            for data_montaggio, cambi in calendario.items():
                for cambio in cambi:
                    righe_cambio = df_cambi[df_cambi['Cambio'] == cambio]
                    for _, riga in righe_cambio.iterrows():
                        unita_montare = riga.get('Testata da montare')
                        unita_smontare = riga.get('Testata da smontare')

                        if str(unita_montare).strip() == '/' and str(unita_smontare).strip() == '/':
                            messaggi_speciali.append(f"Per il cambio {cambio} non sono previsti cambi unità")
                            continue

                        if pd.notna(unita_montare):
                            if "FB-GE o FB-PI se montate" in str(unita_montare):
                                schedule_preparazione.append({
                                    "Linea": "",
                                    "Unità": "FB-GE o FB-PI se montate",
                                    "Data Preparazione": (data_montaggio - timedelta(days=1)).strftime("%d %B"),
                                    "Tempo Preparazione": ""
                                })
                            else:
                                filtro_mont = df_unita[df_unita['Nome Identificativo'].apply(lambda x: clean(x) == clean(unita_montare))]
                                if not filtro_mont.empty:
                                    dett_mont = filtro_mont.iloc[0]
                                    tempo_val = dett_mont.get('Tempo di prep')
                                    tempo_prep = f"{int(tempo_val)} min" if pd.notna(tempo_val) else "-"
                                    schedule_preparazione.append({
                                        "Linea": dett_mont['Linea'],
                                        "Unità": unita_montare,
                                        "Data Preparazione": (data_montaggio - timedelta(days=1)).strftime("%d %B"),
                                        "Tempo Preparazione": tempo_prep
                                    })
                                else:
                                    st.warning(f"Unità non trovata (montaggio): {unita_montare}")

                        if pd.notna(unita_smontare):
                            filtro_smnt = df_unita[df_unita['Nome Identificativo'].apply(lambda x: clean(x) == clean(unita_smontare))]
                            if not filtro_smnt.empty:
                                dett_smnt = filtro_smnt.iloc[0]
                                tempo_val = dett_smnt.get('Tempo di pulizia')
                                tempo_pulizia = f"{int(tempo_val)} min" if pd.notna(tempo_val) else "-"
                                schedule_pulizia.append({
                                    "Linea": dett_smnt['Linea'],
                                    "Unità": unita_smontare,
                                    "Data Pulizia": data_montaggio.strftime("%d %B"),
                                    "Tempo Pulizia": tempo_pulizia
                                })
                            else:
                                st.warning(f"Unità non trovata (smontaggio): {unita_smontare}")

            def color_row_by_linea(row):
                colori = {
                    "FGC1": "background-color: #ADD8E6",
                    "FGC2": "background-color: #E6A57E",
                    "FGC3": "background-color: #F3D1FF",
                }
                return [colori.get(row['Linea'], '')] * len(row)

            st.subheader("Preparazione Pre Montaggio")
            df_prep = pd.DataFrame(schedule_preparazione)
            df_prep.index += 1
            st.dataframe(df_prep.style.apply(color_row_by_linea, axis=1))

            st.subheader("Pulizia Post Smontaggio")
            df_pulizia = pd.DataFrame(schedule_pulizia)
            df_pulizia.index += 1
            st.dataframe(df_pulizia.style.apply(color_row_by_linea, axis=1))

            st.subheader("⏱️ Riepilogo Tempi Totali per Giorno")
            df_prep['Tempo (int)'] = df_prep['Tempo Preparazione'].str.extract(r'(\d+)').astype(float)
            df_pulizia['Tempo (int)'] = df_pulizia['Tempo Pulizia'].str.extract(r'(\d+)').astype(float)

            giorni = sorted(set(df_prep['Data Preparazione']).union(df_pulizia['Data Pulizia']))
            totali_prep, totali_puli, totali = [], [], []
            max_val, min_val = 0, float('inf')
            numeri_totali = []

            for giorno in giorni:
                t_prep = df_prep[df_prep['Data Preparazione'] == giorno]['Tempo (int)'].sum()
                t_puli = df_pulizia[df_pulizia['Data Pulizia'] == giorno]['Tempo (int)'].sum()
                total = t_prep + t_puli
                numeri_totali.append(total)
                max_val = max(max_val, total)
                min_val = min(min_val, total) if total > 0 else min_val
                totali_prep.append(f"{int(t_prep)} min" if t_prep else "-")
                totali_puli.append(f"{int(t_puli)} min" if t_puli else "-")
                totali.append(total)

            df_riepilogo = pd.DataFrame({
                'Giorno': ['Totale Preparazione', 'Totale Pulizia', 'Totale Giorno']
            })
            for i, giorno in enumerate(giorni):
                df_riepilogo[giorno] = [totali_prep[i], totali_puli[i], f"{int(totali[i])} min" if totali[i] else "-"]

            df_riepilogo = df_riepilogo.set_index('Giorno')

            def background_gradient(val):
                if isinstance(val, str) and 'min' in val:
                    num = int(val.replace(' min', ''))
                    norm = (num - min_val) / (max_val - min_val) if max_val != min_val else 0.5
                    if norm <= 0.2:
                        return 'background-color: #c6e2b3'  # verde chiaro
                    elif norm <= 0.4:
                        return 'background-color: #fff8b0'  # giallo chiaro
                    elif norm <= 0.6:
                        return 'background-color: #ffe080'  # giallo medio
                    elif norm <= 0.8:
                        return 'background-color: #ffb347'  # arancio chiaro
                    else:
                        return 'background-color: #e57373'  # rosso tenue
                return ''

            styled = df_riepilogo.style.applymap(background_gradient, subset=pd.IndexSlice[['Totale Giorno'], :])
            styled.set_table_styles([{ 'selector': 'th.col_heading', 'props': [('font-weight', 'bold')] }])
            st.dataframe(styled)

            if messaggi_speciali:
                st.subheader("ℹ️ Note sui cambi senza unità")
                for msg in messaggi_speciali:
                    st.info(msg)

            with st.expander("📥 Esporta Logbook in Excel"):
                import io
                from pandas import ExcelWriter
                buffer = io.BytesIO()
                with ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_prep.drop(columns='Tempo (int)', errors='ignore').to_excel(writer, sheet_name='Preparazione', index=False)
                    df_pulizia.drop(columns='Tempo (int)', errors='ignore').to_excel(writer, sheet_name='Pulizia', index=False)
                    df_riepilogo.to_excel(writer, sheet_name='Riepilogo')
                st.download_button(
                    label="Scarica Logbook",
                    data=buffer,
                    file_name="logbook_settimanale.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

