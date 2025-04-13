import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

st.set_page_config(layout="wide")

oggi = datetime.today().strftime("%A %d %B %Y")
st.title("📅 Logbook unità")
st.caption(f"Oggi: {oggi}")

uploaded_file = st.file_uploader("Carica il file Excel Backend:", type="xlsx")

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    df_cambi = xls.parse("Matrice Cambio")
    df_cambi.columns = df_cambi.columns.str.strip()
    df_cambi['Cambio'] = df_cambi['Cambio'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()

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
        linee = ["FGC1", "FGC2", "FGC3"]
        tier_map = {
            "FGC1": ["S1 Seta", "S1 Idea", "S1 Petalo", "S1 Natura", "S1 Seta Lady"],
            "FGC2": ["S3 Seta", "S3 Idea", "S3 Petalo", "S3 Twiggy"],
            "FGC3": ["S2 Seta", "S2 Idea", "S2 Petalo", "S5 Seta", "S5 Idea", "S5 Petalo"]
        }

        today = datetime.today()
        base_lunedi = today - timedelta(days=today.weekday())
        settimana_1 = [base_lunedi + timedelta(days=i) for i in range(5)]
        settimana_2 = [base_lunedi + timedelta(days=7+i) for i in range(5)]
        tutte_le_date = settimana_1 + settimana_2

        calendario = {linea: {} for linea in linee}

        st.subheader("📋 Seleziona i cambi per ogni linea")
        for linea in linee:
            with st.expander(f" {linea}"):
                tier_possibili = tier_map[linea]
                tier_iniziale = st.selectbox(f"Seleziona il Tier iniziale per {linea}", tier_possibili, key=f"tier_iniziale_{linea}")
                tier_corrente = tier_iniziale

                st.markdown("**Settimana 1**")
                cols1 = st.columns(5)
                for i, giorno in enumerate(settimana_1):
                    giorno_str = giorno.strftime("%a %d/%m")
                    with cols1[i]:
                        opzioni = ["" ] + [t for t in tier_possibili if t != tier_corrente]
                        tier_scelto = st.selectbox(f"{giorno_str}", options=opzioni, key=f"{linea}_settimana1_{giorno_str}")
                        calendario[linea][giorno] = (tier_corrente, tier_scelto if tier_scelto else None)
                        if tier_scelto:
                            tier_corrente = tier_scelto

                st.markdown("**Settimana 2**")
                cols2 = st.columns(5)
                for i, giorno in enumerate(settimana_2):
                    giorno_str = giorno.strftime("%a %d/%m")
                    with cols2[i]:
                        opzioni = ["" ] + [t for t in tier_possibili if t != tier_corrente]
                        tier_scelto = st.selectbox(f"{giorno_str}", options=opzioni, key=f"{linea}_settimana2_{giorno_str}")
                        calendario[linea][giorno] = (tier_corrente, tier_scelto if tier_scelto else None)
                        if tier_scelto:
                            tier_corrente = tier_scelto

        if st.button("Genera Logbook"):
            schedule_preparazione = []
            schedule_pulizia = []
            messaggi_speciali = []

            def clean(s):
                if pd.isna(s):
                    return None
                return str(s).replace(")", ") ").strip().upper()

            for linea in linee:
                for giorno, (tier_da, tier_a) in calendario[linea].items():
                    if not tier_a:
                        continue
                    cambio_cercato = f"{tier_da} > {tier_a}"
                    cambio_cercato = re.sub(r'\s+', ' ', cambio_cercato).strip()
                    cambio_match = df_cambi[df_cambi['Cambio'] == cambio_cercato]
                    if cambio_match.empty:
                        st.warning(f"Cambio non trovato per {linea}: {tier_da} > {tier_a}")
                        continue

                    for _, riga in cambio_match.iterrows():
                        unita_montare = riga.get('Testata da montare')
                        unita_smontare = riga.get('Testata da smontare')

                        if str(unita_montare).strip() == '/' and str(unita_smontare).strip() == '/':
                            messaggi_speciali.append(f"{linea} - {giorno.strftime('%d/%m')}: nessun cambio previsto per {tier_da} > {tier_a}")
                            continue

                        giorno_montaggio = giorno.strftime("%d %B")
                        giorno_preparazione_data = giorno - timedelta(days=3) if giorno.weekday() == 0 else giorno - timedelta(days=1)
                        giorno_preparazione = giorno_preparazione_data.strftime("%d %B")

                        if pd.notna(unita_montare):
                            filtro_mont = df_unita[df_unita['Nome Identificativo'].apply(lambda x: clean(x) == clean(unita_montare))]
                            if not filtro_mont.empty:
                                dett_mont = filtro_mont.iloc[0]
                                tempo_val = dett_mont.get('Tempo di prep')
                                tempo_prep = f"{int(tempo_val)} min" if pd.notna(tempo_val) else "-"
                                schedule_preparazione.append({
                                    "Linea": linea,
                                    "Unità": unita_montare,
                                    "Data Preparazione": giorno_preparazione,
                                    "Priorità Montaggio": giorno_montaggio,
                                    "Tempo Preparazione": tempo_prep
                                })
                        if pd.notna(unita_smontare):
                            filtro_smnt = df_unita[df_unita['Nome Identificativo'].apply(lambda x: clean(x) == clean(unita_smontare))]
                            if not filtro_smnt.empty:
                                dett_smnt = filtro_smnt.iloc[0]
                                tempo_val = dett_smnt.get('Tempo di pulizia')
                                tempo_pulizia = f"{int(tempo_val)} min" if pd.notna(tempo_val) else "-"
                                schedule_pulizia.append({
                                    "Linea": linea,
                                    "Unità": unita_smontare,
                                    "Data Pulizia": giorno.strftime("%d %B"),
                                    "Priorità Rimontaggio": "",
                                    "Tempo Pulizia": tempo_pulizia
                                })

            df_prep = pd.DataFrame(schedule_preparazione)
            df_pulizia = pd.DataFrame(schedule_pulizia)

            df_pulizia['Tempo (int)'] = df_pulizia['Tempo Pulizia'].str.extract(r'(\d+)').astype(float)
            df_prep['Tempo (int)'] = df_prep['Tempo Preparazione'].str.extract(r'(\d+)').astype(float)

            unita_to_montaggi = df_prep.groupby('Unità')['Priorità Montaggio'].apply(list).to_dict()

            for idx, row in df_pulizia.iterrows():
                unita = row['Unità']
                data_pulizia = datetime.strptime(row['Data Pulizia'], "%d %B")
                date_montaggi = [datetime.strptime(d, "%d %B") for d in unita_to_montaggi.get(unita, []) if datetime.strptime(d, "%d %B") > data_pulizia]
                if date_montaggi:
                    prox = min(date_montaggi)
                    df_pulizia.at[idx, 'Priorità Rimontaggio'] = prox.strftime("%d %B")
                else:
                    df_pulizia.at[idx, 'Priorità Rimontaggio'] = "next week"

            def color_row_by_linea(row):
                colori = {
                    "FGC1": "background-color: #ADD8E6",
                    "FGC2": "background-color: #E6A57E",
                    "FGC3": "background-color: #F3D1FF",
                }
                return [colori.get(row['Linea'], '')] * len(row)

            st.subheader("Preparazione pre Montaggio")
            df_prep.index += 1
            st.dataframe(df_prep.drop(columns='Tempo (int)').style.apply(color_row_by_linea, axis=1), use_container_width=True)

            st.subheader("Pulizia Post Smontaggio")
            df_pulizia.index += 1
            st.dataframe(df_pulizia[['Linea', 'Unità', 'Data Pulizia', 'Priorità Rimontaggio', 'Tempo Pulizia']].style.apply(color_row_by_linea, axis=1), use_container_width=True)

            st.subheader("⏱️ Riepilogo Tempi Totali per Giorno")
            giorni = sorted(set(df_prep['Data Preparazione']).union(df_pulizia['Data Pulizia']))
            totali_prep, totali_puli, totali = [], [], []
            max_val, min_val = 0, float('inf')

            for giorno in giorni:
                t_prep = df_prep[df_prep['Data Preparazione'] == giorno]['Tempo (int)'].sum()
                t_puli = df_pulizia[df_pulizia['Data Pulizia'] == giorno]['Tempo (int)'].sum()
                total = t_prep + t_puli
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
                        return 'background-color: #c6e2b3'
                    elif norm <= 0.4:
                        return 'background-color: #fff8b0'
                    elif norm <= 0.6:
                        return 'background-color: #ffe080'
                    elif norm <= 0.8:
                        return 'background-color: #ffb347'
                    else:
                        return 'background-color: #e57373'
                return ''

            styled = df_riepilogo.style.applymap(background_gradient, subset=pd.IndexSlice[['Totale Giorno'], :])
            styled.set_table_styles([{ 'selector': 'th.col_heading', 'props': [('font-weight', 'bold')] }])
            st.dataframe(styled, use_container_width=True)

            if messaggi_speciali:
                st.subheader("ℹ️ Note sui cambi senza unità")
                for msg in messaggi_speciali:
                    st.info(msg)

            with st.expander("📅 Esporta Logbook in Excel"):
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

