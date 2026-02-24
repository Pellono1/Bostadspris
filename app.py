"""
app.py – Streamlit-app for norsk boligprisstatistikk
======================================================
Kjør:
    pip install streamlit openpyxl
    streamlit run app.py
"""

import sqlite3
import io
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

DB_FIL = "boligpriser.db"

st.set_page_config(
    page_title="Norsk Boligprisstatistikk",
    page_icon="🏠",
    layout="wide"
)

# ── HENT DATA ─────────────────────────────────────────────────────────────────
@st.cache_data
def hent_region():
    conn = sqlite3.connect(DB_FIL)
    df = pd.read_sql("SELECT * FROM region_arsvis", conn)
    conn.close()
    return df

@st.cache_data
def hent_kvartal():
    conn = sqlite3.connect(DB_FIL)
    df = pd.read_sql("SELECT * FROM nasjonalt_kvartal", conn)
    conn.close()
    return df

# ── EXCEL-EKSPORT ─────────────────────────────────────────────────────────────
def lag_excel(df, arknavn):
    wb = Workbook()
    ws = wb.active
    ws.title = arknavn
    hdr_font = Font(name="Arial", bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", start_color="1F4E79", fgColor="1F4E79")
    for col, h in enumerate(df.columns, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center")
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col, val in enumerate(row, 1):
            ws.cell(row=row_idx, column=col, value=val)
    for i in range(1, len(df.columns)+1):
        ws.column_dimensions[get_column_letter(i)].width = 20
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ── LAYOUT ────────────────────────────────────────────────────────────────────
st.title("🏠 Norsk Boligprisstatistikk")
st.caption("Datakilde: SSB (Statistics Norway) · Gratis og åpen data")

try:
    df_region  = hent_region()
    df_kvartal = hent_kvartal()
except Exception as e:
    st.error(f"Kunne ikke lese databasen: {e}")
    st.info("Kjør norway_prototype.py først for å lage databasen.")
    st.stop()

tab1, tab2 = st.tabs(["📍 Pris per region (årsvis)", "📈 Nasjonal prisutvikling (kvartal)"])

# ── TAB 1: REGION ─────────────────────────────────────────────────────────────
with tab1:
    col1, col2, col3 = st.columns(3)

    with col1:
        regioner = sorted(df_region["region"].unique())
        valgte_regioner = st.multiselect("Region", regioner, default=regioner)

    with col2:
        boligtyper = sorted(df_region["boligtype"].unique())
        valgte_bt = st.multiselect("Boligtype", boligtyper, default=boligtyper)

    with col3:
        perioder = sorted(df_region["periode"].unique())
        valgt_periode = st.selectbox("År", ["Alle"] + list(reversed(perioder)))

    filtrert = df_region[
        df_region["region"].isin(valgte_regioner) &
        df_region["boligtype"].isin(valgte_bt)
    ]
    if valgt_periode != "Alle":
        filtrert = filtrert[filtrert["periode"] == valgt_periode]

    filtrert = filtrert[["region","boligtype","periode","pris_m2"]].sort_values(
        ["periode","region"], ascending=[False, True]
    ).rename(columns={"region":"Region","boligtype":"Boligtype",
                       "periode":"År","pris_m2":"Pris per m² (NOK)"})

    st.markdown(f"**{len(filtrert)} rader** vises")
    st.dataframe(
        filtrert.style.format({"Pris per m² (NOK)": "{:,.0f}"}),
        use_container_width=True, hide_index=True
    )

    excel = lag_excel(filtrert, "Region årsvis")
    st.download_button(
        "⬇️ Last ned som Excel",
        data=excel,
        file_name=f"boligpris_region_{valgt_periode}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Enkel graf
    if valgt_periode == "Alle" and valgte_regioner and valgte_bt:
        st.subheader("Prisutvikling per region")
        pivot = filtrert.pivot_table(
            index="År", columns="Region",
            values="Pris per m² (NOK)", aggfunc="mean"
        )
        st.line_chart(pivot)

# ── TAB 2: KVARTAL ────────────────────────────────────────────────────────────
with tab2:
    col1, col2 = st.columns(2)

    with col1:
        bt_kv = sorted(df_kvartal["boligtype"].unique())
        valgte_bt_kv = st.multiselect("Boligtype", bt_kv, default=bt_kv, key="bt_kv")

    with col2:
        antall = st.slider("Antall kvartaler bakover", 4, 20, 12)

    perioder_kv = sorted(df_kvartal["periode"].unique())[-antall:]
    filtrert_kv = df_kvartal[
        df_kvartal["boligtype"].isin(valgte_bt_kv) &
        df_kvartal["periode"].isin(perioder_kv)
    ][["boligtype","periode","pris_m2"]].sort_values(
        ["periode","boligtype"], ascending=[False, True]
    ).rename(columns={"boligtype":"Boligtype","periode":"Kvartal","pris_m2":"Pris per m² (NOK)"})

    st.markdown(f"**{len(filtrert_kv)} rader** vises")
    st.dataframe(
        filtrert_kv.style.format({"Pris per m² (NOK)": "{:,.0f}"}),
        use_container_width=True, hide_index=True
    )

    excel_kv = lag_excel(filtrert_kv, "Nasjonalt kvartal")
    st.download_button(
        "⬇️ Last ned som Excel",
        data=excel_kv,
        file_name="boligpris_kvartal.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_kv"
    )

    # Graf
    if valgte_bt_kv:
        st.subheader("Prisutvikling nasjonalt")
        pivot_kv = filtrert_kv.pivot_table(
            index="Kvartal", columns="Boligtype",
            values="Pris per m² (NOK)", aggfunc="mean"
        )
        st.line_chart(pivot_kv)

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("Prototype · Data fra SSB tabell 06035 og 07241 · Oppdater data ved å kjøre norway_prototype.py")