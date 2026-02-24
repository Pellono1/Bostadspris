"""
app.py – Streamlit-app for norsk boligprisstatistikk
======================================================
Kjør:
    pip install streamlit openpyxl pandas requests
    streamlit run app.py
"""

import sqlite3
import io
import requests
import streamlit as st
import pandas as pd
from itertools import product
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

DB_FIL      = "boligpriser.db"
URL_REGION  = "https://data.ssb.no/api/v0/en/table/06035"
URL_KVARTAL = "https://data.ssb.no/api/v0/en/table/07241"
ONSKEDE     = ["Oslo", "Bergen", "Trondheim", "Stavanger", "Hele landet", "The whole country"]

st.set_page_config(page_title="Norsk Boligprisstatistikk", page_icon="🏠", layout="wide")

# ── DATAFUNKSJONAR ────────────────────────────────────────────────────────────
def fix(t):
    return t.replace("km²","m²").replace("km2","m²").replace("KM2","m²")

def hent_meta(url):
    r = requests.get(url, timeout=30); r.raise_for_status()
    return {v["code"]: {"values": v["values"], "labels": v["valueTexts"]}
            for v in r.json()["variables"]}

def hent_json(url, query):
    r = requests.post(url, json={"query": query, "response": {"format": "json-stat2"}}, timeout=30)
    r.raise_for_status(); return r.json()

def parse(data):
    dims = data["dimension"]; vals = data["value"]
    keys = list(dims.keys())
    items = [list(dims[k]["category"]["label"].items()) for k in keys]
    rows, idx = [], 0
    for combo in product(*items):
        v = vals[idx] if idx < len(vals) else None
        idx += 1
        if v is not None:
            row = {keys[i]: combo[i][1] for i in range(len(keys))}
            row["verdi"] = v
            rows.append(row)
    return rows

def oppdater_db():
    meta_r = hent_meta(URL_REGION)
    reg_k = [k for k, l in zip(meta_r["Region"]["values"], meta_r["Region"]["labels"])
             if any(l.lower().startswith(o.lower()) or o.lower() in l.lower() for o in ONSKEDE)]
    if not reg_k:
        reg_k = meta_r["Region"]["values"][:5]

    reg_rows = parse(hent_json(URL_REGION, [
        {"code": "Region",       "selection": {"filter": "item", "values": reg_k}},
        {"code": "Boligtype",    "selection": {"filter": "item", "values": meta_r["Boligtype"]["values"][:3]}},
        {"code": "ContentsCode", "selection": {"filter": "item", "values": [meta_r["ContentsCode"]["values"][0]]}},
        {"code": "Tid",          "selection": {"filter": "item", "values": meta_r["Tid"]["values"][-8:]}},
    ]))

    meta_k = hent_meta(URL_KVARTAL)
    kv_rows = parse(hent_json(URL_KVARTAL, [
        {"code": "Boligtype",    "selection": {"filter": "item", "values": meta_k["Boligtype"]["values"][:3]}},
        {"code": "ContentsCode", "selection": {"filter": "item", "values": [meta_k["ContentsCode"]["values"][0]]}},
        {"code": "Tid",          "selection": {"filter": "item", "values": meta_k["Tid"]["values"][-12:]}},
    ]))

    conn = sqlite3.connect(DB_FIL)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS region_arsvis (
        id INTEGER PRIMARY KEY AUTOINCREMENT, region TEXT, boligtype TEXT,
        periode TEXT, pris_m2 REAL, land TEXT DEFAULT 'Norge',
        kilde TEXT DEFAULT 'SSB 06035', hentet TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS nasjonalt_kvartal (
        id INTEGER PRIMARY KEY AUTOINCREMENT, boligtype TEXT, periode TEXT,
        pris_m2 REAL, land TEXT DEFAULT 'Norge',
        kilde TEXT DEFAULT 'SSB 07241', hentet TEXT)""")
    c.execute("DELETE FROM region_arsvis")
    c.execute("DELETE FROM nasjonalt_kvartal")
    hentet = date.today().isoformat()
    for row in reg_rows:
        c.execute("INSERT INTO region_arsvis (region, boligtype, periode, pris_m2, hentet) VALUES (?,?,?,?,?)",
                  (row.get("Region"), row.get("Boligtype"), row.get("Tid"), row["verdi"], hentet))
    for row in kv_rows:
        c.execute("INSERT INTO nasjonalt_kvartal (boligtype, periode, pris_m2, hentet) VALUES (?,?,?,?)",
                  (row.get("Boligtype"), row.get("Tid"), row["verdi"], hentet))
    conn.commit(); conn.close()
    return len(reg_rows), len(kv_rows)

@st.cache_data
def hent_region():
    conn = sqlite3.connect(DB_FIL)
    df = pd.read_sql("SELECT * FROM region_arsvis", conn)
    conn.close()
    df["region"] = df["region"].str.split(" - ").str[0]
    return df

@st.cache_data
def hent_kvartal():
    conn = sqlite3.connect(DB_FIL)
    df = pd.read_sql("SELECT * FROM nasjonalt_kvartal", conn)
    conn.close(); return df

def lag_excel(df, arknavn):
    wb = Workbook(); ws = wb.active; ws.title = arknavn
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
    buffer = io.BytesIO(); wb.save(buffer); buffer.seek(0)
    return buffer

# ── LAYOUT ────────────────────────────────────────────────────────────────────
st.title("🏠 Norsk Boligprisstatistikk")

# Refresh-knapp i toppen
col_tittel, col_knapp = st.columns([4, 1])
with col_knapp:
    if st.button("🔄 Hent ny data", use_container_width=True):
        with st.spinner("Henter data fra SSB..."):
            try:
                r, k = oppdater_db()
                st.cache_data.clear()
                st.success(f"Oppdatert! {r} regionsrader, {k} kvartalsrader")
            except Exception as e:
                st.error(f"Feil: {e}")

with col_tittel:
    st.caption("Datakilde: SSB (Statistics Norway) · Gratis og åpen data")

try:
    df_region  = hent_region()
    df_kvartal = hent_kvartal()
except Exception:
    st.warning("Ingen data funnet. Klikk 'Hent ny data' for å laste inn data.")
    st.stop()

tab1, tab2 = st.tabs(["📍 Pris per region (årsvis)", "📈 Nasjonal prisutvikling (kvartal)"])

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
        ["periode","region"], ascending=[False,True]
    ).rename(columns={"region":"Region","boligtype":"Boligtype",
                       "periode":"År","pris_m2":"Pris per m² (NOK)"})

    st.markdown(f"**{len(filtrert)} rader** vises")
    st.dataframe(filtrert.style.format({"Pris per m² (NOK)": "{:,.0f}"}),
                 use_container_width=True, hide_index=True)
    st.download_button("⬇️ Last ned som Excel", data=lag_excel(filtrert, "Region årsvis"),
                       file_name=f"boligpris_region_{valgt_periode}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if valgt_periode == "Alle" and valgte_regioner and valgte_bt:
        st.subheader("Prisutvikling per region")
        pivot = filtrert.pivot_table(index="År", columns="Region",
                                     values="Pris per m² (NOK)", aggfunc="mean")
        st.line_chart(pivot)

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
        ["periode","boligtype"], ascending=[False,True]
    ).rename(columns={"boligtype":"Boligtype","periode":"Kvartal","pris_m2":"Pris per m² (NOK)"})

    st.markdown(f"**{len(filtrert_kv)} rader** vises")
    st.dataframe(filtrert_kv.style.format({"Pris per m² (NOK)": "{:,.0f}"}),
                 use_container_width=True, hide_index=True)
    st.download_button("⬇️ Last ned som Excel", data=lag_excel(filtrert_kv, "Nasjonalt kvartal"),
                       file_name="boligpris_kvartal.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="dl_kv")
    if valgte_bt_kv:
        st.subheader("Prisutvikling nasjonalt")
        pivot_kv = filtrert_kv.pivot_table(index="Kvartal", columns="Boligtype",
                                            values="Pris per m² (NOK)", aggfunc="mean")
        st.line_chart(pivot_kv)

st.divider()
st.caption("Prototype · SSB tabell 06035 og 07241 · Klikk 'Hent ny data' for å oppdatere")