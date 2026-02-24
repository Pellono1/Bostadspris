"""
app.py – Streamlit-app for nordisk boligprisstatistikk
=======================================================
Norge: SSB tabell 06035 (region, årsvis) + 07241 (nasjonalt, kvartal)
Sverige: SCB BO0501A2 (bostadsrätter per län, kvartal)

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
URL_SCB     = "https://api.scb.se/OV0104/v1/doris/sv/ssd/BO/BO0501/BO0501A/FastprisHelAr"
ONSKEDE     = ["Oslo", "Bergen", "Trondheim", "Stavanger", "Hele landet", "The whole country"]

st.set_page_config(page_title="Nordisk Boligprisstatistikk", page_icon="🏠", layout="wide")

# ── HJELPEFUNKSJONAR ──────────────────────────────────────────────────────────
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

# ── NORGE: HENT DATA ──────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def hent_no_region():
    meta = hent_meta(URL_REGION)
    reg_k = [k for k, l in zip(meta["Region"]["values"], meta["Region"]["labels"])
             if any(l.lower().startswith(o.lower()) or o.lower() in l.lower() for o in ONSKEDE)]
    if not reg_k:
        reg_k = meta["Region"]["values"][:5]
    rows = parse(hent_json(URL_REGION, [
        {"code": "Region",       "selection": {"filter": "item", "values": reg_k}},
        {"code": "Boligtype",    "selection": {"filter": "item", "values": meta["Boligtype"]["values"][:3]}},
        {"code": "ContentsCode", "selection": {"filter": "item", "values": [meta["ContentsCode"]["values"][0]]}},
        {"code": "Tid",          "selection": {"filter": "item", "values": meta["Tid"]["values"][-8:]}},
    ]))
    df = pd.DataFrame(rows).rename(columns={"Region":"region","Boligtype":"boligtype","Tid":"periode","verdi":"pris_m2"})
    df["region"] = df["region"].str.split(" - ").str[0]
    df["land"] = "Norge"
    return df

@st.cache_data(ttl=3600)
def hent_no_kvartal():
    meta = hent_meta(URL_KVARTAL)
    rows = parse(hent_json(URL_KVARTAL, [
        {"code": "Boligtype",    "selection": {"filter": "item", "values": meta["Boligtype"]["values"][:3]}},
        {"code": "ContentsCode", "selection": {"filter": "item", "values": [meta["ContentsCode"]["values"][0]]}},
        {"code": "Tid",          "selection": {"filter": "item", "values": meta["Tid"]["values"][-12:]}},
    ]))
    df = pd.DataFrame(rows).rename(columns={"Boligtype":"boligtype","Tid":"periode","verdi":"pris_m2"})
    df["land"] = "Norge"
    return df, fix(meta["ContentsCode"]["labels"][0])

# ── SVERIGE: HENT DATA ────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def hent_se():
    try:
        meta = hent_meta(URL_SCB)
        # Hent siste 8 perioder
        perioder = meta["Tid"]["values"][-8:]
        # Hent alle regioner (län)
        regioner = meta["Region"]["values"]
        # Hent første innholdsvariabel (pris)
        innhold  = meta["ContentsCode"]["values"][:1]

        rows = parse(hent_json(URL_SCB, [
            {"code": "Region",       "selection": {"filter": "item", "values": regioner}},
            {"code": "ContentsCode", "selection": {"filter": "item", "values": innhold}},
            {"code": "Tid",          "selection": {"filter": "item", "values": perioder}},
        ]))
        df = pd.DataFrame(rows).rename(columns={"Region":"region","Tid":"periode","verdi":"pris_m2"})
        df["land"] = "Sverige"
        df["boligtype"] = "Alla"
        innhold_namn = meta["ContentsCode"]["labels"][0]
        return df, innhold_namn
    except Exception as e:
        return pd.DataFrame(), f"Fel: {e}"

# ── EXCEL-EKSPORT ─────────────────────────────────────────────────────────────
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
st.title("🏠 Nordisk Boligprisstatistikk")

col_tittel, col_knapp = st.columns([4, 1])
with col_knapp:
    if st.button("🔄 Hent ny data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
with col_tittel:
    st.caption("Norge: SSB (gratis) · Sverige: SCB (gratis)")

tab1, tab2, tab3 = st.tabs([
    "🇳🇴 Norge – region (årsvis)",
    "🇳🇴 Norge – nasjonalt (kvartal)",
    "🇸🇪 Sverige – län (årsvis)"
])

# ── TAB 1: NORGE REGION ───────────────────────────────────────────────────────
with tab1:
    try:
        df_no_r = hent_no_region()
        col1, col2, col3 = st.columns(3)
        with col1:
            reg = sorted(df_no_r["region"].unique())
            valgte_reg = st.multiselect("Region", reg, default=reg)
        with col2:
            bt = sorted(df_no_r["boligtype"].unique())
            valgte_bt = st.multiselect("Boligtype", bt, default=bt)
        with col3:
            per = sorted(df_no_r["periode"].unique())
            valgt_per = st.selectbox("År", ["Alle"] + list(reversed(per)))

        f = df_no_r[df_no_r["region"].isin(valgte_reg) & df_no_r["boligtype"].isin(valgte_bt)]
        if valgt_per != "Alle":
            f = f[f["periode"] == valgt_per]
        f = f[["region","boligtype","periode","pris_m2"]].sort_values(["periode","region"], ascending=[False,True])
        f = f.rename(columns={"region":"Region","boligtype":"Boligtype","periode":"År","pris_m2":"Pris per m² (NOK)"})

        st.markdown(f"**{len(f)} rader**")
        st.dataframe(f.style.format({"Pris per m² (NOK)": "{:,.0f}"}), use_container_width=True, hide_index=True)
        st.download_button("⬇️ Last ned Excel", data=lag_excel(f, "NO Region"),
                           file_name="no_region.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if valgt_per == "Alle":
            pivot = f.pivot_table(index="År", columns="Region", values="Pris per m² (NOK)", aggfunc="mean")
            st.line_chart(pivot)
    except Exception as e:
        st.error(f"Kunne ikke hente data: {e}")

# ── TAB 2: NORGE KVARTAL ──────────────────────────────────────────────────────
with tab2:
    try:
        df_no_k, innhold_no = hent_no_kvartal()
        col1, col2 = st.columns(2)
        with col1:
            bt_k = sorted(df_no_k["boligtype"].unique())
            valgte_bt_k = st.multiselect("Boligtype", bt_k, default=bt_k, key="bt_k")
        with col2:
            antall = st.slider("Antall kvartaler", 4, 20, 12)

        per_k = sorted(df_no_k["periode"].unique())[-antall:]
        fk = df_no_k[df_no_k["boligtype"].isin(valgte_bt_k) & df_no_k["periode"].isin(per_k)]
        fk = fk[["boligtype","periode","pris_m2"]].sort_values(["periode","boligtype"], ascending=[False,True])
        fk = fk.rename(columns={"boligtype":"Boligtype","periode":"Kvartal","pris_m2":innhold_no})

        st.markdown(f"**{len(fk)} rader**")
        st.dataframe(fk.style.format({innhold_no: "{:,.0f}"}), use_container_width=True, hide_index=True)
        st.download_button("⬇️ Last ned Excel", data=lag_excel(fk, "NO Kvartal"),
                           file_name="no_kvartal.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_k")
        pivot_k = fk.pivot_table(index="Kvartal", columns="Boligtype", values=innhold_no, aggfunc="mean")
        st.line_chart(pivot_k)
    except Exception as e:
        st.error(f"Kunne ikke hente data: {e}")

# ── TAB 3: SVERIGE ────────────────────────────────────────────────────────────
with tab3:
    try:
        df_se, innhold_se = hent_se()
        if df_se.empty:
            st.error(f"Kunde inte hämta SCB-data: {innhold_se}")
        else:
            col1, col2 = st.columns(2)
            with col1:
                lan = sorted(df_se["region"].unique())
                valgte_lan = st.multiselect("Län", lan, default=lan[:8])
            with col2:
                per_se = sorted(df_se["periode"].unique())
                valgt_per_se = st.selectbox("År", ["Alla"] + list(reversed(per_se)))

            fs = df_se[df_se["region"].isin(valgte_lan)]
            if valgt_per_se != "Alla":
                fs = fs[fs["periode"] == valgt_per_se]
            fs = fs[["region","periode","pris_m2"]].sort_values(["periode","region"], ascending=[False,True])
            fs = fs.rename(columns={"region":"Län","periode":"År","pris_m2":innhold_se})

            st.markdown(f"**{len(fs)} rader**")
            st.dataframe(fs.style.format({innhold_se: "{:,.0f}"}), use_container_width=True, hide_index=True)
            st.download_button("⬇️ Ladda ned Excel", data=lag_excel(fs, "SE Län"),
                               file_name="se_lan.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_se")
            if valgt_per_se == "Alla":
                pivot_se = fs.pivot_table(index="År", columns="Län", values=innhold_se, aggfunc="mean")
                st.line_chart(pivot_se)
    except Exception as e:
        st.error(f"Kunde inte hämta SCB-data: {e}")

st.divider()
st.caption(f"Uppdaterad: {date.today().strftime('%d.%m.%Y')} · Klikk 'Hent ny data' for å oppdatere")