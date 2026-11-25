# -*- coding: utf-8 -*-
"""Export — Dashboard KPI & alertes"""
import streamlit as st
import pandas as pd
from datetime import date

import sys
import os
_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
from export import export_data as ex


st.set_page_config(page_title="Export Dashboard", page_icon="🚢", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

st.title("🚢 Export — Dashboard")
st.caption("Pilotage bookings · stuffing · shipping · alertes opérationnelles")

try:
    df = ex.charger_bookings()
except FileNotFoundError as e:
    st.error(str(e))
    st.info("Placez **export_data.xlsx** à la racine du projet, puis rechargez.")
    st.stop()

# Filtres
c1, c2, c3, c4 = st.columns(4)
with c1:
    commodities = ["Tous"] + sorted(df["Commodity"].dropna().unique().tolist()) if not df.empty else ["Tous"]
    f_comm = st.selectbox("Commodity", commodities)
with c2:
    buyers = ["Tous"] + sorted(df["Buyer"].dropna().unique().tolist()) if not df.empty else ["Tous"]
    f_buyer = st.selectbox("Buyer", buyers)
with c3:
    fwds = ["Tous"] + sorted(df["Forwarder"].dropna().unique().tolist()) if not df.empty else ["Tous"]
    f_fwd = st.selectbox("Forwarder", fwds)
with c4:
    f_stuff = st.selectbox("Stuffing", ["Tous"] + ex.STUFFING_STATUSES)

vue = df.copy()
if not vue.empty:
    if f_comm != "Tous":
        vue = vue[vue["Commodity"] == f_comm]
    if f_buyer != "Tous":
        vue = vue[vue["Buyer"] == f_buyer]
    if f_fwd != "Tous":
        vue = vue[vue["Forwarder"] == f_fwd]
    if f_stuff != "Tous":
        vue = vue[vue["Stuffing_Status"] == f_stuff]

ind = ex.indicateurs(vue)

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Bookings", ind["bookings"])
k2.metric("Ctr booked / loaded", f"{ind['ctr_booked']} / {ind['ctr_loaded']}")
k3.metric("% Load", f"{ind['pct_load']*100:.1f} %")
k4.metric("Qty nette (MT)", f"{ind['qty_mt']:,.2f}".replace(",", " "))
k5.metric("À empoter", ind["pending_stuff"])

if ind["ctr_booked"] > 0:
    st.progress(min(ind["pct_load"], 1.0), text=f"{ind['pct_load']*100:.0f} % conteneurs empotés")

st.markdown("---")

# Alertes
st.subheader("🚨 Alertes")
alerts = ex.alertes(vue)
if not alerts:
    st.success("Aucune alerte sur le filtre courant.")
else:
    crit = [a for a in alerts if a["Niveau"] == "CRITIQUE"]
    warn = [a for a in alerts if a["Niveau"] != "CRITIQUE"]
    if crit:
        st.error(f"**{len(crit)} alerte(s) critique(s)**")
        st.dataframe(pd.DataFrame(crit), hide_index=True, width="stretch")
    if warn:
        st.warning(f"**{len(warn)} avertissement(s)**")
        st.dataframe(pd.DataFrame(warn), hide_index=True, width="stretch")

st.markdown("---")
st.subheader("📋 Bookings — vue opérationnelle")
if vue.empty:
    st.info("Aucun booking. Saisissez des données dans la page **Export Bookings**.")
else:
    aff = vue.copy()
    for c in ["ETD", "ETA", "Date_Creation"]:
        if c in aff.columns:
            aff[c] = pd.to_datetime(aff[c], errors="coerce").dt.strftime("%d/%m/%Y")
    cols = [
        "Booking_No", "BL_No", "Commodity", "Buyer", "Forwarder", "Liner",
        "Ctr_Booked", "Ctr_Loaded", "Balance", "Net_Qty_KG", "Storage_Tank",
        "Stuffing_Status", "Vessel", "ETD", "ETA", "Shipping_Status",
    ]
    cols = [c for c in cols if c in aff.columns]
    st.dataframe(aff[cols], hide_index=True, width="stretch", height=420)

    st.download_button(
        "⬇️ Export Excel (filtre courant)",
        data=ex.to_excel_bytes(vue),
        file_name=f"Export_Bookings_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# Répartition forwarder
if not vue.empty:
    st.markdown("---")
    st.subheader("📦 Stuffing par transitaire")
    g = (
        vue.groupby("Forwarder", dropna=False)
        .agg(
            Bookings=("Booking_No", "count"),
            Completed=("Stuffing_Status", lambda s: (s == "COMPLETED").sum()),
            Pending=("Stuffing_Status", lambda s: (s != "COMPLETED").sum()),
            Ctr_Booked=("Ctr_Booked", "sum"),
            Ctr_Loaded=("Ctr_Loaded", "sum"),
        )
        .reset_index()
    )
    st.dataframe(g, hide_index=True, width="stretch")

st.caption("NAP SARL — Export Master · source export_data.xlsx")
