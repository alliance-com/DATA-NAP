# -*- coding: utf-8 -*-
"""Export — Stuffing & Conteneurs"""
import streamlit as st
import pandas as pd
from datetime import date

import sys
import os
_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
from export import export_data as ex


st.set_page_config(page_title="Export Stuffing", page_icon="📦", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

st.title("📦 Export — Stuffing & Conteneurs")

try:
    df_b = ex.charger_bookings()
    df_c = ex.charger_containers()
except FileNotFoundError as e:
    st.error(str(e))
    st.stop()

pending = ex.bookings_pending_stuffing(df_b)
k1, k2, k3 = st.columns(3)
k1.metric("Bookings en attente d'empotage", len(pending))
k2.metric("Conteneurs saisis", len(df_c))
k3.metric(
    "Ctr restants (balance)",
    int(pending["Balance"].sum()) if not pending.empty else 0,
)

st.markdown("---")
st.subheader("➕ Enregistrer un conteneur empoté")

opts = [""]
if not df_b.empty:
    opts += [
        f"{r.Booking_No} | {r.Forwarder} | {r.Stuffing_Status} | bal={r.Balance}"
        for _, r in df_b.iterrows()
    ]

with st.form("form_ctr"):
    choix = st.selectbox("Booking *", opts)
    c1, c2, c3 = st.columns(3)
    with c1:
        cno = st.text_input("N° Conteneur *")
        ctype = st.selectbox("Type", ["20'", "40'", "40'HC", "Flexi", "Autre"])
        seal = st.text_input("Seal")
    with c2:
        weight = st.number_input("Poids (KG)", min_value=0.0, step=100.0)
        tank = st.selectbox("Tank", [""] + ex.liste_tanks())
        d_st = st.date_input("Date stuffing", value=date.today())
    with c3:
        status = st.selectbox("Status load", ["LOADED", "PENDING", "REJECTED"])
        notes = st.text_input("Notes")
    ok = st.form_submit_button("💾 Enregistrer conteneur", type="primary")

if ok:
    if not choix or not cno.strip():
        st.error("Booking et n° conteneur obligatoires.")
    else:
        bno = choix.split("|")[0].strip()
        bl = ""
        row = df_b[df_b["Booking_No"] == bno]
        if not row.empty:
            bl = row.iloc[0]["BL_No"]
        success, msg = ex.ajouter_container({
            "Container_No": cno.strip(),
            "Booking_No": bno,
            "BL_No": bl,
            "Type_Ctr": ctype,
            "Seal": seal,
            "Weight_KG": weight,
            "Status_Load": status,
            "Tank": tank,
            "Date_Stuffing": d_st,
            "Notes": notes,
        })
        if success:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

st.markdown("---")
st.subheader("Bookings non complétés")
if pending.empty:
    st.success("Tous les bookings sont empotés (COMPLETED).")
else:
    aff = pending.copy()
    aff["ETD"] = pd.to_datetime(aff["ETD"], errors="coerce").dt.strftime("%d/%m/%Y")
    st.dataframe(
        aff[[
            "Booking_No", "BL_No", "Commodity", "Buyer", "Forwarder",
            "Ctr_Booked", "Ctr_Loaded", "Balance", "Storage_Tank",
            "Stuffing_Status", "Vessel", "ETD",
        ]],
        hide_index=True, width="stretch",
    )

st.markdown("---")
st.subheader("Conteneurs enregistrés")
if df_c.empty:
    st.info("Aucun conteneur saisi.")
else:
    affc = df_c.copy()
    affc["Date_Stuffing"] = pd.to_datetime(affc["Date_Stuffing"], errors="coerce").dt.strftime("%d/%m/%Y")
    st.dataframe(affc, hide_index=True, width="stretch", height=360)

st.caption("NAP SARL — Stuffing · export_data.xlsx")
