# -*- coding: utf-8 -*-
"""
NAP SARL — Conteneurs / Stuffing
1 LIGNE = 1 CONTENEUR (comme SHEA BUTTER stuffing report)
BL et Booking se répètent sur plusieurs lignes.
"""
import streamlit as st
import pandas as pd
from datetime import date
import importlib
import export_data as ed
importlib.reload(ed)

st.set_page_config(page_title="Export Conteneurs", page_icon="📦", layout="wide")
if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

st.title("📦 Export — Conteneurs / Stuffing · STE NAP SARL")
st.caption("**1 ligne = 1 conteneur** · même Booking/BL sur plusieurs lignes · comme le Stuffing Report")

bk = ed.charger_bookings()
opts_bk = {
    f"{r.Booking} · {r.Forwarder} · {r.Commodity} · booked {int(r.Container_Booked)}": str(r.Booking)
    for r in bk.itertuples()
} if not bk.empty else {}

st.sidebar.header("➕ Nouveau conteneur")
with st.sidebar.form("form_ctr", clear_on_submit=True):
    date_d = st.date_input("Date stuffing", value=date.today())
    bk_sel = st.selectbox("Booking lié", ["— Manuel —"] + list(opts_bk.keys()))
    booking_m = st.text_input("N° Booking (manuel si besoin)")
    bl_m = st.text_input("N° BL", placeholder="Souvent = Booking")
    camion = st.text_input("Camion Number")
    conteneur = st.text_input("Container Number *", placeholder="CAAU2346542")
    plomb = st.text_input("Plomb / Seal")
    nb_sacs = st.number_input("No. Of Sacs", min_value=0, value=1, step=1)
    gross = st.number_input("Gross Wt (kg)", min_value=0.0, step=10.0, format="%.0f")
    tare = st.number_input("Tare Wt (kg)", min_value=0.0, step=10.0, format="%.0f")
    net = st.number_input("Net Weight (kg)", min_value=0.0, value=max(gross - tare, 0.0), step=10.0, format="%.0f")
    storage = st.text_input("Loading Wh / Tank", placeholder="TANK 4")
    rotation = st.number_input("Rotation", min_value=0, value=1, step=1)
    buyer = st.text_input("Buyer Name")
    contract = st.text_input("Contract No")
    commodity = st.selectbox("Commodity", ed.COMMODITES)
    forwarder = st.text_input("Forwarder / Agent")
    invoice = st.text_input("Invoice No")
    statut = st.selectbox("Status", ed.STATUTS_STUFFING, index=2)
    notes = st.text_input("Notes")
    if st.form_submit_button("💾 Enregistrer conteneur", width="stretch"):
        if not conteneur.strip():
            st.error("N° conteneur obligatoire.")
        else:
            booking = opts_bk[bk_sel] if bk_sel != "— Manuel —" else booking_m.strip()
            # auto-fill from booking
            if booking and not bk.empty:
                m = bk[bk["Booking"].astype(str) == str(booking)]
                if not m.empty:
                    mr = m.iloc[0]
                    buyer = buyer or str(mr["Buyer"] or "")
                    contract = contract or str(mr["Contract_No"] or "")
                    commodity = commodity or str(mr["Commodity"] or commodity)
                    forwarder = forwarder or str(mr["Forwarder"] or "")
                    bl_m = bl_m or str(mr["BL"] or booking)
            ok, msg = ed.ajouter_conteneur({
                "Date": date_d, "Camion": camion, "Container_No": conteneur, "Plomb": plomb,
                "Nb_Sacs": nb_sacs, "Gross_KG": gross, "Tare_KG": tare, "Net_KG": net,
                "Product_KG": net, "Storage": storage, "Rotation": rotation,
                "Buyer": buyer, "Contract_No": contract, "Booking": booking,
                "BL": bl_m or booking, "Commodity": commodity, "Forwarder": forwarder,
                "Invoice_No": invoice, "Stuffing_Status": statut, "Notes": notes,
            }, name)
            st.success(msg) if ok else st.error(msg)
            if ok:
                st.rerun()

df = ed.charger_conteneurs()
if not df.empty:
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Conteneurs", len(df))
    k2.metric("Bookings couverts", int(df["Booking"].replace("", pd.NA).nunique(dropna=True)))
    k3.metric("Net total (T)", f"{df['Net_KG'].sum()/1000:,.2f}".replace(",", " "))
    k4.metric("BL distincts", int(df["BL"].replace("", pd.NA).nunique(dropna=True)))

st.subheader("📋 Liste conteneurs (1 ligne = 1 conteneur)")
if df.empty:
    st.info("Aucun conteneur. Liez-les à un Booking déjà créé pour mettre à jour le loaded automatiquement.")
else:
    f1, f2, f3, f4 = st.columns(4)
    f_bk = f1.selectbox("Booking", ["Tous"] + sorted([x for x in df["Booking"].unique() if x]))
    f_fw = f2.selectbox("Forwarder", ["Tous"] + sorted([x for x in df["Forwarder"].unique() if x]))
    f_bl = f3.selectbox("BL", ["Tous"] + sorted([x for x in df["BL"].unique() if x]))
    f_q = f4.text_input("Recherche", placeholder="Conteneur, camion, plomb…")
    v = df.copy()
    if f_bk != "Tous":
        v = v[v["Booking"] == f_bk]
    if f_fw != "Tous":
        v = v[v["Forwarder"] == f_fw]
    if f_bl != "Tous":
        v = v[v["BL"] == f_bl]
    if f_q.strip():
        q = f_q.strip().lower()
        v = v[v.apply(lambda r: q in " ".join(r.astype(str)).lower(), axis=1)]
    v = v.sort_values(["Booking", "Date", "ID"], ascending=[False, False, False])
    st.caption(f"{len(v)} ligne(s) — BL/Booking peuvent se répéter")
    st.dataframe(
        v[["ID", "Date", "Camion", "Container_No", "Plomb", "Net_KG", "Gross_KG", "Tare_KG",
           "Storage", "Booking", "BL", "Contract_No", "Buyer", "Forwarder", "Commodity",
           "Invoice_No", "Stuffing_Status"]],
        hide_index=True, width="stretch", height=450,
        column_config={
            "Date": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
            "Container_No": "Container",
            "Net_KG": st.column_config.NumberColumn("Net", format="%d"),
            "Gross_KG": st.column_config.NumberColumn("Gross", format="%d"),
            "Tare_KG": st.column_config.NumberColumn("Tare", format="%d"),
        },
    )
    st.download_button(
        "⬇️ Excel Conteneurs",
        data=ed.to_excel_bytes({"Conteneurs": v}),
        file_name="conteneurs_stuffing.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

st.markdown("---")
st.subheader("✏️ Modifier un conteneur")
if not df.empty:
    opts = {
        f"#{int(r.ID)} · {r.Container_No} · BK:{r.Booking or '—'} · BL:{r.BL or '—'}": int(r.ID)
        for r in df.itertuples()
    }
    choix = st.selectbox("Ligne", [""] + list(opts.keys()), format_func=lambda x: "— Choisir —" if x == "" else x)
    if choix:
        rid = opts[choix]
        row = df[df["ID"] == rid].iloc[0]
        with st.form("mod_ctr"):
            a, b, c = st.columns(3)
            with a:
                e_date = st.date_input("Date", value=row["Date"].date() if pd.notna(row["Date"]) else date.today())
                e_ctr = st.text_input("Container", value=str(row["Container_No"] or ""))
                e_camion = st.text_input("Camion", value=str(row["Camion"] or ""))
                e_plomb = st.text_input("Plomb", value=str(row["Plomb"] or ""))
            with b:
                e_bk = st.text_input("Booking", value=str(row["Booking"] or ""))
                e_bl = st.text_input("BL", value=str(row["BL"] or ""))
                e_net = st.number_input("Net kg", min_value=0.0, value=float(row["Net_KG"] or 0), step=10.0)
                e_storage = st.text_input("Storage", value=str(row["Storage"] or ""))
            with c:
                e_fwd = st.text_input("Forwarder", value=str(row["Forwarder"] or ""))
                e_inv = st.text_input("Invoice", value=str(row["Invoice_No"] or ""))
                e_st = st.selectbox("Status", ed.STATUTS_STUFFING,
                                   index=ed.STATUTS_STUFFING.index(row["Stuffing_Status"]) if row["Stuffing_Status"] in ed.STATUTS_STUFFING else 0)
            if st.form_submit_button("💾 Enregistrer", type="primary"):
                ok, msg = ed.modifier_conteneur(rid, {
                    "Date": e_date, "Container_No": e_ctr, "Camion": e_camion, "Plomb": e_plomb,
                    "Booking": e_bk, "BL": e_bl, "Net_KG": e_net, "Product_KG": e_net,
                    "Storage": e_storage, "Forwarder": e_fwd, "Invoice_No": e_inv, "Stuffing_Status": e_st,
                }, name)
                st.success(msg) if ok else st.error(msg)
                if ok:
                    st.rerun()

if not df.empty:
    with st.expander("🗑️ Supprimer un conteneur"):
        opts_s = {f"#{int(r.ID)} · {r.Container_No}": int(r.ID) for r in df.itertuples()}
        cs = st.selectbox("À supprimer", list(opts_s.keys()), key="supc")
        if st.checkbox("Confirmer", key="confc") and st.button("Supprimer", type="primary", key="btnc"):
            ok, msg = ed.supprimer_conteneur(opts_s[cs])
            st.success(msg) if ok else st.error(msg)
            st.rerun()

st.caption("NAP SARL — Conteneurs = Stuffing Report · 1 ligne = 1 conteneur")
