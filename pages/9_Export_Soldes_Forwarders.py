# -*- coding: utf-8 -*-
"""Soldes forwarders / transitaires — montants sur bookings + page Paiements."""
import streamlit as st
import importlib
import export_data as ed
import transitaires_data as td
importlib.reload(ed)
importlib.reload(td)

st.set_page_config(page_title="Soldes Forwarders", page_icon="💰", layout="wide")
if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

st.title("💰 Export — Soldes Forwarders / Transitaires")
st.caption("Montants forwarder saisis sur les Bookings · paiements via la page Paiements (Type EXPORT)")

sit = ed.situation_forwarders()
trs = td.charger_transitaires(actifs_seulement=False)

if sit.empty:
    st.info("Aucun booking avec forwarder.")
else:
    k1, k2, k3 = st.columns(3)
    k1.metric("Forwarders", len(sit))
    k2.metric("Facturé", ed.fmt(sit["Montant_Facture"].sum(), "FCFA"))
    k3.metric("Solde", ed.fmt(sit["Solde"].sum(), "FCFA"))
    st.dataframe(sit, hide_index=True, width="stretch", column_config={
        "Montant_Facture": st.column_config.NumberColumn("Facturé", format="%d"),
        "Paye": st.column_config.NumberColumn("Payé", format="%d"),
        "Solde": st.column_config.NumberColumn("Solde", format="%d"),
        "Ctr_Booked": "Booked",
        "Ctr_Loaded": "Loaded",
    })
    st.download_button(
        "⬇️ Excel soldes",
        data=ed.to_excel_bytes({"Soldes_Forwarders": sit, "Bookings": ed.charger_bookings()}),
        file_name="soldes_forwarders.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

st.info(
    "Pour payer un forwarder : page **Paiements** → Type **EXPORT** → Prestataire = nom du forwarder "
    "→ renseigner aussi **Forwarder Amount / Paid** sur le booking concerné."
)
st.caption("NAP SARL — Soldes forwarders")
