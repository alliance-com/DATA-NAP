# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import presence_data as db

st.set_page_config(page_title="Ma présence — Exposé N° 3", page_icon="📌", layout="wide")

user = st.session_state.get("user")
if not user:
    st.warning("Veuillez vous connecter.")
    st.page_link("pages/1_Connexion.py", label="🔑 Connexion")
    st.stop()

st.title("📌 Ma présence")
st.caption(f"{user.get('Prenom')} {user.get('Nom')} · {user.get('Profession', '')}")

d_dim = db.dimanche_courant_ou_dernier()
d_proch = db.prochain_dimanche()
opts = [d_dim]
if d_proch != d_dim:
    opts.append(d_proch)

choix = st.radio(
    "Dimanche",
    opts,
    format_func=lambda d: d.strftime("%A %d/%m/%Y").capitalize(),
)
notes = st.text_input("Note (optionnel)")
if st.button("✋ Signaler ma présence", type="primary"):
    ok, msg = db.signaler_presence(int(user["ID"]), choix, notes=notes)
    if ok:
        st.success(msg)
    else:
        st.warning(msg)

st.markdown("#### Historique")
mes = db.charger_presences(id_membre=int(user["ID"]))
if mes.empty:
    st.caption("Aucune présence signalée.")
else:
    aff = mes.head(20).copy()
    aff["Date_Dimanche"] = pd.to_datetime(aff["Date_Dimanche"], errors="coerce").dt.strftime("%d/%m/%Y")
    st.dataframe(
        aff[["Date_Dimanche", "Statut", "Signale_Le", "Valide_Le", "Valide_Par"]],
        hide_index=True,
        width="stretch",
    )
