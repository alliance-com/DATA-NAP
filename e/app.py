# -*- coding: utf-8 -*-
"""
Exposé N° 3 — URHC Paroisse Zogbadjé
Point d'entrée : streamlit run app.py
"""
import streamlit as st

st.set_page_config(
    page_title="Exposé N° 3 — URHC Zogbadjé",
    page_icon="⛪",
    layout="wide",
    initial_sidebar_state="expanded",
)

import presence_data as db

try:
    db.ensure_file()
except Exception as e:
    st.error(f"Initialisation base : {e}")

# Session
if "user" not in st.session_state:
    st.session_state["user"] = None
if "go_login" not in st.session_state:
    st.session_state["go_login"] = False

st.title("⛪ EXPOSÉ N° 3 — URHC PAROISSE ZOGBADJÉ")
st.caption("Gestion des présences du dimanche")

if st.session_state["user"]:
    u = st.session_state["user"]
    st.success(
        f"Connecté : **{u.get('Prenom', '')} {u.get('Nom', '')}** "
        f"({u.get('Role', 'MEMBRE')})"
    )
    st.markdown(
        """
Utilisez le **menu de gauche** :
- **Ma présence** — signaler votre présence
- **Rapport** — tableau du mois / PDF A3
- **Admin** — faire l’appel et valider les présences (responsables uniquement)
"""
    )
    if st.button("Se déconnecter"):
        st.session_state["user"] = None
        st.rerun()
else:
    st.info(
        "Vous n’êtes pas connecté. "
        "Allez dans **Connexion** ou **Inscription** via le menu latéral."
    )
    c1, c2 = st.columns(2)
    with c1:
        st.page_link("pages/1_Connexion.py", label="🔑 Connexion", icon="🔑")
    with c2:
        st.page_link("pages/2_Inscription.py", label="📝 Inscription", icon="📝")
