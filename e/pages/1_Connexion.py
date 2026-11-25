# -*- coding: utf-8 -*-
import streamlit as st
import presence_data as db

st.set_page_config(page_title="Connexion — Exposé N° 3", page_icon="🔑", layout="centered")

if st.session_state.get("user"):
    st.success("Vous êtes déjà connecté.")
    st.page_link("app.py", label="← Accueil", icon="🏠")
    st.stop()

st.title("🔑 Connexion")
st.caption("Téléphone (10 chiffres) + email + code reçu à l’inscription")

# Préremplir si inscription vient de réussir
pre_tel = st.session_state.pop("prefill_tel", "")
pre_email = st.session_state.pop("prefill_email", "")
info_code = st.session_state.pop("inscription_code_msg", "")

if info_code:
    st.success(info_code)

with st.form("form_login"):
    tel = st.text_input("N° Téléphone *", value=pre_tel, max_chars=14, placeholder="0169060724")
    email = st.text_input("Email *", value=pre_email, placeholder="exemple@email.com")
    mdp = st.text_input(
        "Mot de passe / code",
        type="password",
        help="Code reçu à l’inscription, ou 4 derniers chiffres du téléphone",
    )
    ok = st.form_submit_button("Se connecter", type="primary", use_container_width=True)

if ok:
    user, err = db.authentifier(tel, email, mdp)
    if err:
        st.error(err)
    else:
        st.session_state["user"] = user
        st.success(f"Bienvenue {user.get('Prenom', '')} !")
        st.switch_page("app.py")

st.markdown("---")
st.page_link("pages/2_Inscription.py", label="Pas encore de compte ? S’inscrire", icon="📝")
