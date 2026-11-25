# -*- coding: utf-8 -*-
import streamlit as st
import presence_data as db

st.set_page_config(page_title="Inscription — Exposé N° 3", page_icon="📝", layout="centered")

if st.session_state.get("user"):
    st.info("Vous êtes déjà connecté.")
    st.page_link("app.py", label="← Accueil", icon="🏠")
    st.stop()

st.title("📝 Inscription — Exposé N° 3 · URHC Zogbadjé")
st.caption("Le premier inscrit devient responsable (admin).")

# Checkbox HORS formulaire pour afficher le sous-bloc immédiatement
est_etudiant = st.checkbox("Je suis étudiant(e)", key="chk_etudiant")

with st.form("form_insc"):
    c1, c2 = st.columns(2)
    with c1:
        nom = st.text_input("Nom *")
        prenom = st.text_input("Prénom *")
        tel = st.text_input("N° Téléphone (10 chiffres) *", max_chars=14)
        email = st.text_input("Email *")
    with c2:
        profession = st.text_input("Profession *", placeholder="Ex. Commerçant, Étudiant, Artisan…")
        mdp = st.text_input(
            "Mot de passe (optionnel)",
            type="password",
            help="Si vide : 4 derniers chiffres du téléphone",
        )

    univ = fac = fil = annee = ""
    if est_etudiant:
        st.markdown("##### Parcours étudiant")
        s1, s2 = st.columns(2)
        with s1:
            univ = st.text_input("Université / établissement *")
            fac = st.text_input("Faculté ou école")
        with s2:
            fil = st.text_input("Filière *")
            annee = st.text_input("Année d’étude", placeholder="Ex. Licence 2")

    ok = st.form_submit_button("S’inscrire", type="primary", use_container_width=True)

if ok:
    ok2, msg, mid = db.inscrire_membre(
        nom=nom,
        prenom=prenom,
        telephone=tel,
        email=email,
        profession=profession,
        est_etudiant=est_etudiant,
        universite=univ,
        faculte=fac,
        filiere=fil,
        annee_etude=annee,
        mot_de_passe=mdp,
    )
    if ok2:
        st.session_state["prefill_tel"] = db.normaliser_tel(tel)
        st.session_state["prefill_email"] = (email or "").strip().lower()
        st.session_state["inscription_code_msg"] = msg
        st.success(msg)
        st.switch_page("pages/1_Connexion.py")
    else:
        st.error(msg)

st.markdown("---")
st.page_link("pages/1_Connexion.py", label="Déjà inscrit ? Se connecter", icon="🔑")
