# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import presence_data as db

st.set_page_config(page_title="Admin — Exposé N° 3", page_icon="✅", layout="wide")

user = st.session_state.get("user")
if not user:
    st.warning("Veuillez vous connecter.")
    st.page_link("pages/1_Connexion.py", label="🔑 Connexion")
    st.stop()

if str(user.get("Role", "")).upper() != "ADMIN":
    st.error("Accès réservé au responsable (admin).")
    st.stop()

st.title("✅ Administration")
nom_aff = f"{user.get('Prenom', '')} {user.get('Nom', '')}".strip()

t0, t1, t2 = st.tabs(["📋 Faire l'appel", "Présences à valider", "Membres"])

# ───────────────────────── APPEL (sans internet côté membres) ─────────────────────────
with t0:
    st.subheader("Appel du dimanche")
    st.caption("Appelez un nom : la personne répond « présent » → cochez. "
               "La présence est validée immédiatement.")

    d = st.date_input("Dimanche", value=db.dimanche_courant_ou_dernier(),
                      format="DD/MM/YYYY", key="appel_date")
    if d.weekday() != 6:
        st.error("Choisissez un dimanche.")
    else:
        membres = db.charger_membres(actifs_only=True)
        if membres.empty:
            st.info("Aucun membre inscrit.")
        else:
            pres = db.charger_presences(date_debut=d, date_fin=d)
            valides, attente = set(), set()
            if not pres.empty:
                valides = set(pres.loc[pres["Statut"] == "VALIDE", "ID_Membre"].astype(int))
                attente = set(pres.loc[pres["Statut"] == "EN_ATTENTE", "ID_Membre"].astype(int))

            q = st.text_input("🔎 Rechercher un nom", key="appel_q").strip().lower()
            vue = membres
            if q:
                vue = membres[
                    membres["Nom"].str.lower().str.contains(q, regex=False)
                    | membres["Prenom"].str.lower().str.contains(q, regex=False)
                ]

            st.write(f"**Présents validés : {len(valides)} / {len(membres)}**")

            with st.form("form_appel"):
                cols = st.columns(2)
                for i, (_, m) in enumerate(vue.iterrows()):
                    mid = int(m["ID"])
                    label = f"{m['Nom']} {m['Prenom']}"
                    if mid in attente:
                        label += "  ·  (signalé en ligne)"
                    cols[i % 2].checkbox(
                        label,
                        value=(mid in valides) or (mid in attente),
                        key=f"appel_{d.isoformat()}_{mid}",
                    )
                ok = st.form_submit_button("💾 Enregistrer l'appel", type="primary",
                                           use_container_width=True)

            if ok:
                presents = [
                    int(m["ID"]) for _, m in vue.iterrows()
                    if st.session_state.get(f"appel_{d.isoformat()}_{int(m['ID'])}")
                ]
                affiches = [int(x) for x in vue["ID"]]
                good, msg = db.appliquer_appel(d, presents, affiches, nom_aff)
                if good:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

# ───────────────────────── VALIDATION des signalements en ligne ─────────────────────────
with t1:
    att = db.presences_en_attente()
    membres = db.charger_membres()
    noms = {}
    if not membres.empty:
        for _, r in membres.iterrows():
            noms[int(r["ID"])] = f"{r['Prenom']} {r['Nom']}"
    if att.empty:
        st.success("Aucune présence en attente.")
    else:
        for _, row in att.iterrows():
            mid = int(row["ID_Membre"])
            dtxt = pd.Timestamp(row["Date_Dimanche"]).strftime("%d/%m/%Y") if pd.notna(row["Date_Dimanche"]) else "?"
            c1, c2, c3, c4 = st.columns([3, 2, 1, 1])
            c1.write(f"**{noms.get(mid, '#'+str(mid))}**")
            c2.write(f"{dtxt} · {row['Signale_Le']}")
            if c3.button("✓ Valider", key=f"v{row['ID']}", type="primary"):
                ok, msg = db.valider_presence(int(row["ID"]), nom_aff, True)
                st.success(msg)
                st.rerun()
            if c4.button("✗ Refuser", key=f"r{row['ID']}"):
                ok, msg = db.valider_presence(int(row["ID"]), nom_aff, False)
                st.warning(msg)
                st.rerun()

# ───────────────────────── MEMBRES ─────────────────────────
with t2:
    mb = db.charger_membres()
    if mb.empty:
        st.info("Aucun membre.")
    else:
        cols = [c for c in [
            "ID", "Nom", "Prenom", "Telephone", "Email", "Profession",
            "Est_Etudiant", "Universite", "Filiere", "Annee_Etude", "Role", "Actif",
        ] if c in mb.columns]
        st.dataframe(mb[cols], hide_index=True, width="stretch")
