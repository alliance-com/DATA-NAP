# -*- coding: utf-8 -*-
"""NAP SARL — Gestion des Transitaires (comme les fournisseurs)."""
import streamlit as st
import pandas as pd
import importlib

import transitaires_data as td
importlib.reload(td)

st.set_page_config(page_title="Transitaires", page_icon="🚢", layout="wide")

if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"
username = st.session_state.get("username") or ""
est_admin = username == "admin"

st.title("🚢 Transitaires — STE NAP SARL")
st.caption("Création et gestion des comptes transitaires · Fichier : `transitaires.xlsx`")

# ---------- Sidebar : nouveau transitaire ----------
st.sidebar.header("➕ Nouveau transitaire")
with st.sidebar.form("form_nouveau_trs", clear_on_submit=True):
    nom = st.text_input("Nom / Raison sociale *")
    tel = st.text_input("Téléphone")
    email = st.text_input("Email")
    adresse = st.text_input("Adresse")
    rccm = st.text_input("RCCM")
    ifu = st.text_input("IFU")
    submitted = st.form_submit_button("💾 Enregistrer", width="stretch")
    if submitted:
        if not nom.strip():
            st.error("Le nom est obligatoire.")
        else:
            ok, msg, code = td.ajouter_transitaire(nom, tel, email, adresse, rccm, ifu, name)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

# ---------- Liste ----------
df = td.charger_transitaires()
st.subheader(f"📋 Liste des transitaires ({len(df)})")

if df.empty:
    st.info("Aucun transitaire. Créez-en un via le menu de gauche.")
else:
    q = st.text_input("Rechercher", placeholder="Nom, code, téléphone…")
    vue = df.copy()
    if q.strip():
        qq = q.strip().lower()
        vue = vue[vue.apply(lambda r: qq in " ".join(r.astype(str)).lower(), axis=1)]
    st.dataframe(
        vue[["Code", "Nom", "Telephone", "Email", "Adresse", "Actif", "Cree_Par", "Date_Creation"]],
        hide_index=True, width="stretch", height=350
    )
    from io import BytesIO
    from openpyxl.styles import Font, PatternFill
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        vue.to_excel(w, sheet_name="Transitaires", index=False)
        ws = w.sheets["Transitaires"]
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="87CEEB")
            cell.font = Font(bold=True, color="1B4F72")
    st.download_button(
        "⬇️ Exporter les transitaires (Excel)",
        data=buf.getvalue(),
        file_name="transitaires.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

# ---------- Modifier ----------
st.markdown("---")
st.subheader("✏️ Modifier un transitaire")
if df.empty:
    st.caption("Aucun transitaire à modifier.")
else:
    options = {f"{r['Code']} — {r['Nom']}": r["Code"] for _, r in df.iterrows()}
    choix = st.selectbox("Transitaire", [""] + list(options.keys()),
                         format_func=lambda x: "— Choisir —" if x == "" else x)
    if choix:
        code = options[choix]
        row = df[df["Code"] == code].iloc[0]
        with st.form("form_mod_trs"):
            c1, c2 = st.columns(2)
            with c1:
                e_nom = st.text_input("Nom *", value=row["Nom"])
                e_tel = st.text_input("Téléphone", value=row["Telephone"])
                e_email = st.text_input("Email", value=row["Email"])
            with c2:
                e_adr = st.text_input("Adresse", value=row["Adresse"])
                e_rccm = st.text_input("RCCM", value=row["RCCM"])
                e_ifu = st.text_input("IFU", value=row["IFU"])
                e_actif = st.selectbox("Actif", ["OUI", "NON"],
                                       index=0 if str(row["Actif"]).upper() == "OUI" else 1)
            if st.form_submit_button("💾 Enregistrer les modifications", type="primary"):
                if not e_nom.strip():
                    st.error("Le nom est obligatoire.")
                else:
                    ok, msg = td.modifier_transitaire(code, {
                        "Nom": e_nom, "Telephone": e_tel, "Email": e_email,
                        "Adresse": e_adr, "RCCM": e_rccm, "IFU": e_ifu, "Actif": e_actif,
                    })
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

st.caption("NAP SARL — Transitaires · les codes TRS-xxx sont utilisés dans les dossiers Export et le suivi des soldes")
