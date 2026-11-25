# -*- coding: utf-8 -*-
"""Gestion des clients LOCAL / INTERNATIONAL — NAP SARL"""
import streamlit as st
import clients_data as cd

st.set_page_config(page_title="Clients", page_icon="👥", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "User"

st.title("👥 Clients — Local & International")
st.caption("Utilisés sur les déchargements **VENTE** et pour l'export.")

df = cd.charger()

c1, c2, c3 = st.columns(3)
c1.metric("Total clients", len(df))
c2.metric("Local", int((df["Type"] == "LOCAL").sum()) if not df.empty else 0)
c3.metric("International", int((df["Type"] == "INTERNATIONAL").sum()) if not df.empty else 0)

tab1, tab2, tab3 = st.tabs(["📋 Liste", "➕ Nouveau client", "✏️ Modifier"])

with tab1:
    f1, f2 = st.columns(2)
    with f1:
        ft = st.selectbox("Type", ["Tous", "LOCAL", "INTERNATIONAL"])
    with f2:
        q = st.text_input("Recherche", placeholder="Nom, code, pays…")
    vue = df.copy()
    if ft != "Tous":
        vue = vue[vue["Type"] == ft]
    if q.strip():
        ql = q.strip().lower()
        vue = vue[vue.apply(lambda r: ql in " ".join(r.astype(str)).lower(), axis=1)]
    st.dataframe(
        vue[["Code", "Nom", "Type", "Pays", "Ville", "Telephone", "Email", "Contact", "Actif"]],
        hide_index=True, width="stretch", height=420,
    )
    st.download_button(
        "⬇️ Excel",
        data=cd.to_excel_bytes(vue),
        file_name="clients_nap.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab2:
    with st.form("form_cli"):
        t = st.radio("Type *", cd.TYPES, horizontal=True)
        n1, n2 = st.columns(2)
        with n1:
            nom = st.text_input("Nom / Raison sociale *")
            pays = st.text_input("Pays", value="Bénin" if t == "LOCAL" else "")
            ville = st.text_input("Ville")
            adresse = st.text_area("Adresse", height=70)
        with n2:
            contact = st.text_input("Contact")
            tel = st.text_input("Téléphone")
            email = st.text_input("Email")
            nif = st.text_input("NIF / RCCM")
        notes = st.text_input("Notes")
        if st.form_submit_button("💾 Enregistrer", type="primary"):
            ok, msg = cd.ajouter({
                "Nom": nom, "Type": t, "Pays": pays, "Ville": ville,
                "Adresse": adresse, "Contact": contact, "Telephone": tel,
                "Email": email, "NIF": nif, "Notes": notes,
            }, user=name)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

with tab3:
    if df.empty:
        st.info("Aucun client.")
    else:
        opts = {f"{r.Code} — {r.Nom} ({r.Type})": r.Code for _, r in df.iterrows()}
        choix = st.selectbox("Client", list(opts.keys()))
        code = opts[choix]
        row = df[df["Code"] == code].iloc[0]
        with st.form("form_edit_cli"):
            t2 = st.radio(
                "Type", cd.TYPES,
                index=0 if row["Type"] != "INTERNATIONAL" else 1,
                horizontal=True,
            )
            e1, e2 = st.columns(2)
            with e1:
                nom2 = st.text_input("Nom", value=row["Nom"])
                pays2 = st.text_input("Pays", value=row["Pays"])
                ville2 = st.text_input("Ville", value=row["Ville"])
                adr2 = st.text_area("Adresse", value=row["Adresse"], height=70)
            with e2:
                contact2 = st.text_input("Contact", value=row["Contact"])
                tel2 = st.text_input("Téléphone", value=row["Telephone"])
                email2 = st.text_input("Email", value=row["Email"])
                nif2 = st.text_input("NIF", value=row["NIF"])
                actif2 = st.selectbox("Actif", ["OUI", "NON"], index=0 if row["Actif"] == "OUI" else 1)
            notes2 = st.text_input("Notes", value=row["Notes"])
            if st.form_submit_button("💾 Mettre à jour", type="primary"):
                ok, msg = cd.maj(code, {
                    "Nom": nom2, "Type": t2, "Pays": pays2, "Ville": ville2,
                    "Adresse": adr2, "Contact": contact2, "Telephone": tel2,
                    "Email": email2, "NIF": nif2, "Notes": notes2, "Actif": actif2,
                })
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

st.caption("NAP SARL — clients.xlsx (racine du projet)")
