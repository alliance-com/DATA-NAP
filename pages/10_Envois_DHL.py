# -*- coding: utf-8 -*-
"""
NAP SARL — Gestion des envois DHL
Date, Waybill, type de document, nb docs, poids, montant, mode paiement, pièces jointes.
"""
import streamlit as st
import pandas as pd
from datetime import date, datetime
import os

import dhl_data as dhl

st.set_page_config(page_title="Envois DHL", page_icon="📦", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

st.title("📦 Gestion des envois DHL")
st.caption(
    f"Registre des envois · pièces jointes dans `{dhl.DOSSIER_DOCS}/` · fichier `{dhl.FICHIER_DHL}`"
)

df = dhl.charger()

# ==================== SIDEBAR : nouvel envoi ====================
st.sidebar.header("➕ Nouvel envoi DHL")

with st.sidebar.form("form_dhl", clear_on_submit=True):
    date_envoi = st.date_input("Date d'envoi *", value=date.today())
    waybill = st.text_input("N° Waybill *", placeholder="ex. 1234567890")
    type_doc = st.selectbox("Type de document *", dhl.TYPES_DOC)
    nb_docs = st.number_input("Nombre de documents *", min_value=1, value=1, step=1)
    poids = st.number_input("Poids (kg) *", min_value=0.0, value=0.5, step=0.1, format="%.2f")
    montant = st.number_input("Montant (FCFA) *", min_value=0, value=0, step=500)
    mode = st.selectbox("Mode de paiement *", dhl.MODES_PAIEMENT)
    destinataire = st.text_input("Destinataire", placeholder="Nom / société destinataire")
    pays = st.text_input("Pays de destination", placeholder="ex. Chine, Inde, UE…")
    description = st.text_area("Description / objet", placeholder="Précisez le contenu de l'envoi…")
    remarque = st.text_input("Remarque", placeholder="Optionnel")

    st.markdown("**Documents envoyés** (chargement des fichiers)")
    fichiers = st.file_uploader(
        "Charger les documents (PDF, images, Word…)",
        type=None,
        accept_multiple_files=True,
        help="Ces fichiers sont enregistrés avec l'envoi (pas dans le bloc du haut du tableau).",
        key="up_new",
    )

    submitted = st.form_submit_button("💾 Enregistrer l'envoi", type="primary", use_container_width=True)

if submitted:
    if not waybill.strip():
        st.sidebar.error("Le N° Waybill est obligatoire.")
    elif poids <= 0:
        st.sidebar.error("Le poids doit être supérieur à 0.")
    else:
        ok, msg, eid = dhl.ajouter(
            {
                "Date_Envoi": date_envoi,
                "Waybill": waybill.strip(),
                "Type_Document": type_doc,
                "Nb_Documents": nb_docs,
                "Poids_KG": poids,
                "Montant": montant,
                "Mode_Paiement": mode,
                "Destinataire": destinataire.strip(),
                "Pays_Destination": pays.strip(),
                "Description": description.strip(),
                "Remarque": remarque.strip(),
            },
            uploaded_files=fichiers,
            utilisateur=name,
        )
        if ok:
            st.sidebar.success(msg)
            st.rerun()
        else:
            st.sidebar.error(msg)

# ==================== KPI ====================
k1, k2, k3, k4 = st.columns(4)
k1.metric("Envois", len(df))
k2.metric("Poids total (kg)", f"{df['Poids_KG'].sum():,.1f}".replace(",", " ") if not df.empty else "0")
k3.metric("Montant total (FCFA)", f"{df['Montant'].sum():,.0f}".replace(",", " ") if not df.empty else "0")
k4.metric(
    "Ce mois",
    int(
        (
            (df["Date_Envoi"].dt.month == date.today().month)
            & (df["Date_Envoi"].dt.year == date.today().year)
        ).sum()
    )
    if not df.empty
    else 0,
)

st.markdown("---")

# ==================== Filtres ====================
f1, f2, f3, f4 = st.columns(4)
with f1:
    f_way = st.text_input("🔍 Waybill", placeholder="Rechercher…")
with f2:
    types = ["Tous"] + sorted({t for t in df["Type_Document"].tolist() if t}) if not df.empty else ["Tous"]
    f_type = st.selectbox("Type de document", types)
with f3:
    modes = ["Tous"] + dhl.MODES_PAIEMENT
    f_mode = st.selectbox("Mode de paiement", modes)
with f4:
    f_dest = st.text_input("Destinataire", placeholder="Filtrer…")

vue = df.copy()
if not vue.empty:
    if f_way.strip():
        vue = vue[vue["Waybill"].str.contains(f_way.strip(), case=False, na=False)]
    if f_type != "Tous":
        vue = vue[vue["Type_Document"] == f_type]
    if f_mode != "Tous":
        vue = vue[vue["Mode_Paiement"] == f_mode]
    if f_dest.strip():
        vue = vue[vue["Destinataire"].str.contains(f_dest.strip(), case=False, na=False)]

# ==================== Tableau ====================
st.subheader("📋 Registre des envois")

if vue.empty:
    st.info("Aucun envoi enregistré. Utilisez le formulaire à gauche.")
else:
    aff = vue.copy()
    aff["Date_Envoi"] = pd.to_datetime(aff["Date_Envoi"], errors="coerce").dt.strftime("%d/%m/%Y")
    aff["Nb_fichiers"] = aff["Fichiers"].map(lambda x: len([p for p in str(x).split("|") if p.strip()]) if x else 0)
    st.dataframe(
        aff[
            [
                "ID", "Date_Envoi", "Waybill", "Type_Document", "Nb_Documents",
                "Poids_KG", "Montant", "Mode_Paiement", "Destinataire", "Pays_Destination",
                "Description", "Nb_fichiers", "Cree_Par",
            ]
        ],
        hide_index=True,
        width="stretch",
        height=380,
        column_config={
            "ID": "ID",
            "Date_Envoi": "Date d'envoi",
            "Waybill": "N° Waybill",
            "Type_Document": "Type de document",
            "Nb_Documents": st.column_config.NumberColumn("Nb docs", format="%d"),
            "Poids_KG": st.column_config.NumberColumn("Poids (kg)", format="%.2f"),
            "Montant": st.column_config.NumberColumn("Montant", format="%d"),
            "Mode_Paiement": "Mode paiement",
            "Destinataire": "Destinataire",
            "Pays_Destination": "Pays",
            "Description": "Description",
            "Nb_fichiers": st.column_config.NumberColumn("Fichiers joints", format="%d"),
            "Cree_Par": "Créé par",
        },
    )

# ==================== Détail / documents / modification ====================
st.markdown("---")
st.subheader("📎 Détail & documents d'un envoi")

if df.empty:
    st.caption("Aucun envoi à afficher.")
else:
    opts = {
        f"#{int(r.ID)} · {r.Waybill} · {r.Type_Document} · "
        f"{r.Date_Envoi.strftime('%d/%m/%Y') if pd.notna(r.Date_Envoi) else '—'}"
        : int(r.ID)
        for _, r in df.iterrows()
    }
    choix = st.selectbox("Sélectionner un envoi", list(opts.keys()))
    eid = opts[choix]
    row = df[df["ID"] == eid].iloc[0]

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            f"""
| | |
|---|---|
| **Waybill** | `{row['Waybill']}` |
| **Date** | {row['Date_Envoi'].strftime('%d/%m/%Y') if pd.notna(row['Date_Envoi']) else '—'} |
| **Type** | {row['Type_Document']} |
| **Nb documents** | {int(row['Nb_Documents'])} |
| **Poids** | {row['Poids_KG']} kg |
| **Montant** | {int(row['Montant']):,} FCFA |
| **Paiement** | {row['Mode_Paiement']} |
| **Destinataire** | {row['Destinataire'] or '—'} |
| **Pays** | {row['Pays_Destination'] or '—'} |
""".replace(",", " ")
        )
        if row["Description"]:
            st.write("**Description :**", row["Description"])
        if row["Remarque"]:
            st.write("**Remarque :**", row["Remarque"])

    with c2:
        st.markdown("**Documents chargés pour cet envoi**")
        fichiers = dhl.liste_fichiers(row["Fichiers"])
        if not fichiers:
            st.info("Aucun fichier joint pour le moment.")
        else:
            for fp in fichiers:
                nom = os.path.basename(fp)
                try:
                    with open(fp, "rb") as f:
                        st.download_button(
                            f"⬇️ {nom}",
                            data=f.read(),
                            file_name=nom,
                            key=f"dl_{eid}_{nom}",
                        )
                except Exception as e:
                    st.caption(f"{nom} — erreur lecture : {e}")

        st.markdown("**Ajouter d'autres documents**")
        extra = st.file_uploader(
            "Joindre des fichiers à cet envoi",
            accept_multiple_files=True,
            key=f"up_extra_{eid}",
        )
        if st.button("📎 Ajouter les fichiers", key=f"btn_extra_{eid}"):
            if not extra:
                st.warning("Choisissez au moins un fichier.")
            else:
                ok, msg = dhl.modifier(
                    eid,
                    {
                        "Date_Envoi": row["Date_Envoi"] if pd.notna(row["Date_Envoi"]) else date.today(),
                        "Waybill": row["Waybill"],
                        "Type_Document": row["Type_Document"],
                        "Nb_Documents": row["Nb_Documents"],
                        "Poids_KG": row["Poids_KG"],
                        "Montant": row["Montant"],
                        "Mode_Paiement": row["Mode_Paiement"],
                        "Destinataire": row["Destinataire"],
                        "Pays_Destination": row["Pays_Destination"],
                        "Description": row["Description"],
                        "Remarque": row["Remarque"],
                    },
                    uploaded_files=extra,
                    utilisateur=name,
                )
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

    with st.expander("✏️ Modifier cet envoi"):
        with st.form(f"mod_{eid}"):
            m_date = st.date_input(
                "Date d'envoi",
                value=row["Date_Envoi"].date() if pd.notna(row["Date_Envoi"]) else date.today(),
            )
            m_way = st.text_input("Waybill", value=row["Waybill"])
            m_type = st.selectbox(
                "Type de document",
                dhl.TYPES_DOC,
                index=dhl.TYPES_DOC.index(row["Type_Document"]) if row["Type_Document"] in dhl.TYPES_DOC else 0,
            )
            m_nb = st.number_input("Nb documents", min_value=1, value=max(1, int(row["Nb_Documents"])))
            m_poids = st.number_input("Poids (kg)", min_value=0.0, value=float(row["Poids_KG"]), format="%.2f")
            m_mont = st.number_input("Montant", min_value=0, value=int(row["Montant"]))
            m_mode = st.selectbox(
                "Mode paiement",
                dhl.MODES_PAIEMENT,
                index=dhl.MODES_PAIEMENT.index(row["Mode_Paiement"]) if row["Mode_Paiement"] in dhl.MODES_PAIEMENT else 0,
            )
            m_dest = st.text_input("Destinataire", value=row["Destinataire"])
            m_pays = st.text_input("Pays", value=row["Pays_Destination"])
            m_desc = st.text_area("Description", value=row["Description"])
            m_rem = st.text_input("Remarque", value=row["Remarque"])
            if st.form_submit_button("Enregistrer les modifications"):
                ok, msg = dhl.modifier(
                    eid,
                    {
                        "Date_Envoi": m_date,
                        "Waybill": m_way,
                        "Type_Document": m_type,
                        "Nb_Documents": m_nb,
                        "Poids_KG": m_poids,
                        "Montant": m_mont,
                        "Mode_Paiement": m_mode,
                        "Destinataire": m_dest,
                        "Pays_Destination": m_pays,
                        "Description": m_desc,
                        "Remarque": m_rem,
                    },
                    utilisateur=name,
                )
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

    if st.button("🗑️ Supprimer cet envoi du registre", key=f"del_{eid}"):
        ok, msg = dhl.supprimer(eid)
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

# ==================== Export ====================
st.markdown("---")
st.subheader("⬇️ Export")
if not df.empty:
    st.download_button(
        "Télécharger le registre DHL (Excel)",
        data=dhl.to_excel_bytes(df),
        file_name=f"Envois_DHL_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

st.caption("NAP SARL — Envois DHL · waybill · documents joints · modes de paiement")
