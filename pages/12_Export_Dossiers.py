# -*- coding: utf-8 -*-
"""
NAP SARL — Export Conteneurs
1 LIGNE = 1 CONTENEUR
Le même BL et le même Booking peuvent apparaître sur plusieurs lignes.
"""
import streamlit as st
import pandas as pd
from datetime import date
import importlib

import export_dossiers_data as ed
import transitaires_data as td
importlib.reload(ed)
importlib.reload(td)

st.set_page_config(page_title="Export Conteneurs", page_icon="📦", layout="wide")

if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"
username = st.session_state.get("username") or ""
est_admin = username == "admin"

ICONES_STATUT = {
    "En cours": "🟠 En cours",
    "Chargé": "🟡 Chargé",
    "Embarqué": "🔵 Embarqué",
    "Arrivé": "🟣 Arrivé",
    "Livré": "🟢 Livré",
    "Annulé": "⚫ Annulé",
}

st.title("📦 Export — Conteneurs (1 ligne = 1 conteneur)")
st.caption(
    "Même **BL** et même **Booking** peuvent se répéter sur plusieurs lignes. "
    "Fichier : `export_dossiers.xlsx` · feuille **Conteneurs**"
)

# ---------- Sidebar : nouveau conteneur ----------
st.sidebar.header("➕ Nouveau conteneur")

opts_trs = td.options_selection(actifs_seulement=True)
labels_trs = ["— Choisir —"] + list(opts_trs.keys())
if len(opts_trs) == 0:
    st.sidebar.error("Aucun transitaire. Créez-en dans **Transitaires**.")
else:
    st.sidebar.caption(f"{len(opts_trs)} transitaire(s)")

with st.sidebar.form("form_nouveau_ctr", clear_on_submit=True):
    date_d = st.date_input("Date *", value=date.today())
    contrat = st.text_input("Contrat / PO")
    produit = st.selectbox("Produit *", ed.PRODUITS_EXPORT)
    description = st.text_input("Description")
    trs_sel = st.selectbox("Transitaire *", labels_trs)

    st.markdown("**Références compagnie (peuvent se répéter)**")
    num_booking = st.text_input("N° Booking", placeholder="Identique pour plusieurs conteneurs")
    num_bl = st.text_input("N° BL", placeholder="Identique pour plusieurs conteneurs")

    st.markdown("**Ce conteneur**")
    num_ctr = st.text_input("N° Conteneur *", placeholder="MSCU1234567")
    type_ctr = st.selectbox("Type", ed.TYPES_CTR, index=1)
    scelle = st.text_input("N° Scellé")
    poids = st.number_input("Poids (KG)", min_value=0.0, step=100.0, format="%.0f")
    prix = st.number_input("Prix (FCFA)", min_value=0.0, step=1000.0, format="%.0f")
    statut = st.selectbox("Statut", ed.STATUTS)
    compagnie = st.text_input("Compagnie maritime", placeholder="MSC, Maersk…")
    pol = st.text_input("Port chargement", value="Cotonou")
    pod = st.text_input("Port déchargement")
    navire = st.text_input("Navire")
    voyage = st.text_input("N° voyage")
    notes = st.text_area("Notes", height=50)

    st.caption(f"Montant ligne = **{prix:,.0f}** FCFA".replace(",", " "))
    submitted = st.form_submit_button("💾 Enregistrer ce conteneur", width="stretch")

    if submitted:
        if trs_sel == "— Choisir —":
            st.error("Choisissez un transitaire.")
        elif not num_ctr.strip():
            st.error("N° conteneur obligatoire.")
        else:
            code_trs, nom_trs = opts_trs[trs_sel]
            ok, msg = ed.ajouter_conteneur({
                "Date": date_d,
                "Contrat": contrat,
                "Produit": produit,
                "Description": description,
                "Code_Transitaire": code_trs,
                "Transitaire": nom_trs,
                "Numero_Booking": num_booking,
                "Numero_BL": num_bl,
                "Numero_Conteneur": num_ctr,
                "Type": type_ctr,
                "Scelle": scelle,
                "Poids_KG": poids,
                "Prix": prix,
                "Montant": prix,
                "Statut": statut,
                "Compagnie": compagnie,
                "Port_Chargement": pol,
                "Port_Dechargement": pod,
                "Navire": navire,
                "Voyage": voyage,
                "Notes": notes,
            }, name)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

# ---------- Données ----------
df = ed.charger_conteneurs()

if not df.empty:
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Conteneurs (lignes)", len(df))
    k2.metric("BL distincts", int(df["Numero_BL"].replace("", pd.NA).nunique(dropna=True)))
    k3.metric("Montant total", ed.fmt(df["Montant"].sum(), "FCFA"))
    k4.metric("Poids total", ed.fmt(df["Poids_KG"].sum(), "kg"))
else:
    st.info("Aucun conteneur. Utilisez le formulaire à gauche : **1 enregistrement = 1 conteneur**.")

# ---------- Liste principale ----------
st.markdown("---")
st.subheader("📋 Liste des conteneurs (1 ligne = 1 conteneur)")

if not df.empty:
    f1, f2, f3, f4, f5 = st.columns(5)
    f_trs = f1.selectbox("Transitaire", ["Tous"] + sorted([x for x in df["Transitaire"].dropna().unique() if x]))
    f_prod = f2.selectbox("Produit", ["Tous"] + sorted([x for x in df["Produit"].dropna().unique() if x]))
    f_stat = f3.selectbox("Statut", ["Tous"] + ed.STATUTS)
    f_bl = f4.selectbox("BL", ["Tous"] + sorted([x for x in df["Numero_BL"].dropna().unique() if x]))
    f_q = f5.text_input("Recherche", placeholder="Conteneur, Booking, contrat…")

    v = df.copy()
    if f_trs != "Tous":
        v = v[v["Transitaire"] == f_trs]
    if f_prod != "Tous":
        v = v[v["Produit"] == f_prod]
    if f_stat != "Tous":
        v = v[v["Statut"] == f_stat]
    if f_bl != "Tous":
        v = v[v["Numero_BL"] == f_bl]
    if f_q.strip():
        qq = f_q.strip().lower()
        v = v[v.apply(lambda r: qq in " ".join(r.astype(str)).lower(), axis=1)]

    v = v.sort_values(["Numero_BL", "Numero_Booking", "ID"], ascending=[False, False, False])
    aff = v.copy()
    aff["Statut"] = aff["Statut"].map(lambda s: ICONES_STATUT.get(s, s))

    st.caption(f"{len(v)} ligne(s) · BL et Booking peuvent se répéter · montant filtré {ed.fmt(v['Montant'].sum(), 'FCFA')}")
    st.dataframe(
        aff[[
            "ID", "Date", "Numero_BL", "Numero_Booking", "Numero_Conteneur", "Type",
            "Transitaire", "Produit", "Contrat", "Poids_KG", "Prix", "Montant",
            "Statut", "Compagnie", "Scelle",
        ]],
        hide_index=True, width="stretch", height=420,
        column_config={
            "Date": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
            "Numero_BL": "N° BL",
            "Numero_Booking": "N° Booking",
            "Numero_Conteneur": "N° Conteneur",
            "Poids_KG": st.column_config.NumberColumn("Poids (kg)", format="%d"),
            "Prix": st.column_config.NumberColumn("Prix", format="%d"),
            "Montant": st.column_config.NumberColumn("Montant", format="%d"),
        },
    )

    st.markdown("#### ⬇️ Exporter Excel")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button(
            "📋 Liste affichée",
            data=ed.to_excel_bytes({"Conteneurs": v}),
            file_name="conteneurs_filtres.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
    with c2:
        st.download_button(
            "📦 Tous les conteneurs",
            data=ed.to_excel_bytes({"Conteneurs": df}),
            file_name="conteneurs_complet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
    with c3:
        resume = ed.resume_par_bl(df)
        trs_df = td.charger_transitaires()
        st.download_button(
            "🚢 Conteneurs + Résumé BL + Transitaires",
            data=ed.to_excel_bytes({
                "Conteneurs": df,
                "Resume_par_BL": resume,
                "Transitaires": trs_df,
            }),
            file_name="export_conteneurs_bl_transitaires.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
            type="primary",
        )

# ---------- Résumé par BL ----------
st.markdown("---")
st.subheader("📑 Vue résumé par BL (agrégation)")
st.caption("Une ligne = un BL · Nb conteneurs, poids et montant cumulés")
resume = ed.resume_par_bl(df) if not df.empty else pd.DataFrame()
if resume.empty:
    st.caption("Pas encore de données.")
else:
    st.dataframe(
        resume,
        hide_index=True, width="stretch",
        column_config={
            "Nb_Conteneurs": "Nb ctr",
            "Poids_Total_KG": st.column_config.NumberColumn("Poids total", format="%d"),
            "Montant_Total": st.column_config.NumberColumn("Montant total", format="%d"),
            "Date_Premier": st.column_config.DateColumn("1re date", format="DD/MM/YYYY"),
            "Date_Dernier": st.column_config.DateColumn("Dernière", format="DD/MM/YYYY"),
        },
    )

# ---------- Modifier ----------
st.markdown("---")
st.subheader("✏️ Modifier un conteneur (une ligne)")
if df.empty:
    st.caption("Aucune ligne.")
else:
    opts = {
        f"#{int(r.ID)} · {r.Numero_Conteneur} · BL:{r.Numero_BL or '—'} · BK:{r.Numero_Booking or '—'} · {r.Transitaire}": int(r.ID)
        for r in df.itertuples()
    }
    choix = st.selectbox(
        "Ligne à modifier",
        [""] + list(opts.keys()),
        format_func=lambda x: "— Choisir —" if x == "" else x,
        key="mod_ctr",
    )
    if choix:
        id_l = opts[choix]
        row = df[df["ID"] == id_l].iloc[0]
        with st.form("form_mod_ctr"):
            c1, c2, c3 = st.columns(3)
            with c1:
                e_date = st.date_input("Date", value=row["Date"].date() if pd.notna(row["Date"]) else date.today())
                e_bl = st.text_input("N° BL", value=str(row["Numero_BL"] or ""))
                e_bk = st.text_input("N° Booking", value=str(row["Numero_Booking"] or ""))
                e_ctr = st.text_input("N° Conteneur", value=str(row["Numero_Conteneur"] or ""))
                e_type = st.selectbox("Type", ed.TYPES_CTR,
                                     index=ed.TYPES_CTR.index(row["Type"]) if row["Type"] in ed.TYPES_CTR else 1)
            with c2:
                e_contrat = st.text_input("Contrat", value=str(row["Contrat"] or ""))
                e_prod = st.selectbox("Produit", ed.PRODUITS_EXPORT,
                                     index=ed.PRODUITS_EXPORT.index(row["Produit"]) if row["Produit"] in ed.PRODUITS_EXPORT else 0)
                e_poids = st.number_input("Poids (KG)", min_value=0.0, value=float(row["Poids_KG"] or 0), step=100.0, format="%.0f")
                e_prix = st.number_input("Prix (FCFA)", min_value=0.0, value=float(row["Prix"] or 0), step=1000.0, format="%.0f")
                e_statut = st.selectbox("Statut", ed.STATUTS,
                                       index=ed.STATUTS.index(row["Statut"]) if row["Statut"] in ed.STATUTS else 0)
            with c3:
                e_scelle = st.text_input("Scellé", value=str(row["Scelle"] or ""))
                e_comp = st.text_input("Compagnie", value=str(row["Compagnie"] or ""))
                e_pol = st.text_input("POL", value=str(row["Port_Chargement"] or ""))
                e_pod = st.text_input("POD", value=str(row["Port_Dechargement"] or ""))
                e_notes = st.text_input("Notes", value=str(row["Notes"] or ""))
            if st.form_submit_button("💾 Enregistrer", type="primary", width="stretch"):
                ok, msg = ed.modifier_conteneur(id_l, {
                    "Date": e_date, "Numero_BL": e_bl, "Numero_Booking": e_bk,
                    "Numero_Conteneur": e_ctr, "Type": e_type, "Contrat": e_contrat,
                    "Produit": e_prod, "Poids_KG": e_poids, "Prix": e_prix, "Montant": e_prix,
                    "Statut": e_statut, "Scelle": e_scelle, "Compagnie": e_comp,
                    "Port_Chargement": e_pol, "Port_Dechargement": e_pod, "Notes": e_notes,
                }, name)
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

if not df.empty:
    with st.expander("🗑️ Supprimer un conteneur (une ligne)"):
        opts_sup = {
            f"#{int(r.ID)} · {r.Numero_Conteneur} · BL:{r.Numero_BL or '—'}": int(r.ID)
            for r in df.itertuples()
        }
        choix_sup = st.selectbox("Ligne à supprimer", list(opts_sup.keys()), key="sup_ctr")
        if st.checkbox("Je confirme", key="conf_sup") and st.button("🗑️ Supprimer", type="primary"):
            ok, msg = ed.supprimer_conteneur(opts_sup[choix_sup])
            st.success(msg) if ok else st.error(msg)
            st.rerun()

st.caption("NAP SARL — 1 ligne = 1 conteneur · BL et Booking répétables · Excel local exploitable")
