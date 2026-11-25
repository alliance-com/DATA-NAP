# -*- coding: utf-8 -*-
"""
NAP SARL — Rapport fournisseurs
Synthèse + détail BC + liste plate « 1 camion non déchargé = 1 ligne » (Excel prêt à filtrer).
"""
import re
import streamlit as st
import pandas as pd
from datetime import date

import dechargement_data as dd

st.set_page_config(page_title="Rapport Fournisseurs", page_icon="📊", layout="wide")

if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

# ---------- Produit actif ----------
if "produit_actif" not in st.session_state:
    st.session_state["produit_actif"] = dd.PRODUIT_DEFAUT
produit = st.sidebar.selectbox(
    "📦 Produit",
    dd.PRODUITS,
    index=dd.PRODUITS.index(st.session_state["produit_actif"]),
    format_func=dd.libelle_produit,
)
st.session_state["produit_actif"] = produit
dd.set_produit(produit)

# ---------- Données ----------
df, aff = dd.charger_complet()
bons = dd.charger_bons_commande()
suivi = dd.suivi_par_bc(df, bons)

# Enrichir suivi avec camion + date émission du bon
if not suivi.empty and not bons.empty:
    bmap = bons.drop_duplicates("Reference").set_index("Reference")
    suivi = suivi.copy()
    if "Numero_Tracteur" in bmap.columns:
        suivi["Camion_BC"] = suivi["Reference"].map(bmap["Numero_Tracteur"].map(dd._s)).fillna("")
    else:
        suivi["Camion_BC"] = ""
    if "Date_Emission" in bmap.columns:
        suivi["Date_Emission"] = pd.to_datetime(
            suivi["Reference"].map(bmap["Date_Emission"]), errors="coerce", dayfirst=True
        )
    else:
        suivi["Date_Emission"] = pd.NaT
    if "Chauffeur" in bmap.columns:
        suivi["Chauffeur"] = suivi["Reference"].map(bmap["Chauffeur"].map(dd._s)).fillna("")
    else:
        suivi["Chauffeur"] = ""
else:
    suivi = suivi.copy() if not suivi.empty else pd.DataFrame()
    for c in ("Camion_BC", "Chauffeur"):
        if not suivi.empty and c not in suivi.columns:
            suivi[c] = ""
    if not suivi.empty and "Date_Emission" not in suivi.columns:
        suivi["Date_Emission"] = pd.NaT


def liste_camions_non_decharges(suivi_bc):
    """Table plate : 1 ligne = 1 camion non déchargé (fournisseur peut se répéter).

    Source = bons EN ATTENTE ou EN COURS ayant un N° camion sur le bon.
    Si plusieurs plaques séparées par / ou &, une ligne par plaque.
    """
    cols = [
        "Fournisseur", "N_Camion", "Ref_BC", "Date_Emission", "Chauffeur",
        "Qte_Commandee_KG", "Poids_Livre_KG", "Reste_KG", "Statut_Livraison", "Produit",
    ]
    if suivi_bc is None or suivi_bc.empty:
        return pd.DataFrame(columns=cols)

    rows = []
    sous = suivi_bc[suivi_bc["Statut_Livraison"].isin(["EN ATTENTE", "EN COURS"])].copy()
    for _, r in sous.iterrows():
        frs = dd._s(r.get("Fournisseur"))
        cam_brut = dd._s(r.get("Camion_BC"))
        ref = dd._s(r.get("Reference"))
        if not cam_brut:
            # bon sans camion saisi : une ligne avec camion vide pour ne pas perdre le BC
            plaques = [""]
        else:
            # une ligne par plaque (BP1681RB/BT3066RB → 2 lignes éventuelles, ou 1 ligne regroupée)
            # Par défaut : 1 ligne = texte camion complet du bon (plus simple pour l'exploitation)
            # Option "éclater" gérée par le paramètre éclater plus bas → on produit le détail ici
            plaques = [cam_brut]

        qte = float(r.get("Qte_Commandee_KG") or 0)
        liv = float(r.get("Poids_Livre_KG") or 0)
        reste = qte - liv
        date_e = r.get("Date_Emission")
        chauffeur = dd._s(r.get("Chauffeur"))
        statut = dd._s(r.get("Statut_Livraison"))
        prod = dd._s(r.get("Produit")) or dd.get_produit()

        for cam in plaques:
            rows.append({
                "Fournisseur": frs,
                "N_Camion": cam,
                "Ref_BC": ref,
                "Date_Emission": date_e,
                "Chauffeur": chauffeur,
                "Qte_Commandee_KG": qte,
                "Poids_Livre_KG": liv,
                "Reste_KG": reste,
                "Statut_Livraison": statut,
                "Produit": prod,
            })

    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values(["Fournisseur", "Ref_BC"]).reset_index(drop=True)
    return out


def eclater_plaques(df_cam):
    """Découpe 'BP1681RB/BT3066RB' en 2 lignes (même fournisseur, même BC)."""
    if df_cam is None or df_cam.empty:
        return df_cam
    rows = []
    for _, r in df_cam.iterrows():
        cam = dd._s(r.get("N_Camion"))
        parts = [p for p in re.split(r"[/&,;]+", cam) if p.strip()] if cam else [""]
        if not parts:
            parts = [""]
        for p in parts:
            d = r.to_dict()
            d["N_Camion"] = p.strip()
            rows.append(d)
    return pd.DataFrame(rows)


def synthese_fournisseurs(suivi_bc, liste_cam):
    if suivi_bc is None or suivi_bc.empty:
        return pd.DataFrame()

    s = suivi_bc.copy()
    s["Fournisseur"] = s["Fournisseur"].map(dd._s)
    s = s[s["Fournisseur"] != ""]

    # nb camions non déch. depuis la liste plate
    nb_cam = {}
    if liste_cam is not None and not liste_cam.empty:
        nb_cam = liste_cam.groupby("Fournisseur").size().to_dict()
        cam_txt = (
            liste_cam.groupby("Fournisseur")["N_Camion"]
            .apply(lambda s: " · ".join(x for x in s if x))
            .to_dict()
        )
    else:
        cam_txt = {}

    rows = []
    for frs, g in s.groupby("Fournisseur", sort=True):
        nb_bons = len(g)
        nb_livres = int((g["Statut_Livraison"] == "LIVRE").sum())
        nb_cours = int((g["Statut_Livraison"] == "EN COURS").sum())
        nb_att = int((g["Statut_Livraison"] == "EN ATTENTE").sum())
        nb_dech = int((g["Statut_Livraison"] == "DECHARGE").sum())
        qte_cmd = float(g["Qte_Commandee_KG"].sum())
        qte_liv = float(g["Poids_Livre_KG"].sum())
        reste = qte_cmd - qte_liv
        pct = round(qte_liv / qte_cmd * 100, 1) if qte_cmd > 0 else 0
        du = float(g["Montant_Du"].sum())
        paye = float(g["Montant_Paye"].sum())

        if qte_cmd <= 0:
            alerte = "—"
        elif reste > 0 and (nb_att > 0 or nb_cours > 0):
            alerte = "🔴 Reste à livrer"
        elif reste > 0:
            alerte = "🟠 Écart poids"
        elif reste < 0:
            alerte = "🔵 Surplus livré"
        elif pct >= 98:
            alerte = "🟢 Complet"
        else:
            alerte = "🟡 En cours"

        rows.append({
            "Fournisseur": frs,
            "Nb_Bons": nb_bons,
            "Nb_Bons_Livres": nb_livres + nb_dech,
            "Nb_Bons_En_Cours": nb_cours,
            "Nb_Bons_Attente": nb_att,
            "Qte_Commandee_KG": qte_cmd,
            "Qte_Livree_KG": qte_liv,
            "Reste_KG": reste if qte_cmd > 0 else 0,
            "Pct_Livre": pct,
            "Nb_Camions_Non_Decharges": int(nb_cam.get(frs, 0)),
            "Camions_Non_Decharges": cam_txt.get(frs, ""),
            "Montant_Du": du,
            "Montant_Paye": paye,
            "Solde_Paiement": du - paye,
            "Alerte": alerte,
        })
    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values("Reste_KG", ascending=False).reset_index(drop=True)
    return out


liste_cam = liste_camions_non_decharges(suivi)
syn = synthese_fournisseurs(suivi, liste_cam)

# ---------- UI ----------
st.title(f"📊 Rapport fournisseurs — {dd.libelle_produit(produit)}")
st.caption(
    "Synthèse · détail BC · **liste plate 1 camion = 1 ligne** (téléchargeable Excel). "
    f"`{dd.fichier_actif()}` · `{dd.FICHIER_BONS}`"
)

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Fournisseurs", dd.fmt(len(syn)))
k2.metric("Bons émis", dd.fmt(int(syn["Nb_Bons"].sum()) if not syn.empty else 0))
k3.metric("Bons en attente", dd.fmt(int(syn["Nb_Bons_Attente"].sum()) if not syn.empty else 0))
k4.metric("Qté commandée", f"{(syn['Qte_Commandee_KG'].sum()/1000 if not syn.empty else 0):,.1f} T".replace(",", " "))
k5.metric("Reste à livrer", f"{(syn['Reste_KG'].sum()/1000 if not syn.empty else 0):,.1f} T".replace(",", " "))
k6.metric("Lignes camions non déch.", dd.fmt(len(liste_cam)))

tab1, tab2, tab3 = st.tabs([
    "🚛 Camions non déchargés (1 ligne = 1 camion)",
    "👥 Synthèse fournisseurs",
    "📑 Détail par bon de commande",
])

# ==================== TAB 1 : liste plate ====================
with tab1:
    st.subheader("🚛 Camions non déchargés — format Excel (1 camion / 1 ligne)")
    st.caption(
        "Le **nom du fournisseur se répète** autant de fois qu’il a de camions non déchargés. "
        "Idéal pour filtre, TCD et suivi terrain."
    )

    eclater = st.checkbox(
        "Éclater les doubles plaques (ex. BP1681RB/BT3066RB → 2 lignes)",
        value=False,
        help="Si coché, chaque plaque séparée par / ou & devient une ligne.",
    )
    data_cam = eclater_plaques(liste_cam) if eclater else liste_cam.copy()

    t1, t2 = st.columns(2)
    with t1:
        frs_cam = st.selectbox(
            "Fournisseur",
            ["Tous"] + (sorted(data_cam["Fournisseur"].unique().tolist()) if not data_cam.empty else []),
            key="cam_frs",
        )
    with t2:
        stat_cam = st.multiselect(
            "Statut",
            ["EN ATTENTE", "EN COURS"],
            default=["EN ATTENTE", "EN COURS"],
            key="cam_stat",
        )

    vue_cam = data_cam.copy() if not data_cam.empty else pd.DataFrame()
    if not vue_cam.empty:
        if frs_cam != "Tous":
            vue_cam = vue_cam[vue_cam["Fournisseur"] == frs_cam]
        if stat_cam:
            vue_cam = vue_cam[vue_cam["Statut_Livraison"].isin(stat_cam)]

    if vue_cam.empty:
        st.success("Aucun camion non déchargé pour ce filtre.")
    else:
        st.caption(
            f"**{len(vue_cam)} ligne(s)** · **{vue_cam['Fournisseur'].nunique()} fournisseur(s)** · "
            f"reste total **{dd.fmt(vue_cam['Reste_KG'].sum(), 'kg')}**"
        )
        aff_cam = vue_cam.copy()
        st.dataframe(
            aff_cam,
            hide_index=True,
            width="stretch",
            height=480,
            column_config={
                "Fournisseur": st.column_config.TextColumn("Fournisseur", width="medium"),
                "N_Camion": st.column_config.TextColumn("N° Camion", width="medium"),
                "Ref_BC": "N° Bon de commande",
                "Date_Emission": st.column_config.DateColumn("Date émission BC", format="DD/MM/YYYY"),
                "Chauffeur": "Chauffeur",
                "Qte_Commandee_KG": st.column_config.NumberColumn("Commandé (kg)", format="%d"),
                "Poids_Livre_KG": st.column_config.NumberColumn("Déjà livré (kg)", format="%d"),
                "Reste_KG": st.column_config.NumberColumn("Reste (kg)", format="%d"),
                "Statut_Livraison": "Statut",
                "Produit": "Produit",
            },
        )

        # Export dédié — colonnes dans un ordre clair pour Excel
        export_cam = vue_cam.copy()
        export_cam = export_cam.rename(columns={
            "Fournisseur": "Fournisseur",
            "N_Camion": "N_Camion",
            "Ref_BC": "N_Bon_Commande",
            "Date_Emission": "Date_Emission_BC",
            "Chauffeur": "Chauffeur",
            "Qte_Commandee_KG": "Qte_Commandee_KG",
            "Poids_Livre_KG": "Poids_Livre_KG",
            "Reste_KG": "Reste_KG",
            "Statut_Livraison": "Statut_Livraison",
            "Produit": "Produit",
        })
        # ordre colonnes
        ord_cols = [
            "Fournisseur", "N_Camion", "N_Bon_Commande", "Date_Emission_BC", "Chauffeur",
            "Qte_Commandee_KG", "Poids_Livre_KG", "Reste_KG", "Statut_Livraison", "Produit",
        ]
        export_cam = export_cam[[c for c in ord_cols if c in export_cam.columns]]

        cdl1, cdl2 = st.columns(2)
        with cdl1:
            st.download_button(
                "⬇️ Télécharger camions non déchargés (Excel)",
                data=dd.to_excel_bytes({"Camions_Non_Decharges": export_cam}),
                file_name=f"Camions_Non_Decharges_{produit}_{date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="dl_cam",
            )
        with cdl2:
            st.download_button(
                "⬇️ Tout le rapport (synthèse + camions + bons)",
                data=dd.to_excel_bytes({
                    "Camions_Non_Decharges": export_cam,
                    "Synthese_Fournisseurs": syn if not syn.empty else pd.DataFrame(),
                    "Detail_Bons": suivi if not suivi.empty else pd.DataFrame(),
                }),
                file_name=f"Rapport_Fournisseurs_{produit}_{date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_all",
            )

# ==================== TAB 2 : synthèse ====================
with tab2:
    st.subheader("👥 Synthèse par fournisseur")
    f1, f2 = st.columns(2)
    with f1:
        f_frs = st.selectbox(
            "Fournisseur",
            ["Tous"] + (sorted(syn["Fournisseur"].tolist()) if not syn.empty else []),
            key="syn_frs",
        )
    with f2:
        f_alerte = st.selectbox(
            "Alerte",
            ["Toutes", "🔴 Reste à livrer", "🟠 Écart poids", "🟡 En cours", "🟢 Complet", "🔵 Surplus livré"],
            key="syn_alerte",
        )
    vue = syn.copy() if not syn.empty else pd.DataFrame()
    if not vue.empty:
        if f_frs != "Tous":
            vue = vue[vue["Fournisseur"] == f_frs]
        if f_alerte != "Toutes":
            vue = vue[vue["Alerte"] == f_alerte]

    if vue.empty:
        st.info("Aucune donnée.")
    else:
        aff_syn = vue.copy()
        aff_syn["Commandé (T)"] = (aff_syn["Qte_Commandee_KG"] / 1000).round(2)
        aff_syn["Livré (T)"] = (aff_syn["Qte_Livree_KG"] / 1000).round(2)
        aff_syn["Reste (T)"] = (aff_syn["Reste_KG"] / 1000).round(2)
        st.dataframe(
            aff_syn[[
                "Fournisseur", "Nb_Bons", "Nb_Bons_Livres", "Nb_Bons_En_Cours", "Nb_Bons_Attente",
                "Commandé (T)", "Livré (T)", "Reste (T)", "Pct_Livre",
                "Nb_Camions_Non_Decharges", "Camions_Non_Decharges",
                "Montant_Du", "Montant_Paye", "Solde_Paiement", "Alerte",
            ]],
            hide_index=True, width="stretch", height=400,
            column_config={
                "Nb_Bons": st.column_config.NumberColumn("Bons reçus", format="%d"),
                "Nb_Bons_Livres": st.column_config.NumberColumn("Bons livrés", format="%d"),
                "Nb_Bons_En_Cours": st.column_config.NumberColumn("En cours", format="%d"),
                "Nb_Bons_Attente": st.column_config.NumberColumn("En attente", format="%d"),
                "Pct_Livre": st.column_config.NumberColumn("% livré", format="%.1f"),
                "Nb_Camions_Non_Decharges": st.column_config.NumberColumn("Nb camions non déch.", format="%d"),
                "Camions_Non_Decharges": "N° Camions (aperçu)",
                "Montant_Du": st.column_config.NumberColumn("Dû", format="%d"),
                "Montant_Paye": st.column_config.NumberColumn("Payé", format="%d"),
                "Solde_Paiement": st.column_config.NumberColumn("Solde", format="%d"),
            },
        )

# ==================== TAB 3 : détail BC ====================
with tab3:
    st.subheader("📑 Détail par bon de commande")
    d1, d2 = st.columns(2)
    with d1:
        d_frs = st.selectbox(
            "Fournisseur",
            ["Tous"] + (sorted(suivi["Fournisseur"].unique().tolist()) if not suivi.empty else []),
            key="det_frs",
        )
    with d2:
        d_stat = st.multiselect(
            "Statut livraison",
            ["EN ATTENTE", "EN COURS", "LIVRE", "DECHARGE"],
            default=["EN ATTENTE", "EN COURS"],
            key="det_stat",
        )
    det = suivi.copy() if not suivi.empty else pd.DataFrame()
    if not det.empty:
        if d_frs != "Tous":
            det = det[det["Fournisseur"] == d_frs]
        if d_stat:
            det = det[det["Statut_Livraison"].isin(d_stat)]
        det = det.copy()
        det["Reste_KG"] = det["Qte_Commandee_KG"] - det["Poids_Livre_KG"]
        det["Camion_Non_Decharge"] = det.apply(
            lambda r: dd._s(r.get("Camion_BC")) if dd._s(r.get("Statut_Livraison")) in ("EN ATTENTE", "EN COURS") else "",
            axis=1,
        )
        st.dataframe(
            det[[
                "Reference", "Fournisseur", "Camion_BC", "Camion_Non_Decharge",
                "Qte_Commandee_KG", "Poids_Livre_KG", "Reste_KG", "Pct_Livre",
                "Nb_Camions", "Statut_Livraison", "Statut_Paiement",
                "Montant_Du", "Montant_Paye", "Solde",
            ]],
            hide_index=True, width="stretch", height=420,
            column_config={
                "Reference": "N° Bon de commande",
                "Camion_BC": "Camion sur le bon",
                "Camion_Non_Decharge": "N° Camion non déchargé",
                "Qte_Commandee_KG": st.column_config.NumberColumn("Commandé (kg)", format="%d"),
                "Poids_Livre_KG": st.column_config.NumberColumn("Livré (kg)", format="%d"),
                "Reste_KG": st.column_config.NumberColumn("Reste (kg)", format="%d"),
                "Pct_Livre": st.column_config.NumberColumn("% livré", format="%.1f"),
                "Nb_Camions": st.column_config.NumberColumn("Camions déch.", format="%d"),
                "Montant_Du": st.column_config.NumberColumn("Dû", format="%d"),
                "Montant_Paye": st.column_config.NumberColumn("Payé", format="%d"),
                "Solde": st.column_config.NumberColumn("Solde", format="%d"),
            },
        )
        attente = det[det["Statut_Livraison"] == "EN ATTENTE"]
        if not attente.empty:
            st.error(
                f"**{len(attente)} bon(s) sans livraison** — quantité attendue : "
                f"**{dd.fmt(attente['Qte_Commandee_KG'].sum(), 'kg')}**"
            )

st.caption(
    "Onglet **Camions non déchargés** : chaque ligne = un camion encore attendu. "
    "Le fournisseur se répète. Export Excel prêt pour filtres / TCD."
)
