import streamlit as st
import pandas as pd

import dechargement_data as dd
import importlib
importlib.reload(dd)   # relit le fichier à chaque affichage (évite l'ancienne version en mémoire)

st.set_page_config(page_title="Relevé fournisseur", page_icon="📒", layout="wide")

# --- Sécurité : recharge les fichiers modifiés et vérifie qu'ils sont à jour ---
if getattr(dd, "VERSION_DATA", 0) < 9:
    st.error(f"⚠️ Le fichier **dechargement_data.py** n'est pas à jour (fichier utilisé : "
             f"`{getattr(dd, '__file__', '?')}`). Remplacez-le par la dernière version, au même endroit que "
             "`app.py`, puis arrêtez et relancez Streamlit.")
    st.stop()

# ---------- Authentification (comme les autres pages) ----------
if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

# ---------- Produit actif (un fichier de déchargements par produit) ----------
if "produit_actif" not in st.session_state:
    st.session_state["produit_actif"] = dd.PRODUIT_DEFAUT
produit = st.sidebar.selectbox("📦 Produit", dd.PRODUITS, index=dd.PRODUITS.index(st.session_state["produit_actif"]),
                               format_func=dd.libelle_produit)
st.session_state["produit_actif"] = produit
dd.set_produit(produit)
st.sidebar.divider()

ICONES = {"PAYE": "🟢 PAYE", "PARTIEL": "🟡 PARTIEL", "NON PAYE": "🔴 NON PAYE", "SURPLUS": "🔵 SURPLUS", "—": "—"}

df, aff = dd.charger_complet()
bons = dd.charger_bons_commande()

st.title(f"📒 Relevé de compte fournisseur — {dd.libelle_produit(produit)}")

if df.empty:
    st.info(f"Aucun déchargement enregistré pour {dd.libelle_produit(produit)}. Ajoutez-en (ou importez votre Excel) dans la page « Déchargements ».")
    st.stop()

tab1, tab2 = st.tabs(["📒 Relevé d'un fournisseur", "👥 Situation de tous les fournisseurs"])

# ---------------------------------------------------------------- TAB 2 : tous
with tab2:
    st.subheader("Situation de tous les fournisseurs")
    g = (df.groupby(["Code_Frs", "Fournisseur"], dropna=False)
           .agg(Camions=("Transaction", "count"), Bons_Commande=("Ref_BC", lambda s: len({x for x in s if x})),
                Poids_KG=("Poids_Net_KG", "sum"), Montant_Du=("Montant_Du", "sum"),
                Montant_Paye=("Montant_Paye", "sum"), Solde=("Solde", "sum"))
           .reset_index().sort_values("Solde", ascending=False))
    t1, t2, t3 = st.columns(3)
    t1.metric("Fournisseurs", dd.fmt(len(g)))
    t2.metric("Total dû (FCFA)", dd.fmt(g["Montant_Du"].sum()))
    t3.metric("Reste à payer (FCFA)", dd.fmt(g["Solde"].sum()))
    st.dataframe(g, hide_index=True, width="stretch", height=480, column_config={
        "Code_Frs": "Code", "Bons_Commande": "Bons de commande",
        "Poids_KG": st.column_config.NumberColumn("Poids (kg)", format="%d"),
        "Montant_Du": st.column_config.NumberColumn("Dû", format="%d"),
        "Montant_Paye": st.column_config.NumberColumn("Payé", format="%d"),
        "Solde": st.column_config.NumberColumn("Reste à payer", format="%d")})
    st.caption("Solde positif = nous devons ce montant au fournisseur · négatif = surplus déjà payé.")
    st.download_button("⬇️ Télécharger (Excel)", data=dd.to_excel_bytes({"Fournisseurs": g}),
                       file_name="situation_fournisseurs.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------------- TAB 1 : un fournisseur
with tab1:
    frs = dd.tous_les_fournisseurs(df)
    opts = {dd.label_fournisseur(k, n): (k, n) for k, n in sorted(frs.items(), key=lambda x: x[1])}
    choix = st.selectbox("Fournisseur", list(opts.keys()))
    cle, nom = opts[choix]

    sub = dd.filtrer_fournisseur(df, cle, nom).sort_values(["Date_Dechargement", "Transaction"]).reset_index(drop=True)
    suivi = dd.suivi_par_bc(df, bons)
    suivi = suivi[(suivi["Code_Frs"] == cle) | suivi["Fournisseur"].map(lambda x: str(x).casefold() == nom.casefold())]

    if sub.empty and suivi.empty:
        st.info(f"Aucun déchargement ni bon de commande pour {nom}.")
        st.stop()

    solde_total = sub["Solde"].sum()
    st.markdown(f"### {nom}" + (f"  ·  `{cle}`" if cle != nom else ""))
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Camions", dd.fmt(len(sub)))
    k2.metric("Bons de commande", dd.fmt(len({x for x in sub["Ref_BC"] if x})))
    k3.metric("Poids net", f"{sub['Poids_Net_KG'].sum() / 1000:,.1f} T".replace(",", " "))
    k4.metric("Total dû (FCFA)", dd.fmt(sub["Montant_Du"].sum()))
    k5.metric("Total payé (FCFA)", dd.fmt(sub["Montant_Paye"].sum()))
    k6.metric("Reste à payer (FCFA)" if solde_total >= 0 else "Surplus payé (FCFA)", dd.fmt(abs(solde_total)))

    if solde_total > 0:
        st.warning(f"💰 Nous devons encore **{dd.fmt(solde_total)} FCFA** à {nom}.")
    elif solde_total < 0:
        st.info(f"🔵 {nom} a reçu **{dd.fmt(-solde_total)} FCFA** de plus que le montant dû (surplus / avance).")
    else:
        st.success(f"🟢 Compte de {nom} soldé.")

    # --- Relevé détaillé ---
    st.markdown("#### 📄 Relevé détaillé (une ligne par camion)")
    rel = sub.copy()
    rel["Solde_Cumule"] = rel["Solde"].cumsum()
    rel["Statut_Paiement"] = rel["Statut_Paiement"].map(ICONES)
    rel["Alerte_Camion"] = ["⚠️ Différent" if x else "" for x in rel["Camion_Diff"]]
    rel = rel[["Date_Dechargement", "Transaction", "Ref_BC", "N_Camion", "Camion_BC", "Alerte_Camion",
               "Weighment_Ticket", "Poids_Net_KG",
               "PU_FCFA_KG", "Montant_Du", "Montant_Paye", "Solde", "Solde_Cumule", "Statut_Paiement",
               "Bons_Paiement"]]
    st.dataframe(rel, hide_index=True, width="stretch", height=420, column_config={
        "Date_Dechargement": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
        "Ref_BC": "N° Bon de commande", "N_Camion": "Camion livré", "Camion_BC": "Camion sur le bon",
        "Alerte_Camion": "Alerte", "Weighment_Ticket": "Ticket pesée",
        "Poids_Net_KG": st.column_config.NumberColumn("Poids net (kg)", format="%d"),
        "PU_FCFA_KG": st.column_config.NumberColumn("PU", format="%d"),
        "Montant_Du": st.column_config.NumberColumn("Dû", format="%d"),
        "Montant_Paye": st.column_config.NumberColumn("Payé", format="%d"),
        "Solde": st.column_config.NumberColumn("Solde ligne", format="%d"),
        "Solde_Cumule": st.column_config.NumberColumn("Solde cumulé", format="%d"),
        "Statut_Paiement": "Statut", "Bons_Paiement": "Bon(s) de paiement"})

    # --- Paiements approuvés ---
    st.markdown("#### 💳 Paiements approuvés (bons de paiement rattachés)")
    pay = aff[aff["Transaction"].isin(sub["Transaction"])].merge(
        sub[["Transaction", "Ref_BC", "N_Camion"]], on="Transaction", how="left").sort_values("ID", ascending=False)
    pay = pay[["Date_Approbation", "Bons_Paiement", "Transaction", "Ref_BC", "N_Camion", "Montant", "Approuve_Par"]]
    if pay.empty:
        st.caption("Aucun paiement approuvé pour ce fournisseur.")
    else:
        st.dataframe(pay, hide_index=True, width="stretch", column_config={
            "Date_Approbation": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
            "Bons_Paiement": "Bon(s) de paiement", "Ref_BC": "N° Bon de commande", "N_Camion": "N° Camion",
            "Montant": st.column_config.NumberColumn("Montant", format="%d"), "Approuve_Par": "Approuvé par"})

    # --- Bons de commande ---
    st.markdown("#### 📑 Bons de commande du fournisseur")
    if suivi.empty:
        st.caption("Aucun bon de commande rattaché.")
    else:
        sv = suivi.copy()
        sv["Statut_Paiement"] = sv["Statut_Paiement"].map(ICONES)
        st.dataframe(sv[["Reference", "Produit", "Qte_Commandee_KG", "Poids_Livre_KG", "Ecart_KG", "Nb_Camions",
                         "Montant_Du", "Montant_Paye", "Solde", "Statut_Livraison", "Statut_Paiement"]],
                     hide_index=True, width="stretch", column_config={
            "Reference": "N° Bon de commande",
            "Qte_Commandee_KG": st.column_config.NumberColumn("Commandé (kg)", format="%d"),
            "Poids_Livre_KG": st.column_config.NumberColumn("Livré (kg)", format="%d"),
            "Ecart_KG": st.column_config.NumberColumn("Écart (kg)", format="%d"),
            "Nb_Camions": st.column_config.NumberColumn("Camions", format="%d"),
            "Montant_Du": st.column_config.NumberColumn("Dû", format="%d"),
            "Montant_Paye": st.column_config.NumberColumn("Payé", format="%d"),
            "Solde": st.column_config.NumberColumn("Solde", format="%d"),
            "Statut_Livraison": "Livraison", "Statut_Paiement": "Paiement"})

    # --- Paiements du journal pas encore rattachés ---
    st.markdown("#### 🔎 Paiements du journal pas encore rattachés à une livraison")
    journal = dd.bons_paiement_achats(dd.charger_journal())
    utilises = {b for t in aff["Bons_Paiement"] for b in dd.liste_bons_texte(t)}
    libres = journal[journal["Prestataire"].map(lambda p: dd.meme_nom(p, nom)) & ~journal["No_Transaction"].isin(utilises)]
    if libres.empty:
        st.caption("Rien à rattacher : tous les bons de paiement du journal pour ce nom sont déjà utilisés.")
    else:
        st.caption("Avances ou paiements trouvés dans « Suivi de paiement » au nom de ce fournisseur (nom rapproché "
                   "automatiquement — à vérifier). Rattachez-les dans la page « Déchargements » › « Approuver un paiement ».")
        st.dataframe(libres, hide_index=True, width="stretch", column_config={
            "No_Transaction": "N° bon de paiement", "Date": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
            "Montant": st.column_config.NumberColumn("Montant", format="%d")})
        st.metric("Total non rattaché (FCFA)", dd.fmt(libres["Montant"].sum()))

    st.download_button(
        "⬇️ Télécharger le relevé (Excel)",
        data=dd.to_excel_bytes({"Releve": rel.assign(Statut_Paiement=sub["Statut_Paiement"]), "Paiements": pay,
                                "Bons de commande": suivi}),
        file_name=f"releve_{nom.replace(' ', '_').replace('/', '-')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("---")
st.caption(f"NAP SARL — Relevé de compte fournisseur {dd.libelle_produit(produit)} · données : `{dd.fichier_actif()}`, `donnees_bons.xlsx`, `Suivi de paiement.xlsx`")
