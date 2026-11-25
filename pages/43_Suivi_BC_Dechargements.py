# -*- coding: utf-8 -*-
"""
NAP SARL — Suivi Bons de Commande ↔ Déchargements
Filtres · Excel conditionnel · PDF A3 · Rattachement BC
"""
from datetime import date, datetime

import pandas as pd
import streamlit as st

import dechargement_data as dd
import suivi_bc_dechargement as sbc

st.set_page_config(page_title="Suivi BC / Déchargements", page_icon="📋", layout="wide")

if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"
username = st.session_state.get("username") or ""

st.title("📋 Suivi Bons de Commande ↔ Déchargements")
st.caption(
    "Statut de chaque bon : déchargé / non déchargé / partiel / camion différent. "
    "Export Excel avec couleurs + impression PDF A3."
)

# ---------- Filtres ----------
f1, f2, f3, f4 = st.columns([2, 2, 2, 2])
with f1:
    opts_prod = ["TOUS"] + list(dd.PRODUITS)
    produit_sel = st.selectbox(
        "Produit",
        opts_prod,
        format_func=lambda p: "Tous les produits" if p == "TOUS" else dd.libelle_produit(p),
    )
with f2:
    statut_sel = st.selectbox(
        "Statut déchargement",
        [
            ("TOUS", "Tous"),
            ("DECHARGE", "Déchargé (complet)"),
            ("NON_DECHARGE", "Non déchargé"),
            ("PARTIEL", "Partiel (en cours)"),
            ("CAMION_DIFF", "Camion différent du bon"),
        ],
        format_func=lambda x: x[1],
    )
with f3:
    d_deb = st.date_input("Émis du", value=date(date.today().year, 1, 1), key="sbc_deb")
with f4:
    d_fin = st.date_input("Émis au", value=date.today(), key="sbc_fin")

prod_code = None if produit_sel == "TOUS" else produit_sel
statut_code = statut_sel[0]

df = sbc.construire_suivi(
    produit=prod_code,
    statut=statut_code,
    date_debut=d_deb,
    date_fin=d_fin,
)

# KPIs
if df.empty:
    st.info("Aucun bon de commande pour ces filtres.")
else:
    n_ok = int((df["Statut_Livraison"] == "DECHARGE").sum())
    n_non = int((df["Statut_Livraison"] == "NON DECHARGE").sum())
    n_part = int((df["Statut_Livraison"] == "PARTIEL").sum())
    n_diff = int(df["Camion_Diff"].sum())
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Bons", len(df))
    k2.metric("🟢 Déchargés", n_ok)
    k3.metric("⚪ Non déchargés", n_non)
    k4.metric("🟠 Partiels", n_part)
    k5.metric("🟡 Camion différent", n_diff)

    # Légende
    st.markdown(
        """
| Couleur | Signification |
|--------|----------------|
| 🟢 **Vert** | Bon entièrement déchargé |
| ⚪ **Blanc** | Non déchargé |
| 🟠 **Orange** | Partiellement déchargé |
| 🟡 **Jaune** | Déchargé mais **camion différent** du bon — vérifier / rattacher |
"""
    )

    # Affichage tableau
    aff = df.copy()
    if "Date_Emission" in aff.columns:
        aff["Date_Emission"] = pd.to_datetime(aff["Date_Emission"], errors="coerce").dt.strftime("%d/%m/%Y")
    cols_show = [
        c for c in [
            "Reference", "Date_Emission", "Produit", "Fournisseur", "Quantite",
            "Camion_BC", "Camions_Livres", "Nb_Camions_Decharges", "Poids_Livre_KG",
            "Pct_Livre", "Transactions", "Tickets_Pesee", "Dates_Dechargement", "Statut",
        ] if c in aff.columns
    ]

    def _color_row(row):
        st_ = str(row.get("Statut", ""))
        if st_ == "DECHARGE":
            return ["background-color: #C6EFCE"] * len(row)
        if st_ == "CAMION DIFFERENT":
            return ["background-color: #FFEB9C"] * len(row)
        if st_ == "PARTIEL":
            return ["background-color: #FCE4D6"] * len(row)
        return [""] * len(row)

    try:
        styled = aff[cols_show].style.apply(_color_row, axis=1)
        st.dataframe(styled, hide_index=True, width="stretch", height=420)
    except Exception:
        st.dataframe(aff[cols_show], hide_index=True, width="stretch", height=420)

    # Export
    st.markdown("### 📤 Export")
    titre_f = (
        f"Produit : {dd.libelle_produit(prod_code) if prod_code else 'Tous'} · "
        f"Statut : {statut_sel[1]} · Émis du {d_deb.strftime('%d/%m/%Y')} au {d_fin.strftime('%d/%m/%Y')}"
    )
    e1, e2 = st.columns(2)
    with e1:
        try:
            xls = sbc.excel_conditionnel(df)
            st.download_button(
                "⬇️ Excel (mise en forme conditionnelle)",
                data=xls,
                file_name=f"Suivi_BC_Dechargements_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch",
            )
        except Exception as ex:
            st.error(f"Excel : {ex}")
    with e2:
        try:
            pdf = sbc.pdf_a3(df, titre_filtre=titre_f)
            st.download_button(
                "🖨️ Impression PDF A3",
                data=pdf,
                file_name=f"Suivi_BC_Dechargements_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                width="stretch",
            )
        except Exception as ex:
            st.error(f"PDF : {ex}")

# ---------- Zone jaune : camion différent + rattachement ----------
st.markdown("---")
st.subheader("🟡 Camion différent / Rattacher un N° de bon de commande")

tab_jaune, tab_rattache = st.tabs([
    "Cas camion différent (précision)",
    "Rattacher un BC à un déchargement",
])

with tab_jaune:
    jaunes = sbc.construire_suivi(produit=prod_code, statut="CAMION_DIFF")
    if jaunes.empty:
        st.success("Aucun bon avec camion livré différent du camion prévu sur le BC.")
    else:
        st.warning(
            f"**{len(jaunes)} bon(s)** déchargés avec un N° camion différent. "
            "Vérifiez le déchargement correspondant (transaction / ticket)."
        )
        for _, r in jaunes.iterrows():
            with st.expander(
                f"{r['Reference']} · {r['Fournisseur']} · BC camion **{r['Camion_BC'] or '—'}** → livré **{r['Camions_Livres'] or '—'}**"
            ):
                c1, c2 = st.columns(2)
                c1.markdown(
                    f"""
- **Produit** : {r.get('Produit')}
- **Qté commandée** : {r.get('Quantite')}
- **Camion sur le BC** : `{r.get('Camion_BC') or '—'}`
- **Camion(s) livré(s)** : `{r.get('Camions_Livres') or '—'}`
"""
                )
                c2.markdown(
                    f"""
- **Transaction(s)** : `{r.get('Transactions') or '—'}`
- **Ticket(s) pesée** : `{r.get('Tickets_Pesee') or '—'}`
- **Date(s) déch.** : {r.get('Dates_Dechargement') or '—'}
- **Poids livré** : {r.get('Poids_Livre_KG') or 0:,.0f} kg
- **Statut** : **{r.get('Statut')}**
""".replace(",", " ")
                )
                st.caption(
                    "Si le déchargement est correct mais le BC n’était pas renseigné ou est erroné, "
                    "utilisez l’onglet « Rattacher un BC »."
                )

with tab_rattache:
    st.caption(
        "Utile quand un camion a été déchargé **sans N° BC** (ou avec un mauvais N°) : "
        "vous rattachez le bon de commande à la transaction de pesée."
    )
    sans = sbc.dechargements_sans_bc(prod_code)
    # aussi proposer les jaunes (transaction déjà liée mais on peut corriger)
    opts_tr = {}
    if not sans.empty:
        for _, r in sans.iterrows():
            lab = (
                f"{r.get('Transaction')} · {r.get('Fournisseur')} · "
                f"camion {r.get('N_Camion') or '—'} · ticket {r.get('Weighment_Ticket') or '—'} · "
                f"{'sans BC'}"
            )
            opts_tr[lab] = (str(r.get("Transaction") or "").strip(), r.get("Produit_Fichier"))
    # déchargements liés aux jaunes
    if not jaunes.empty:
        for _, r in jaunes.iterrows():
            for tr in str(r.get("Transactions") or "").split("|"):
                tr = tr.strip()
                if not tr:
                    continue
                lab = f"{tr} · {r.get('Fournisseur')} · BC actuel {r.get('Reference')} (camion diff.)"
                opts_tr[lab] = (tr, r.get("Code_Produit") or prod_code)

    if not opts_tr:
        st.info("Aucun déchargement sans BC (ni cas camion différent) à rattacher pour ce filtre.")
    else:
        choix = st.selectbox("Déchargement (transaction)", list(opts_tr.keys()))
        tr_sel, prod_hint = opts_tr[choix]
        # liste des BC pour rattachement
        bons = sbc.charger_bons_bruts()
        if prod_code and not bons.empty:
            bons = bons[bons["Code_Produit"] == prod_code]
        refs = bons["Reference"].tolist() if not bons.empty else []
        ref_new = st.selectbox("N° Bon de commande à rattacher", refs)
        if st.button("🔗 Rattacher ce BC à la transaction", type="primary"):
            ok, msg = sbc.rattacher_bc_a_transaction(
                tr_sel, ref_new, produit_code=prod_hint or prod_code, utilisateur=name
            )
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

st.markdown("---")
st.caption("NAP SARL — Suivi BC / Déchargements · couleurs : vert = déchargé · jaune = camion différent · orange = partiel · blanc = non déchargé")
