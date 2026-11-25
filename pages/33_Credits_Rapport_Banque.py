# -*- coding: utf-8 -*-
"""
NAP SARL — Crédits bancaires + Rapport débit/crédit/solde + impression A3
"""
import streamlit as st
import pandas as pd
from datetime import date

import credits_bancaires_data as cb

st.set_page_config(page_title="Crédits & Rapport banque", page_icon="💳", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

st.title("💳 Crédits bancaires & Rapport complet")
st.caption(
    "**Débits** = Suivi de paiement · **Crédits** = base Excel + VERSEMENTS Suivi + **relevé banque importé** (rapprochement) · "
    "**Solde** = ouverture + crédits − débits"
)

st.sidebar.header("Filtres rapport")
banque = st.sidebar.selectbox("Banque", cb.BANQUES, index=0)
today = date.today()
periode = st.sidebar.date_input(
    "Période",
    value=(date(today.year, 1, 1), today),
)
if isinstance(periode, (tuple, list)) and len(periode) == 2:
    d_deb, d_fin = periode[0], periode[1]
else:
    d_deb, d_fin = date(today.year, 1, 1), today

resume = cb.resume_solde(banque, d_deb, d_fin)
mvt = cb.construire_mouvements(banque, d_deb, d_fin)

k1, k2, k3, k4 = st.columns(4)
k1.metric("Solde d'ouverture", f"{cb.fmt(resume['ouverture'])} F")
k2.metric("Total crédits (+)", f"{cb.fmt(resume['total_credits'])} F", f"{resume['nb_credits']} lignes")
k3.metric("Total débits (−)", f"{cb.fmt(resume['total_debits'])} F", f"{resume['nb_debits']} lignes")
k4.metric("Solde de clôture", f"{cb.fmt(resume['cloture'])} F")

tab1, tab2, tab3 = st.tabs([
    "➕ Saisir un crédit",
    "📋 Rapport complet (débit + crédit)",
    "🖨️ Impression A3 / Export",
])

# ---- TAB 1 saisie crédit ----
with tab1:
    st.subheader("Nouveau mouvement de crédit")
    st.info(
        "Utilisez cette saisie pour tout ce qui **augmente** le solde : "
        "versement caisse, virement client, prêt bancaire, remboursement, encaissement chèque…"
    )
    with st.form("form_credit"):
        c1, c2, c3 = st.columns(3)
        with c1:
            e_date = st.date_input("Date *", value=today)
            e_banque = st.selectbox("Banque *", cb.BANQUES, index=cb.BANQUES.index(banque) if banque in cb.BANQUES else 0)
            e_type = st.selectbox("Type de crédit *", cb.TYPES_CREDIT)
        with c2:
            e_tiers = st.text_input("Tiers / Déposant", placeholder="Nom client ou banque")
            e_montant = st.number_input("Montant (FCFA) *", min_value=0.0, step=1000.0)
            e_mode = st.selectbox("Mode", ["ESPECES", "VIREMENT", "CHEQUE", "MOBILE MONEY", "AUTRE"])
        with c3:
            e_ref = st.text_input("N° référence (auto si vide)", placeholder=cb.prochain_ref_credit())
            e_piece = st.text_input("N° pièce / chèque")
            e_lien = st.text_input("Lien prêt / dossier", placeholder="Ex. PRET SITE")
        e_motif = st.text_input("Motif / Description *", placeholder="Ex. Encaissement facture export")
        e_obs = st.text_input("Observation")
        if st.form_submit_button("💾 Enregistrer le crédit", type="primary"):
            ok, msg = cb.ajouter_credit({
                "Date": e_date,
                "Ref": e_ref,
                "Banque": e_banque,
                "Type": e_type,
                "Tiers": e_tiers,
                "Motif": e_motif,
                "Montant": e_montant,
                "Mode": e_mode,
                "Piece": e_piece,
                "Lien": e_lien,
                "Observation": e_obs,
            })
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    st.markdown("#### Tous les crédits liés (période / banque)")
    st.info(
        "Sources automatiques : **1)** saisie manuelle (ce formulaire / Excel) · "
        "**2)** feuille VERSEMENTS du Suivi · "
        "**3)** colonne **Credit** du relevé PDF/Excel importé dans **Rapprochement bancaire**."
    )
    credits = cb.charger_credits_tous(banque, d_deb, d_fin)
    if credits.empty:
        st.caption("Aucun crédit sur la période. Importez un relevé ou saisissez un crédit.")
    else:
        aff = credits.copy()
        aff["Date"] = pd.to_datetime(aff["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(aff, hide_index=True, width="stretch", height=320)
        st.caption(f"Total crédits affichés : **{cb.fmt(credits['Montant'].sum())} FCFA** · {len(credits)} ligne(s)")

    st.markdown("#### Fichier base Excel")
    st.caption(
        "Placez `Credits_Bancaires_NAP.xlsx` à la racine du projet DATA. "
        "Vous pouvez aussi saisir directement dans Excel (feuille CREDITS + SOLDES_DEPART)."
    )

# ---- TAB 2 rapport ----
with tab2:
    st.subheader(f"Mouvements {banque} — {d_deb.strftime('%d/%m/%Y')} → {d_fin.strftime('%d/%m/%Y')}")
    if mvt.empty:
        st.warning("Aucun mouvement. Vérifiez la période, la banque, et que débits (Suivi) / crédits sont renseignés.")
    else:
        aff = mvt.copy()
        aff["Date"] = pd.to_datetime(aff["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
        # color hint via sens column
        st.dataframe(
            aff.rename(columns={
                "Sens": "Sens", "Ref": "Réf.", "Tiers": "Tiers", "Libelle": "Libellé",
                "Debit": "Débit", "Credit": "Crédit", "Solde": "Solde", "Piece": "Pièce",
            }),
            hide_index=True, width="stretch", height=420,
        )
        st.markdown(
            f"""
| | Montant |
|--|--:|
| Solde d'ouverture | **{cb.fmt(resume['ouverture'])}** |
| + Crédits | **{cb.fmt(resume['total_credits'])}** |
| − Débits | **{cb.fmt(resume['total_debits'])}** |
| **Solde de clôture** | **{cb.fmt(resume['cloture'])}** |
"""
        )

# ---- TAB 3 impression ----
with tab3:
    st.subheader("Export & impression A3")
    st.write(
        "Le PDF A3 (paysage) regroupe solde d'ouverture, tous les mouvements débit/crédit, "
        "solde courant ligne à ligne et solde de clôture."
    )
    col_a, col_b = st.columns(2)
    with col_a:
        try:
            pdf_bytes = cb.generer_pdf_a3(banque, d_deb, d_fin, mvt, resume)
            st.download_button(
                "🖨️ Télécharger rapport PDF A3",
                data=pdf_bytes,
                file_name=f"Rapport_Banque_{banque}_{d_deb}_{d_fin}.pdf",
                mime="application/pdf",
                type="primary",
            )
        except Exception as e:
            st.error(f"PDF indisponible : {e}")
            st.caption("Vérifiez : pip install reportlab")
    with col_b:
        xls = cb.to_excel_rapport(mvt, resume, banque, d_deb, d_fin)
        st.download_button(
            "⬇️ Export Excel du rapport",
            data=xls,
            file_name=f"Rapport_Banque_{banque}_{d_deb}_{d_fin}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.markdown("---")
st.caption("NAP SARL — Débits (Suivi paiement) + Crédits (base dédiée) = vision claire du solde par banque")
