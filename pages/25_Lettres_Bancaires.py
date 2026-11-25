# -*- coding: utf-8 -*-
"""
NAP SARL — Gestion des lettres bancaires / courriers sortants
Génération Word + PDF (Tirage, paiement fournisseur, LC, justificatifs…)
"""
import streamlit as st
import pandas as pd
from datetime import date
import os

import lettres_data as ld
from generer_lettre_bancaire import generer_lettre_word, generer_lettre_pdf

st.set_page_config(page_title="Lettres Bancaires", page_icon="✉️", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

st.title("✉️ Lettres bancaires & courriers sortants")
st.caption(
    "Registre des lettres · modèles Tirage CFA, paiement fournisseur, extension LC, "
    "acceptation d'écart, réponse justificatifs · export Word et PDF"
)

df = ld.charger()

# KPI
k1, k2, k3 = st.columns(3)
k1.metric("Lettres enregistrées", len(df))
k2.metric("Générées", int((df["Statut"] == "GENERE").sum()) if not df.empty else 0)
k3.metric("Prochaine réf.", ld.prochaine_ref())

st.markdown("---")

tab_new, tab_list, tab_gen = st.tabs(["➕ Nouvelle lettre", "📋 Registre", "🖨️ Générer Word / PDF"])

# ========== NOUVELLE LETTRE ==========
with tab_new:
    st.subheader("Créer une lettre")
    typ = st.selectbox(
        "Type de lettre *",
        ld.TYPES,
        format_func=lambda t: {
            "TIRAGE_CFA": "Tirage CFA (demande de tirage)",
            "PAIEMENT_FOURNISSEUR": "Paiement fournisseur (USD/EUR)",
            "EXTENSION_LC": "Extension de lettre de crédit (LC)",
            "ACCEPTATION_ECART": "Acceptation d'écart documentaire",
            "REPONSE_JUSTIFICATIFS": "Réponse banque — pièces justificatives",
            "LETTRE_GENERALE": "Lettre générale / autre",
        }.get(t, t),
    )

    with st.form("form_lettre"):
        c1, c2, c3 = st.columns(3)
        with c1:
            ref = st.text_input("Référence", value=ld.prochaine_ref(), help="Ex. REF-046-2026")
            d_lettre = st.date_input("Date de la lettre", value=date.today())
            banque = st.selectbox("Banque / Destinataire", ld.BANQUES)
            if banque == "Autre":
                banque = st.text_input("Nom de la banque / destinataire")
        with c2:
            montant = st.number_input("Montant", min_value=0.0, step=1000.0, format="%.2f")
            devise = st.selectbox("Devise", ld.DEVISES, index=0 if typ == "TIRAGE_CFA" else 1)
            compte = st.selectbox("Compte à créditer", [""] + ld.COMPTES)
            if not compte:
                compte = st.text_input("N° compte (saisie libre)")
        with c3:
            benef = st.text_input("Bénéficiaire / Fournisseur", placeholder="ADITYA ENGINEERS…")
            ref_lc = st.text_input("Réf. LC", placeholder="CD123000019")
            ref_fact = st.text_input("Réf. facture", placeholder="A0047B")
            ref_tx = st.text_input("Réf. transaction banque", placeholder="M01INFT…")

        objet = st.text_input(
            "Objet",
            placeholder="Ex. Tirage for Cfa 350,000,000 — Purchase of Shea nuts",
        )
        motif = st.text_area(
            "Motif / précision",
            placeholder="For Purchase of Shea nuts / Construction…",
            height=70,
        )
        corps_libre = st.text_area(
            "Corps libre (optionnel — laisse vide pour modèle automatique)",
            height=120,
            help="Si renseigné, remplace entièrement le texte généré automatiquement.",
        )
        signataire = st.text_input("Signataire", value="AUTHORIZED SIGNATORY")
        ok = st.form_submit_button("💾 Enregistrer la lettre", type="primary")

    if ok:
        if not typ:
            st.error("Type obligatoire.")
        else:
            success, msg, eid, ref_ok = ld.ajouter(
                {
                    "Reference": ref.strip(),
                    "Date_Lettre": d_lettre,
                    "Type": typ,
                    "Banque": banque,
                    "Objet": objet,
                    "Montant": montant,
                    "Devise": devise,
                    "Compte_Credit": compte,
                    "Beneficiaire": benef,
                    "Ref_LC": ref_lc,
                    "Ref_Facture": ref_fact,
                    "Ref_Transaction": ref_tx,
                    "Motif": motif,
                    "Corps_Libre": corps_libre,
                    "Signataire": signataire,
                    "Statut": "BROUILLON",
                },
                user=name,
            )
            if success:
                st.success(msg)
                st.session_state["lettre_sel_id"] = eid
                st.rerun()
            else:
                st.error(msg)

# ========== REGISTRE ==========
with tab_list:
    st.subheader("Registre des lettres")
    if df.empty:
        st.info("Aucune lettre enregistrée.")
    else:
        f1, f2 = st.columns(2)
        with f1:
            f_type = st.selectbox("Filtrer type", ["Tous"] + ld.TYPES, key="flt_t")
        with f2:
            f_banque = st.selectbox(
                "Filtrer banque",
                ["Tous"] + sorted(df["Banque"].dropna().unique().tolist()),
                key="flt_b",
            )
        vue = df.copy()
        if f_type != "Tous":
            vue = vue[vue["Type"] == f_type]
        if f_banque != "Tous":
            vue = vue[vue["Banque"] == f_banque]
        aff = vue.copy()
        aff["Date_Lettre"] = pd.to_datetime(aff["Date_Lettre"], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(
            aff[[
                "ID", "Reference", "Date_Lettre", "Type", "Banque", "Objet",
                "Montant", "Devise", "Beneficiaire", "Statut", "Cree_Par",
            ]],
            hide_index=True,
            width="stretch",
            height=380,
        )
        st.download_button(
            "⬇️ Export Excel du registre",
            data=ld.to_excel_bytes(vue),
            file_name=f"Registre_Lettres_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ========== GENERATION ==========
with tab_gen:
    st.subheader("Générer Word et PDF")
    if df.empty:
        st.info("Enregistrez d'abord une lettre.")
    else:
        options = {
            f"#{int(r.ID)} · {r.Reference} · {r.Type} · "
            f"{r.Date_Lettre.strftime('%d/%m/%Y') if pd.notna(r.Date_Lettre) else '—'} · "
            f"{r.Objet[:40] if r.Objet else ''}": int(r.ID)
            for _, r in df.iterrows()
        }
        default_ix = 0
        if st.session_state.get("lettre_sel_id") in options.values():
            keys = list(options.keys())
            for i, k in enumerate(keys):
                if options[k] == st.session_state["lettre_sel_id"]:
                    default_ix = i
                    break
        choix = st.selectbox("Lettre à générer", list(options.keys()), index=default_ix)
        eid = options[choix]
        row = df[df["ID"] == eid].iloc[0]
        data = row.to_dict()

        st.markdown("**Aperçu des champs**")
        a1, a2, a3 = st.columns(3)
        a1.write(f"Réf. : **{data.get('Reference')}**")
        a2.write(f"Type : **{data.get('Type')}**")
        a3.write(f"Montant : **{data.get('Montant')} {data.get('Devise')}**")
        st.write(f"Banque : {data.get('Banque')}")
        st.write(f"Objet : {data.get('Objet')}")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("📄 Générer Word", type="primary", use_container_width=True):
                try:
                    chemin = generer_lettre_word(data)
                    ld.maj_fichier(eid, chemin)
                    st.success(f"Word généré : `{chemin}`")
                    with open(chemin, "rb") as f:
                        st.download_button(
                            "⬇️ Télécharger Word",
                            data=f.read(),
                            file_name=os.path.basename(chemin),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary",
                        )
                except Exception as e:
                    st.error(f"Erreur Word : {e}")
                    import traceback
                    st.code(traceback.format_exc())
        with c2:
            if st.button("📕 Générer PDF", type="primary", use_container_width=True):
                try:
                    pdf = generer_lettre_pdf(data)
                    nom = f"Lettre_{data.get('Reference', eid)}.pdf".replace(" ", "_")
                    st.success("PDF prêt.")
                    st.download_button(
                        "⬇️ Télécharger PDF",
                        data=pdf,
                        file_name=nom,
                        mime="application/pdf",
                        type="primary",
                    )
                except Exception as e:
                    st.error(f"Erreur PDF : {e}")
                    import traceback
                    st.code(traceback.format_exc())

st.markdown("---")
st.caption(
    "STE NAP SARL — Lettres bancaires · modèles alignés sur vos courriers Tirage / LC / justificatifs ECOBANK & BIIC"
)
