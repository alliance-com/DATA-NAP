# -*- coding: utf-8 -*-
"""
NAP SARL — Tirages bancaires & tableaux d'amortissement
"""
from __future__ import annotations

from datetime import date, datetime
import os

import pandas as pd
import streamlit as st

import emprunts_data as ed
from generer_amortissement_pdf import generer_pdf_amortissement

st.set_page_config(page_title="Emprunts & Amortissement", page_icon="🏦", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

st.title("🏦 Emprunts bancaires — Tirages & Amortissement")
st.caption(
    "Prêts / lignes de crédit · **Tirages** · **Tableau d'amortissement** · Suivi des échéances · "
    f"Fichier : `{ed.FICHIER}`"
)

try:
    ed.ensure_file()
except Exception as e:
    st.error(f"Init fichier : {e}")
    st.stop()

prets = ed.charger_prets()

tab_prets, tab_tirages, tab_amort, tab_suivi = st.tabs([
    "📋 Prêts / Lignes de crédit",
    "💰 Tirages",
    "📊 Tableau d'amortissement",
    "✅ Suivi des échéances",
])

# ========== PRETS ==========
with tab_prets:
    st.subheader("Nouveau prêt / ligne de crédit")
    with st.form("form_pret"):
        c1, c2, c3 = st.columns(3)
        with c1:
            banque = st.selectbox("Banque *", ed.BANQUES)
            reference = st.text_input("Référence contrat / dossier", placeholder="ex. CRD-2026-014")
            objet = st.text_input("Objet", placeholder="ex. Financement stock karité")
        with c2:
            montant = st.number_input("Montant autorisé (FCFA) *", min_value=0.0, step=1_000_000.0, format="%.0f")
            taux = st.number_input("Taux annuel (%) *", min_value=0.0, max_value=100.0, value=8.5, step=0.1, format="%.2f")
            duree = st.number_input("Durée (mois) *", min_value=1, max_value=360, value=12, step=1)
        with c3:
            perio = st.selectbox("Périodicité", ed.PERIODICITES, index=0)
            typ = st.selectbox(
                "Type d'amortissement",
                ed.TYPES_AMORT,
                format_func=lambda x: (
                    "Annuité constante (échéance fixe)"
                    if x == "ANNUITE_CONSTANTE"
                    else "Amortissement constant (capital fixe)"
                ),
            )
            d_sig = st.date_input("Date signature", value=date.today())
            d_deb = st.date_input("Date début amortissement", value=date.today())
        notes = st.text_input("Notes")
        ok = st.form_submit_button("💾 Enregistrer le prêt", type="primary")
    if ok:
        ok2, msg = ed.ajouter_pret({
            "Banque": banque,
            "Reference": reference,
            "Objet": objet,
            "Montant_Autorise": montant,
            "Taux_Annuel_Pct": taux,
            "Duree_Mois": duree,
            "Periodicite": perio,
            "Type_Amort": typ,
            "Date_Signature": d_sig,
            "Date_Debut": d_deb,
            "Notes": notes,
        }, user=name)
        st.success(msg) if ok2 else st.error(msg)
        if ok2:
            st.rerun()

    st.markdown("---")
    st.subheader("Prêts enregistrés")
    if prets.empty:
        st.info("Aucun prêt. Créez une ligne de crédit ci-dessus.")
    else:
        aff = prets.copy()
        aff["Date_Signature"] = pd.to_datetime(aff["Date_Signature"], errors="coerce").dt.strftime("%d/%m/%Y")
        aff["Date_Debut"] = pd.to_datetime(aff["Date_Debut"], errors="coerce").dt.strftime("%d/%m/%Y")
        # enrichir tirages
        aff["Tiré"] = aff["ID_Pret"].map(lambda i: ed.total_tirages(i))
        aff["Reste à tirer"] = aff["Montant_Autorise"] - aff["Tiré"]
        st.dataframe(
            aff[[
                "ID_Pret", "Banque", "Reference", "Objet", "Montant_Autorise", "Tiré", "Reste à tirer",
                "Taux_Annuel_Pct", "Duree_Mois", "Periodicite", "Type_Amort", "Date_Debut", "Statut",
            ]],
            hide_index=True,
            width="stretch",
            column_config={
                "ID_Pret": "N°",
                "Montant_Autorise": st.column_config.NumberColumn("Autorisé", format="%d"),
                "Tiré": st.column_config.NumberColumn("Tiré", format="%d"),
                "Reste à tirer": st.column_config.NumberColumn("Reste à tirer", format="%d"),
                "Taux_Annuel_Pct": st.column_config.NumberColumn("Taux %", format="%.2f"),
                "Duree_Mois": "Durée (mois)",
            },
        )

# ========== TIRAGES ==========
with tab_tirages:
    st.subheader("Enregistrer un tirage")
    if prets.empty:
        st.warning(
            "Aucun prêt / ligne de crédit enregistré. "
            "Allez d’abord dans l’onglet **📋 Prêts / Lignes de crédit** pour créer "
            "par exemple une ligne ECOBANK (montant autorisé), puis revenez ici pour le tirage."
        )
        st.info(
            "Exemple : ligne de crédit ECOBANK 1 000 000 000 F → ensuite tirage de 350 000 000 F "
            "pour *Purchase of Cotton seed* (comme votre lettre REF 044-2026)."
        )
    else:
        opts = {
            f"#{int(r.ID_Pret)} — {r.Banque} — {float(r.Montant_Autorise):,.0f} F ({ed._s(r.Reference) or 'sans réf.'})".replace(",", " "): int(r.ID_Pret)
            for _, r in prets.iterrows()
        }
        choix = st.selectbox("Prêt / ligne concernée", list(opts.keys()), key="tir_pret")
        pid = opts[choix]
        res = ed.resume_pret(pid)
        m1, m2, m3 = st.columns(3)
        m1.metric("Autorisé", f"{res['autorise']:,.0f} F".replace(",", " "))
        m2.metric("Déjà tiré", f"{res['tire']:,.0f} F".replace(",", " "))
        m3.metric("Reste à tirer", f"{res['reste_tirer']:,.0f} F".replace(",", " "))

        with st.form("form_tirage"):
            t1, t2 = st.columns(2)
            with t1:
                d_tir = st.date_input("Date du tirage", value=date.today())
                mt_tir = st.number_input(
                    "Montant tiré (FCFA) *",
                    min_value=0.0,
                    max_value=float(max(res["reste_tirer"], 0)) if res["reste_tirer"] > 0 else 1e15,
                    step=100_000.0,
                    format="%.0f",
                )
            with t2:
                ref_b = st.text_input("Réf. banque / avis de crédit")
                notes_t = st.text_input("Notes")
            ok_t = st.form_submit_button("💾 Enregistrer le tirage", type="primary")
        if ok_t:
            ok2, msg = ed.ajouter_tirage({
                "ID_Pret": pid,
                "Date_Tirage": d_tir,
                "Montant": mt_tir,
                "Reference_Banque": ref_b,
                "Notes": notes_t,
            }, user=name)
            st.success(msg) if ok2 else st.error(msg)
            if ok2:
                # Sync → Lettres bancaires (TIRAGE_CFA)
                try:
                    import lettres_data as ld
                    from generer_lettre_bancaire import generer_lettre_word
                    pret_row = ed.pret_par_id(pid)
                    banque_l = ed._s(pret_row["Banque"]) if pret_row is not None else "ECOBANK BENIN"
                    ref_lettre = (ref_b.strip() if ref_b.strip() else None) or ld.prochaine_ref()
                    # éviter doublon référence
                    exist_l = ld.charger()
                    if not exist_l.empty and ref_lettre in exist_l["Reference"].astype(str).values:
                        ref_lettre = ld.prochaine_ref()
                    ok_l, msg_l, eid_l, ref_ok = ld.ajouter({
                        "Reference": ref_lettre,
                        "Date_Lettre": d_tir,
                        "Type": "TIRAGE_CFA",
                        "Banque": banque_l,
                        "Objet": f"Tirage for CFA {mt_tir:,.0f}".replace(",", " ") + (f" — {notes_t}" if notes_t else ""),
                        "Montant": mt_tir,
                        "Devise": "XOF",
                        "Compte_Credit": "112241067001",
                        "Beneficiaire": "",
                        "Motif": notes_t or (ed._s(pret_row.get("Objet")) if pret_row is not None else ""),
                        "Signataire": "AUTHORIZED SIGNATORY",
                        "Statut": "BROUILLON",
                    }, user=name)
                    if ok_l:
                        data_l = {
                            "Reference": ref_ok,
                            "Date_Lettre": d_tir,
                            "Type": "TIRAGE_CFA",
                            "Banque": banque_l,
                            "Objet": f"Tirage for CFA {mt_tir:,.0f}".replace(",", " "),
                            "Montant": mt_tir,
                            "Devise": "XOF",
                            "Compte_Credit": "112241067001",
                            "Motif": notes_t or "",
                            "Signataire": "AUTHORIZED SIGNATORY",
                        }
                        try:
                            chemin = generer_lettre_word(data_l)
                            ld.maj_fichier(eid_l, chemin)
                            st.info(f"✉️ Lettre bancaire **{ref_ok}** créée et Word généré (module Lettres).")
                            with open(chemin, "rb") as f:
                                st.download_button(
                                    "⬇️ Télécharger la lettre de tirage (Word)",
                                    data=f.read(),
                                    file_name=os.path.basename(chemin),
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key=f"dl_tir_let_{eid_l}",
                                )
                        except Exception as exw:
                            st.warning(f"Lettre enregistrée ({ref_ok}) mais Word : {exw}")
                    else:
                        st.caption(f"Lettre non créée : {msg_l}")
                except Exception as ex:
                    st.caption(f"Sync lettres bancaires : {ex}")
                st.rerun()

        st.markdown("#### Historique des tirages")
        tir = ed.charger_tirages(pid)
        if tir.empty:
            st.caption("Aucun tirage sur ce prêt.")
        else:
            aff = tir.copy()
            aff["Date_Tirage"] = pd.to_datetime(aff["Date_Tirage"], errors="coerce").dt.strftime("%d/%m/%Y")
            st.dataframe(
                aff[["ID_Tirage", "Date_Tirage", "Montant", "Reference_Banque", "Notes"]],
                hide_index=True,
                width="stretch",
                column_config={"Montant": st.column_config.NumberColumn("Montant", format="%d")},
            )

# ========== AMORTISSEMENT ==========
with tab_amort:
    st.subheader("Générer / afficher le tableau d'amortissement")
    if prets.empty:
        st.warning("Créez d'abord un prêt.")
    else:
        opts = {
            f"#{int(r.ID_Pret)} — {r.Banque} — {ed._s(r.Reference) or ed._s(r.Objet) or 'prêt'}": int(r.ID_Pret)
            for _, r in prets.iterrows()
        }
        choix = st.selectbox("Prêt", list(opts.keys()), key="am_pret")
        pid = opts[choix]
        pret = ed.pret_par_id(pid)
        res = ed.resume_pret(pid)

        c1, c2 = st.columns(2)
        with c1:
            mode_cap = st.radio(
                "Capital à amortir",
                ["tirages", "autorise", "manuel"],
                horizontal=True,
                format_func=lambda x: {
                    "tirages": f"Somme des tirages ({res['tire']:,.0f} F)".replace(",", " "),
                    "autorise": f"Montant autorisé ({res['autorise']:,.0f} F)".replace(",", " "),
                    "manuel": "Montant manuel",
                }[x],
            )
            cap_man = 0.0
            if mode_cap == "manuel":
                cap_man = st.number_input("Capital (FCFA)", min_value=0.0, value=float(res["tire"] or res["autorise"]), step=100_000.0, format="%.0f")
            ass = st.number_input("Assurance / frais par échéance (FCFA)", min_value=0.0, value=0.0, step=1000.0, format="%.0f")
        with c2:
            st.info(
                f"**{pret['Type_Amort']}** · {pret['Periodicite']} · "
                f"{int(pret['Duree_Mois'])} mois · taux {float(pret['Taux_Annuel_Pct']):.2f} %\n\n"
                "• *Annuité constante* : même montant d'échéance\n"
                "• *Amortissement constant* : même capital, intérêts dégressifs"
            )
            if st.button("⚙️ Générer / régénérer le tableau", type="primary"):
                if mode_cap == "tirages":
                    cap = res["tire"] if res["tire"] > 0 else res["autorise"]
                elif mode_cap == "autorise":
                    cap = res["autorise"]
                else:
                    cap = cap_man
                ok2, msg, _ = ed.generer_et_sauver_tableau(pid, capital=cap, assurance=ass, ecraser=True, user=name)
                st.success(msg) if ok2 else st.error(msg)
                if ok2:
                    st.rerun()

        ech = ed.charger_echeances(pid)
        res = ed.resume_pret(pid)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Échéances", res["nb_echeances"])
        k2.metric("Total intérêts + capital", f"{res['total_echeances']:,.0f} F".replace(",", " "))
        k3.metric("Déjà payé", f"{res['total_paye']:,.0f} F".replace(",", " "))
        k4.metric("Capital restant", f"{res['capital_restant']:,.0f} F".replace(",", " "))

        if ech.empty:
            st.info("Aucun tableau généré pour ce prêt. Cliquez sur **Générer**.")
        else:
            aff = ech.copy()
            aff["Date_Echeance"] = pd.to_datetime(aff["Date_Echeance"], errors="coerce").dt.strftime("%d/%m/%Y")
            aff["Date_Paiement"] = pd.to_datetime(aff["Date_Paiement"], errors="coerce").dt.strftime("%d/%m/%Y")
            st.dataframe(
                aff[[
                    "N_Echeance", "Date_Echeance", "Capital_Debut", "Amortissement", "Interets",
                    "Assurance", "Echeance_Totale", "Capital_Fin", "Statut", "Montant_Paye",
                ]],
                hide_index=True,
                width="stretch",
                height=min(520, 40 + 28 * min(len(aff), 18)),
                column_config={
                    "N_Echeance": "N°",
                    "Capital_Debut": st.column_config.NumberColumn("Capital début", format="%d"),
                    "Amortissement": st.column_config.NumberColumn("Amortissement", format="%d"),
                    "Interets": st.column_config.NumberColumn("Intérêts", format="%d"),
                    "Assurance": st.column_config.NumberColumn("Assurance", format="%d"),
                    "Echeance_Totale": st.column_config.NumberColumn("Échéance", format="%d"),
                    "Capital_Fin": st.column_config.NumberColumn("Capital fin", format="%d"),
                    "Montant_Paye": st.column_config.NumberColumn("Payé", format="%d"),
                },
            )
            b1, b2 = st.columns(2)
            with b1:
                st.download_button(
                    "⬇️ Excel — tableau d'amortissement",
                    data=ed.to_excel_bytes(ech),
                    file_name=f"Amortissement_Pret_{pid}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            with b2:
                try:
                    pdf = generer_pdf_amortissement(pret, ech, res)
                    st.download_button(
                        "🖨️ PDF — tableau d'amortissement",
                        data=pdf,
                        file_name=f"Amortissement_Pret_{pid}.pdf",
                        mime="application/pdf",
                        type="primary",
                    )
                except Exception as ex:
                    st.error(f"PDF : {ex}")

# ========== SUIVI ECHEANCES ==========
with tab_suivi:
    st.subheader("Marquer une échéance comme payée")
    if prets.empty:
        st.warning("Aucun prêt.")
    else:
        opts = {
            f"#{int(r.ID_Pret)} — {r.Banque}": int(r.ID_Pret)
            for _, r in prets.iterrows()
        }
        choix = st.selectbox("Prêt", list(opts.keys()), key="su_pret")
        pid = opts[choix]
        ech = ed.charger_echeances(pid)
        if ech.empty:
            st.info("Générez d'abord le tableau d'amortissement.")
        else:
            a_payer = ech[ech["Statut"] != "PAYE"]
            if a_payer.empty:
                st.success("Toutes les échéances sont payées.")
            else:
                opt_e = {
                    f"N°{int(r.N_Echeance)} — {pd.Timestamp(r.Date_Echeance).strftime('%d/%m/%Y') if pd.notna(r.Date_Echeance) else '?'} — {float(r.Echeance_Totale):,.0f} F [{r.Statut}]".replace(",", " "): int(r.ID_Echeance)
                    for _, r in a_payer.iterrows()
                }
                sel_e = st.selectbox("Échéance", list(opt_e.keys()))
                eid = opt_e[sel_e]
                row = ech[ech["ID_Echeance"] == eid].iloc[0]
                with st.form("form_pay_ech"):
                    p1, p2 = st.columns(2)
                    with p1:
                        d_pay = st.date_input("Date de paiement", value=date.today())
                        mt_pay = st.number_input(
                            "Montant payé (FCFA)",
                            min_value=0.0,
                            value=float(row["Echeance_Totale"]),
                            step=1000.0,
                            format="%.0f",
                        )
                    with p2:
                        notes_p = st.text_input("Notes")
                    ok_p = st.form_submit_button("✅ Enregistrer le paiement d'échéance", type="primary")
                if ok_p:
                    ok2, msg = ed.marquer_echeance_payee(eid, montant=mt_pay, date_paiement=d_pay, notes=notes_p)
                    st.success(msg) if ok2 else st.error(msg)
                    if ok2:
                        st.rerun()

            st.markdown("#### Échéances du prêt")
            aff = ech.copy()
            aff["Date_Echeance"] = pd.to_datetime(aff["Date_Echeance"], errors="coerce").dt.strftime("%d/%m/%Y")
            st.dataframe(
                aff[["N_Echeance", "Date_Echeance", "Echeance_Totale", "Statut", "Montant_Paye"]],
                hide_index=True,
                width="stretch",
                column_config={
                    "Echeance_Totale": st.column_config.NumberColumn("Échéance", format="%d"),
                    "Montant_Paye": st.column_config.NumberColumn("Payé", format="%d"),
                },
            )



st.markdown("---")
st.subheader("🔗 Synchronisation avec Lettres bancaires")
st.caption(
    "Les tirages enregistrés ici créent automatiquement une lettre type **TIRAGE_CFA** "
    "dans le module **Lettres bancaires**. "
    "Inversement, vous pouvez importer ici les lettres de tirage pas encore liées à un prêt."
)
try:
    import lettres_data as ld
    lettres = ld.charger()
    tirages_all = ed.charger_tirages()
    refs_tir = set()
    if not tirages_all.empty:
        refs_tir = {ed._s(x) for x in tirages_all["Reference_Banque"] if ed._s(x)}
    if lettres is not None and not lettres.empty:
        tir_lettres = lettres[lettres["Type"].astype(str) == "TIRAGE_CFA"].copy()
        if not tir_lettres.empty:
            tir_lettres["_lie"] = tir_lettres["Reference"].astype(str).isin(refs_tir)
            n_nouv = int((~tir_lettres["_lie"]).sum())
            st.write(f"Lettres TIRAGE_CFA : **{len(tir_lettres)}** · déjà liées à un tirage : **{int(tir_lettres['_lie'].sum())}** · non liées : **{n_nouv}**")
            if n_nouv and not prets.empty:
                non_liees = tir_lettres[~tir_lettres["_lie"]]
                opts_p = {f"#{int(r.ID_Pret)} — {r.Banque}": int(r.ID_Pret) for _, r in prets.iterrows()}
                with st.form("form_import_lettre_tir"):
                    pret_cible = st.selectbox("Associer les lettres non liées au prêt", list(opts_p.keys()))
                    sel_refs = st.multiselect(
                        "Lettres à importer comme tirages",
                        options=non_liees["Reference"].astype(str).tolist(),
                    )
                    ok_imp = st.form_submit_button("📥 Importer comme tirages")
                if ok_imp and sel_refs:
                    pid_c = opts_p[pret_cible]
                    n_ok = 0
                    for _, lr in non_liees[non_liees["Reference"].astype(str).isin(sel_refs)].iterrows():
                        ok_i, msg_i = ed.ajouter_tirage({
                            "ID_Pret": pid_c,
                            "Date_Tirage": lr["Date_Lettre"] if pd.notna(lr["Date_Lettre"]) else date.today(),
                            "Montant": float(lr["Montant"] or 0),
                            "Reference_Banque": ed._s(lr["Reference"]),
                            "Notes": ed._s(lr.get("Motif") or lr.get("Objet")),
                        }, user=name)
                        if ok_i:
                            n_ok += 1
                        else:
                            st.warning(f"{lr['Reference']}: {msg_i}")
                    st.success(f"{n_ok} tirage(s) importé(s) depuis les lettres.")
                    if n_ok:
                        st.rerun()
            elif n_nouv and prets.empty:
                st.info("Créez un prêt pour pouvoir importer les lettres de tirage comme tirages.")
        else:
            st.caption("Aucune lettre type TIRAGE_CFA dans le registre lettres pour l’instant.")
    else:
        st.caption("Registre lettres vide ou module lettres_data absent.")
except Exception as ex:
    st.caption(f"Module lettres non disponible : {ex}")


st.caption("STE NAP SARL — Emprunts · Tirages · Tableau d'amortissement")
