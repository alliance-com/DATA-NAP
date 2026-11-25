# -*- coding: utf-8 -*-
"""
NAP SARL — Gestion de la Caisse (dépenses, catégories, comptes client/fournisseur)
"""
import streamlit as st
import pandas as pd
from datetime import date, datetime
import os

import caisse_data as cx
from generer_caisse_rapport import generer_pdf_mouvements, generer_pdf_comptes, generer_pdf_fermeture

st.set_page_config(page_title="Caisse NAP", page_icon="💵", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

st.title("💵 Gestion de la Caisse — NAP SARL")
st.caption(
    "Suivi des mouvements de caisse · Dépenses prioritaires · "
    "Catégories · Comptes Client / Fournisseur"
)

# Init fichier
try:
    cx.ensure_file()
except Exception as e:
    st.error(f"Initialisation caisse : {e}")
    st.stop()

df_all = cx.charger_mouvements()
kpi_all = cx.indicateurs(df_all)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Mouvements", kpi_all["nb"])
m2.metric("Total dépenses", f"{kpi_all['depenses']:,.0f}".replace(",", " ") + " F")
m3.metric("Total entrées", f"{kpi_all['entrees']:,.0f}".replace(",", " ") + " F")
solde = kpi_all["solde"]
m4.metric(
    "Solde (E − D)",
    f"{solde:,.0f}".replace(",", " ") + " F",
    delta="positif" if solde >= 0 else "négatif",
)


# ===================== IMPORT SUIVI PAIEMENT → CAISSE =====================
with st.expander("📥 Importer les paiements caisse depuis Suivi de paiement", expanded=False):
    st.caption(
        "Reprend automatiquement les lignes **Espèces / Mobile Money / Caisse** du fichier "
        "**Suivi de paiement.xlsx** et les enregistre comme **dépenses** dans la Caisse NAP. "
        "Les N° de transaction déjà importés (réf. `SP-…`) sont ignorés."
    )
    only_cash = st.checkbox(
        "Uniquement modes caisse (Espèces, Mobile Money…)",
        value=True,
        key="imp_sp_only_cash",
        help="Décochez pour importer tous les paiements du Suivi comme dépenses caisse.",
    )
    try:
        cand, stats, err_ap = cx.apercu_import_suivi(seulement_caisse=only_cash)
    except Exception as e:
        cand, stats, err_ap = None, {}, str(e)
    if err_ap:
        st.warning(err_ap)
    else:
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Lignes suivi", stats.get("total_suivi", 0))
        s2.metric("Candidats caisse", stats.get("candidats", 0))
        s3.metric("Déjà en caisse", stats.get("deja_importes", 0))
        s4.metric("Nouveaux", stats.get("nouveaux", 0))
        if stats.get("fichier"):
            st.caption(f"Fichier source : `{stats['fichier']}`")
        if cand is not None and not cand.empty:
            prev = cand[~cand["_Deja"]].head(30) if "_Deja" in cand.columns else cand.head(30)
            if not prev.empty:
                show = prev.copy()
                if "Date" in show.columns:
                    show["Date"] = pd.to_datetime(show["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
                cols = [c for c in [
                    "Date", "No_Transaction", "Produit", "Prestataire", "Montant", "Mode", "_Ref_Import"
                ] if c in show.columns]
                st.dataframe(show[cols], hide_index=True, width="stretch", height=220)
        if st.button("📥 Importer maintenant dans la Caisse", type="primary", key="btn_imp_sp"):
            ok_i, msg_i, n_i = cx.importer_depuis_suivi_paiement(
                seulement_caisse=only_cash, user=name
            )
            if ok_i:
                st.success(msg_i)
                st.rerun()
            else:
                st.error(msg_i)


tab_saisie, tab_jour, tab_liste, tab_comptes, tab_cats = st.tabs([
    "➕ Saisie",
    "🗓️ Solde & fermeture du jour",
    "📋 Mouvements & rapports",
    "👤 Comptes Client / Fournisseur",
    "🏷️ Catégories & Tiers",
])

# ===================== SAISIE =====================
with tab_saisie:
    st.subheader("Nouveau mouvement de caisse")
    type_mouv = st.radio(
        "Type",
        ["DEPENSE", "ENTREE"],
        horizontal=True,
        format_func=lambda x: "💸 Dépense" if x == "DEPENSE" else "💰 Entrée",
    )

    # Mode HORS formulaire → sous-bloc chèque s'affiche tout de suite
    mode = st.selectbox("Mode de paiement *", cx.MODES, key="caisse_mode_saisie")
    num_cheque = ""
    nom_percepteur = ""
    if str(mode).casefold().startswith("chèque") or str(mode).casefold().startswith("cheque"):
        st.info("Mode **Chèque** : renseignez le N° de chèque et le nom du percepteur.")
        cq1, cq2 = st.columns(2)
        with cq1:
            num_cheque = st.text_input("N° de chèque *", key="caisse_num_chq", placeholder="Ex. 2000345")
        with cq2:
            nom_percepteur = st.text_input(
                "Nom du percepteur *",
                key="caisse_nom_perc",
                placeholder="Personne qui reçoit le chèque",
            )

    with st.form("form_caisse"):
        c1, c2, c3 = st.columns(3)
        with c1:
            d_mouv = st.date_input("Date *", value=date.today())
            montant = st.number_input("Montant (FCFA) *", min_value=0.0, step=500.0, value=0.0)
            produit_saisie = st.selectbox(
                "Produit",
                ["—"] + list(cx.PRODUITS),
                help="Lier le mouvement à un produit (Soja, Coton, Karité…)",
            )
        with c2:
            cats = cx.liste_categories()
            cat_choix = st.selectbox("Catégorie *", cats + ["➕ Nouvelle catégorie…"])
            nouvelle_cat = ""
            if cat_choix == "➕ Nouvelle catégorie…":
                nouvelle_cat = st.text_input("Nom de la nouvelle catégorie *")
            libelle = st.text_input("Libellé / objet", placeholder="Ex. Envoi documents DHL client X")
            reference = st.text_input("Référence (reçu, facture…)", placeholder="Optionnel")
        with c3:
            type_tiers = st.selectbox("Type de tiers", cx.TYPES_TIERS)
            tiers_existants = cx.liste_tiers(type_tiers)
            tiers_sel = st.selectbox(
                "Client / Fournisseur",
                ["—"] + tiers_existants + ["➕ Nouveau…"],
            )
            nouveau_tiers = ""
            if tiers_sel == "➕ Nouveau…":
                nouveau_tiers = st.text_input(f"Nom du {type_tiers.lower()} *")
            observation = st.text_input("Observation")

        ok = st.form_submit_button("💾 Enregistrer", type="primary", use_container_width=True)

    if ok:
        cat = nouvelle_cat.strip() if cat_choix == "➕ Nouvelle catégorie…" else cat_choix
        tiers = nouveau_tiers.strip() if tiers_sel == "➕ Nouveau…" else ("" if tiers_sel == "—" else tiers_sel)
        if not cat:
            st.error("Catégorie obligatoire.")
        elif montant <= 0:
            st.error("Montant doit être > 0.")
        elif tiers_sel == "➕ Nouveau…" and not tiers:
            st.error("Indiquez le nom du tiers.")
        else:
            # Validation chèque avant enregistrement
            if str(mode).casefold().startswith("chèque") or str(mode).casefold().startswith("cheque"):
                if not str(num_cheque).strip():
                    st.error("N° de chèque obligatoire.")
                    st.stop()
                if not str(nom_percepteur).strip():
                    st.error("Nom du percepteur obligatoire.")
                    st.stop()
            ok2, msg = cx.ajouter_mouvement({
                "Date": d_mouv,
                "Type": type_mouv,
                "Categorie": cat,
                "Libelle": libelle,
                "Tiers": tiers,
                "Type_Tiers": type_tiers if tiers else "",
                "Produit": "" if produit_saisie == "—" else produit_saisie,
                "Montant": montant,
                "Mode": mode,
                "Num_Cheque": num_cheque,
                "Nom_Percepteur": nom_percepteur,
                "Reference": reference,
                "Observation": observation,
            }, user=name)
            if ok2:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)


# ===================== SOLDE & FERMETURE JOUR =====================
with tab_jour:
    st.subheader("Solde de départ du jour")
    st.caption(
        "Chaque matin : saisissez le solde de départ **ou** reprenez automatiquement "
        "le solde de clôture d’hier."
    )
    j1, j2 = st.columns(2)
    with j1:
        jour_sel = st.date_input("Jour concerné", value=date.today(), key="jour_caisse")
    with j2:
        hier_info = cx.cloture_hier(jour_sel)
        st.metric(
            "Clôture de la veille",
            f"{hier_info['solde_cloture']:,.0f}".replace(",", " ") + " F",
            help=f"Veille du {jour_sel.strftime('%d/%m/%Y')}",
        )

    sd_exist = cx.solde_depart_du_jour(jour_sel)
    if sd_exist is not None:
        st.info(
            f"Solde de départ déjà enregistré pour le {jour_sel.strftime('%d/%m/%Y')} : "
            f"**{sd_exist:,.0f} FCFA**".replace(",", " ")
        )

    opt = st.radio(
        "Mode d’ouverture",
        ["report_hier", "manuel"],
        horizontal=True,
        format_func=lambda x: (
            "↩️ Continuer sur le solde d’hier (recommandé)"
            if x == "report_hier"
            else "✏️ Saisir un solde de départ manuel"
        ),
        key="mode_ouv",
    )

    with st.form("form_solde_depart"):
        if opt == "report_hier":
            montant_sd = st.number_input(
                "Solde de départ (report hier)",
                min_value=0.0,
                value=float(max(hier_info["solde_cloture"], 0.0)),
                step=500.0,
                format="%.0f",
            )
            mode_saisie = "REPORT_HIER"
            notes_sd = st.text_input("Notes", value="Report automatique de la clôture de la veille")
        else:
            montant_sd = st.number_input(
                "Solde de départ manuel (FCFA)",
                min_value=0.0,
                value=float(sd_exist or 0.0),
                step=500.0,
                format="%.0f",
            )
            mode_saisie = "MANUEL"
            notes_sd = st.text_input("Notes", value="")
        ok_sd = st.form_submit_button("💾 Enregistrer le solde de départ", type="primary")

    if ok_sd:
        ok, msg = cx.enregistrer_solde_depart(
            jour_sel, montant_sd, mode=mode_saisie, notes=notes_sd, user=name
        )
        st.success(msg) if ok else st.error(msg)
        if ok:
            st.rerun()

    st.markdown("---")
    st.subheader("Fermeture de caisse du jour")
    # solde départ effectif pour calcul
    sd_calc = cx.solde_depart_du_jour(jour_sel)
    if sd_calc is None:
        st.warning(
            "Aucun solde de départ enregistré pour ce jour. "
            "Enregistrez-le ci-dessus avant d’imprimer la fermeture "
            "(sinon le calcul part de 0)."
        )
        sd_calc = 0.0

    cloture = cx.calculer_cloture(jour_sel, solde_depart=sd_calc)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Solde départ", f"{cloture['solde_depart']:,.0f}".replace(",", " ") + " F")
    c2.metric("Entrées", f"{cloture['entrees']:,.0f}".replace(",", " ") + " F")
    c3.metric("Dépenses", f"{cloture['depenses']:,.0f}".replace(",", " ") + " F")
    c4.metric("Solde clôture", f"{cloture['solde_cloture']:,.0f}".replace(",", " ") + " F")

    mvj = cloture["mouvements"]
    if mvj is not None and not mvj.empty:
        affj = mvj.copy()
        affj["Date"] = pd.to_datetime(affj["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(
            affj[["Date", "Type", "Categorie", "Libelle", "Tiers", "Montant", "Mode"]],
            hide_index=True, width="stretch", height=280,
        )
    else:
        st.caption("Aucun mouvement ce jour.")

    infos_f = {
        "Jour": jour_sel.strftime("%d/%m/%Y"),
        "Mode ouverture": "Report hier" if (sd_exist is not None) else "—",
        "Édité le": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Par": name,
    }
    try:
        pdf_f = generer_pdf_fermeture(cloture, infos_f)
        st.download_button(
            "🖨️ Imprimer le rapport de fermeture de caisse (PDF)",
            data=pdf_f,
            file_name=f"Fermeture_Caisse_{jour_sel:%Y%m%d}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )
    except Exception as ex:
        st.error(f"PDF fermeture : {ex}")


# ===================== LISTE & FILTRES =====================
with tab_liste:
    st.subheader("Mouvements — filtre par période & rapports")
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        debut = st.date_input("Du", value=date(date.today().year, 1, 1), key="caisse_d1")
        fin = st.date_input("Au", value=date.today(), key="caisse_d2")
    with f2:
        cats_f = ["Toutes"] + cx.liste_categories()
        cat_f = st.selectbox("Catégorie", cats_f, key="caisse_cat")
        type_f = st.selectbox("Type mouvement", ["Tous", "DEPENSE", "ENTREE"], key="caisse_typ")
        prod_f = st.selectbox(
            "Produit",
            ["Tous"] + list(cx.PRODUITS),
            key="caisse_prod",
            help="Graines de soja, Coton, Karité, Tourteaux…",
        )
    with f3:
        tt_f = st.selectbox("Type tiers", ["Tous"] + cx.TYPES_TIERS, key="caisse_tt")
        # liste tiers selon type
        if tt_f != "Tous":
            opts_tiers = ["Tous"] + cx.liste_tiers(tt_f)
        else:
            opts_tiers = ["Tous"] + sorted(set(cx.liste_tiers()) | set(
                df_all["Tiers"].dropna().astype(str) if not df_all.empty else []
            ))
            opts_tiers = [t for t in opts_tiers if t]
        tiers_f = st.selectbox("Client / Fournisseur", opts_tiers, key="caisse_tiers")
    with f4:
        mode_f = st.selectbox("Mode", ["Tous"] + cx.MODES, key="caisse_mode")
        st.write("")  # spacer

    if debut > fin:
        st.error("La date de début doit être ≤ date de fin.")
        st.stop()

    sous = cx.filtrer(
        df_all,
        debut=debut,
        fin=fin,
        categorie=cat_f,
        tiers=tiers_f,
        type_tiers=tt_f,
        type_mouv=type_f,
        mode=mode_f,
        produit=prod_f,
    )
    kpi = cx.indicateurs(sous)

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Lignes", kpi["nb"])
    k2.metric("Dépenses", f"{kpi['depenses']:,.0f}".replace(",", " "))
    k3.metric("Entrées", f"{kpi['entrees']:,.0f}".replace(",", " "))
    k4.metric("Solde période", f"{kpi['solde']:,.0f}".replace(",", " "))
    k5.metric("Nb dépenses", kpi["nb_dep"])

    if sous.empty:
        st.info("Aucun mouvement pour ces filtres.")
    else:
        aff = sous.copy()
        aff["Date"] = pd.to_datetime(aff["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
        cols = ["ID", "Date", "Type", "Categorie", "Produit", "Libelle", "Tiers", "Type_Tiers",
                "Montant", "Mode", "Num_Cheque", "Nom_Percepteur", "Reference", "Observation"]
        cols = [c for c in cols if c in aff.columns]
        st.dataframe(
            aff[cols],
            hide_index=True,
            width="stretch",
            height=420,
            column_config={
                "Montant": st.column_config.NumberColumn("Montant", format="%d"),
                "Type_Tiers": "Type tiers",
            },
        )
        st.markdown("##### Impression / export du rapport")
        infos_rap = {
            "Période": f"{debut:%d/%m/%Y} → {fin:%d/%m/%Y}",
            "Catégorie": cat_f,
            "Produit": prod_f,
            "Tiers": tiers_f,
            "Type": type_f,
            "Édité le": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Par": name,
        }
        e1, e2 = st.columns(2)
        with e1:
            st.download_button(
                "⬇️ Excel (filtre actif)",
                data=cx.to_excel_bytes(sous),
                file_name=f"Caisse_NAP_{debut:%Y%m%d}_{fin:%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        with e2:
            try:
                pdf_m = generer_pdf_mouvements(sous, infos_rap, kpi)
                st.download_button(
                    "🖨️ Imprimer PDF (mouvements)",
                    data=pdf_m,
                    file_name=f"Caisse_Mouvements_{debut:%Y%m%d}_{fin:%Y%m%d}.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True,
                )
            except Exception as ex:
                st.error(f"PDF mouvements : {ex}")

        with st.expander("Supprimer un mouvement"):
            ids = aff["ID"].tolist()
            if ids:
                sid = st.selectbox("ID à supprimer", ids)
                if st.button("🗑️ Supprimer", type="secondary"):
                    ok, msg = cx.supprimer_mouvement(sid)
                    st.success(msg) if ok else st.error(msg)
                    if ok:
                        st.rerun()

    # Mini synthèse catégories sur filtre
    st.markdown("#### Synthèse par catégorie (filtre actif)")
    syn_cat = cx.situation_par_categorie(sous)
    if syn_cat.empty:
        st.caption("Pas de données.")
    else:
        st.dataframe(syn_cat, hide_index=True, width="stretch")

# ===================== COMPTES =====================
with tab_comptes:
    st.subheader("Situation par client / fournisseur")
    st.caption(
        "Chaque tiers a son « compte » : total des dépenses et entrées liées. "
        "Filtrez pour voir la situation d’un seul compte."
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        d_c1 = st.date_input("Du", value=date(date.today().year, 1, 1), key="cpt_d1")
        d_c2 = st.date_input("Au", value=date.today(), key="cpt_d2")
    with c2:
        tt_c = st.selectbox("Type de compte", ["Tous", "Client", "Fournisseur", "Interne", "Autre"], key="cpt_tt")
    with c3:
        base_cpt = cx.filtrer(df_all, debut=d_c1, fin=d_c2, type_tiers=tt_c if tt_c != "Tous" else None)
        sit = cx.situation_par_tiers(base_cpt)
        noms = ["Tous"] + (sorted(sit["Tiers"].unique().tolist()) if not sit.empty else [])
        tiers_c = st.selectbox("Compte (tiers)", noms, key="cpt_tiers")

    vue = cx.filtrer(
        df_all,
        debut=d_c1,
        fin=d_c2,
        type_tiers=tt_c if tt_c != "Tous" else None,
        tiers=tiers_c if tiers_c != "Tous" else None,
    )
    sit2 = cx.situation_par_tiers(vue)

    if sit2.empty:
        st.info("Aucun mouvement lié à un tiers sur cette période.")
    else:
        st.dataframe(
            sit2,
            hide_index=True,
            width="stretch",
            column_config={
                "Total_Depenses": st.column_config.NumberColumn("Total dépenses", format="%d"),
                "Total_Entrees": st.column_config.NumberColumn("Total entrées", format="%d"),
                "Solde": st.column_config.NumberColumn("Solde (E−D)", format="%d"),
                "Type_Tiers": "Type",
            },
        )
        infos_all = {
            "Période": f"{d_c1:%d/%m/%Y} → {d_c2:%d/%m/%Y}",
            "Type compte": tt_c,
            "Édité le": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Par": name,
        }
        kpi_all_c = cx.indicateurs(vue)
        try:
            pdf_all = generer_pdf_comptes(sit2, None, infos_all, kpi_all_c)
            st.download_button(
                "🖨️ Imprimer PDF — synthèse des comptes",
                data=pdf_all,
                file_name=f"Caisse_Comptes_{d_c1:%Y%m%d}_{d_c2:%Y%m%d}.pdf",
                mime="application/pdf",
                type="primary",
            )
        except Exception as ex:
            st.error(f"PDF comptes : {ex}")

    if tiers_c != "Tous":
        st.markdown(f"#### Détail du compte : **{tiers_c}**")
        det = cx.filtrer(df_all, debut=d_c1, fin=d_c2, tiers=tiers_c)
        kpi_t = cx.indicateurs(det)
        a, b, c = st.columns(3)
        a.metric("Dépenses", f"{kpi_t['depenses']:,.0f}".replace(",", " ") + " F")
        b.metric("Entrées", f"{kpi_t['entrees']:,.0f}".replace(",", " ") + " F")
        c.metric("Solde", f"{kpi_t['solde']:,.0f}".replace(",", " ") + " F")
        if not det.empty:
            aff = det.copy()
            aff["Date"] = pd.to_datetime(aff["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
            st.dataframe(
                aff[["Date", "Type", "Categorie", "Libelle", "Montant", "Mode", "Reference"]],
                hide_index=True,
                width="stretch",
                height=320,
            )
            st.markdown("##### Impression du compte")
            infos_c = {
                "Période": f"{d_c1:%d/%m/%Y} → {d_c2:%d/%m/%Y}",
                "Compte": tiers_c,
                "Édité le": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "Par": name,
            }
            x1, x2 = st.columns(2)
            with x1:
                st.download_button(
                    f"⬇️ Excel — {tiers_c}",
                    data=cx.to_excel_bytes(det),
                    file_name=f"Compte_{tiers_c.replace(' ', '_')}_{d_c1:%Y%m%d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
            with x2:
                try:
                    pdf_c = generer_pdf_comptes(None, det, infos_c, kpi_t, titre_compte=tiers_c)
                    st.download_button(
                        f"🖨️ PDF — {tiers_c}",
                        data=pdf_c,
                        file_name=f"Compte_{tiers_c.replace(' ', '_')}_{d_c1:%Y%m%d}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True,
                    )
                except Exception as ex:
                    st.error(f"PDF compte : {ex}")

# ===================== CATEGORIES & TIERS =====================
with tab_cats:
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Catégories")
        with st.form("form_cat"):
            nc = st.text_input("Nouvelle catégorie")
            notes_c = st.text_input("Notes", key="notes_cat")
            if st.form_submit_button("Ajouter catégorie"):
                ok, msg = cx.ajouter_categorie(nc, notes_c)
                st.success(msg) if ok else st.error(msg)
                if ok:
                    st.rerun()
        cats_df = cx.charger_categories()
        if cats_df.empty:
            st.caption("Catégories par défaut à la première saisie.")
        else:
            st.dataframe(cats_df, hide_index=True, width="stretch")

    with col_b:
        st.subheader("Clients / Fournisseurs")
        with st.form("form_tiers"):
            nt = st.text_input("Nom")
            tt = st.selectbox("Type", cx.TYPES_TIERS)
            tel = st.text_input("Téléphone")
            if st.form_submit_button("Ajouter tiers"):
                ok, msg = cx.ajouter_tiers(nt, tt, tel)
                st.success(msg) if ok else st.error(msg)
                if ok:
                    st.rerun()
        tiers_df = cx.charger_tiers(actifs_only=False)
        if tiers_df.empty:
            st.caption("Aucun tiers enregistré.")
        else:
            st.dataframe(tiers_df, hide_index=True, width="stretch")

st.caption("STE NAP SARL — Caisse · dépenses · catégories · comptes client / fournisseur")
