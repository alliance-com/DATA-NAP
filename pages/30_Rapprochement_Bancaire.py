# -*- coding: utf-8 -*-
"""
NAP SARL — Rapprochement bancaire

Flux voulu :
1) Suivi de paiement (DEPENSES) → alimente le fichier STATEMENT local de la banque
2) Import du vrai relevé PDF (ou Excel) de la banque
3) Comparaison STATEMENT (chez nous) ↔ relevé banque (chez eux)
"""
import streamlit as st
import pandas as pd
from datetime import date
import os

import bank_reconciliation_data as br

st.set_page_config(page_title="Rapprochement bancaire", page_icon="🏦", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

st.title("🏦 Rapprochement bancaire")
st.markdown(
    """
**Logique (3 étapes)**  
1. **Suivi de paiement** → on **remplit** le fichier `… STATEMENT SHEET …xlsx` de la banque  
2. **Relevé PDF/Excel de la banque** → on **importe** les mouvements réels banque  
3. **Rapprochement** → on **compare** les deux pour voir les écarts  
"""
)

st.sidebar.header("Filtres")
banque = st.sidebar.selectbox("Banque *", br.BANQUES, index=0)
today = date.today()
periode = st.sidebar.date_input(
    "Période *",
    value=(date(today.year, today.month, 1), today),
    help="Période pour alimenter le STATEMENT et pour rapprocher.",
)
if isinstance(periode, (tuple, list)) and len(periode) == 2:
    d_deb, d_fin = periode[0], periode[1]
else:
    d_deb, d_fin = date(today.year, 1, 1), today

tol_j = st.sidebar.slider("Tolérance date (jours)", 0, 10, 3)
tol_m = st.sidebar.number_input("Tolérance montant (FCFA)", 0, 50000, 1)

ocr = br.statut_ocr()
with st.sidebar.expander("État OCR"):
    st.write("Tesseract:", ocr["tesseract_info"])
    st.write("Poppler:", ocr["poppler_path"])

# --- Données ---
# Journal complet = débits Suivi + crédits base crédits
journal = br.charger_journal_complet(banque, date_debut=d_deb, date_fin=d_fin)
journal_debits_only = br.charger_journal(banque, date_debut=d_deb, date_fin=d_fin)
statement, path_statement = br.charger_releve_local(banque, date_debut=d_deb, date_fin=d_fin)
n_cred = int((journal["Sens"] == "CREDIT").sum()) if (journal is not None and not journal.empty and "Sens" in journal.columns) else 0
n_deb = int((journal["Sens"] == "DEBIT").sum()) if (journal is not None and not journal.empty and "Sens" in journal.columns) else len(journal_debits_only)

# Relevé banque importé (PDF) — côté "chez la banque"
df_pdf_saved, path_pdf_saved = br.charger_releve_import(banque)
if df_pdf_saved is not None and not df_pdf_saved.empty:
    releve_banque = br.filtrer_periode(df_pdf_saved, d_deb, d_fin)
    src_banque = path_pdf_saved or "Import PDF/Excel"
else:
    releve_banque = pd.DataFrame()
    src_banque = "— (importer le relevé banque)"

c1, c2, c3, c4 = st.columns(4)
c1.metric("Banque", banque)
c2.metric("Chez nous (période)", len(journal), f"{n_deb} débits · {n_cred} crédits")
c3.metric("STATEMENT local (période)", len(statement) if statement is not None else 0)
c4.metric("Relevé banque importé (période)", len(releve_banque) if releve_banque is not None else 0)

tab1, tab2, tab3, tab4 = st.tabs([
    "① Suivi → STATEMENT",
    "② Import relevé banque (PDF)",
    "③ Rapprochement",
    "✏️ Écriture manuelle",
])

# ========== TAB 1 : alimenter STATEMENT depuis Suivi ==========
with tab1:
    st.subheader("Étape 1 — Remplir le STATEMENT local depuis le Suivi de paiement")
    st.info(
        f"Les **débits** du Suivi (DEPENSES) filtrés **{banque}** + période "
        f"seront écrits dans le STATEMENT local. "
        f"Les **crédits** viennent de Credits_Bancaires_NAP.xlsx et sont inclus automatiquement dans le rapprochement."
    )
    st.caption(f"Période : {d_deb.strftime('%d/%m/%Y')} → {d_fin.strftime('%d/%m/%Y')}")

    if journal.empty:
        st.warning("Aucune ligne dans le Suivi pour cette banque / période.")
    else:
        st.write(f"**{len(journal)}** ligne(s) Suivi prêtes à être poussées vers STATEMENT.")
        aff = journal.copy()
        aff["Date"] = pd.to_datetime(aff["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
        cols = [c for c in ["Date", "No_Transaction", "Prestataire", "Description", "Montant", "Num_Cheque", "Mode"] if c in aff.columns]
        st.dataframe(aff[cols], hide_index=True, width="stretch", height=220)

    if st.button("📥 Alimenter le STATEMENT local depuis le Suivi", type="primary"):
        ok, msg, df_out, path = br.alimenter_statement_depuis_suivi(banque, d_deb, d_fin)
        if ok:
            st.success(msg)
            st.session_state.pop("match_result", None)
            st.rerun()
        else:
            st.error(msg)

    if path_statement:
        st.caption(f"Fichier STATEMENT actuel : `{path_statement}`")
        if statement is not None and not statement.empty:
            with st.expander("Aperçu STATEMENT (période)"):
                a = statement.copy()
                a["Date"] = pd.to_datetime(a["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
                st.dataframe(a, hide_index=True, width="stretch", height=200)

# ========== TAB 2 : import PDF banque ==========
with tab2:
    st.subheader("Étape 2 — Importer le **vrai** relevé de la banque (PDF ou Excel banque)")
    st.markdown(
        "Ce n’est **pas** le Suivi de paiement. C’est le document fourni par ECOBANK / ATLANTIC / BIIC."
    )
    if path_pdf_saved:
        st.success(f"Dernier import banque enregistré : `{path_pdf_saved}` ({len(df_pdf_saved)} lignes au total)")
    else:
        st.warning("Aucun relevé banque importé pour l’instant.")

    def _save_bank_import(df, label):
        if df is None or df.empty:
            st.warning("Aucune ligne reconnue.")
            return False
        br.sauver_releve_import(df, banque)
        st.session_state["releve_import"] = df
        st.session_state["releve_import_banque"] = banque
        st.session_state["show_ops_banque"] = True
        st.session_state.pop("match_result", None)
        dmin = pd.to_datetime(df["Date"], errors="coerce").min()
        dmax = pd.to_datetime(df["Date"], errors="coerce").max()
        st.success(
            f"**{len(df)}** mouvements banque enregistrés ({label}). "
            f"Dates : {dmin.strftime('%d/%m/%Y')} → {dmax.strftime('%d/%m/%Y')}"
        )
        return True

    def _afficher_ops_banque(df, key_suffix="main"):
        """Tableau complet de toutes les opérations banque importées."""
        if df is None or df.empty:
            st.info("Aucune opération banque à afficher.")
            return
        aff = df.copy()
        aff["Date"] = pd.to_datetime(aff["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
        preferred = [c for c in ["Date", "Description", "Credit", "Debit", "Solde", "Num_Cheque", "Motif", "Categorie", "Banque", "Source"] if c in aff.columns]
        other = [c for c in aff.columns if c not in preferred]
        aff = aff[preferred + other]
        total_c = float(pd.to_numeric(df["Credit"], errors="coerce").fillna(0).sum()) if "Credit" in df.columns else 0.0
        total_d = float(pd.to_numeric(df["Debit"], errors="coerce").fillna(0).sum()) if "Debit" in df.columns else 0.0
        m1, m2, m3 = st.columns(3)
        m1.metric("Lignes", len(df))
        m2.metric("Total crédits", f"{total_c:,.0f} F".replace(",", " "))
        m3.metric("Total débits", f"{total_d:,.0f} F".replace(",", " "))
        st.dataframe(aff, hide_index=True, width="stretch", height=420)
        from io import BytesIO
        buf = BytesIO()
        df.to_excel(buf, index=False)
        st.download_button(
            "⬇️ Télécharger ces opérations (Excel)",
            data=buf.getvalue(),
            file_name=f"Operations_banque_{banque}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_ops_banque_{key_suffix}",
        )

    up_pdf = st.file_uploader("📄 PDF relevé banque", type=["pdf"], key="pdf_bank")
    if up_pdf and st.button("🔍 Analyser et enregistrer le PDF banque", type="primary"):
        with st.spinner("OCR / lecture PDF…"):
            try:
                df_pdf, texte = br.lire_releve_pdf(up_pdf.read(), banque=banque)
                if df_pdf.empty:
                    st.error("Aucune ligne lue. Ouvrez le debug ou collez le texte.")
                    with st.expander("Texte extrait"):
                        st.text((texte or "")[:5000])
                else:
                    _save_bank_import(df_pdf, "PDF")
                    st.session_state["last_import_df"] = df_pdf
                    # afficher tout de suite sans perdre le tableau
                    st.markdown("#### Toutes les opérations extraites du PDF")
                    _afficher_ops_banque(df_pdf, "after_pdf")
            except Exception as e:
                st.error(str(e))
                import traceback
                st.code(traceback.format_exc())

    up_xls = st.file_uploader("📊 Excel relevé banque (si la banque vous envoie un xlsx)", type=["xlsx", "xls"], key="xls_bank")
    if up_xls and st.button("Importer Excel banque"):
        tmp = f"_tmp_bank_{banque}.xlsx"
        with open(tmp, "wb") as f:
            f.write(up_xls.read())
        try:
            df_i = br.lire_releve_excel(tmp, banque=banque)
            if _save_bank_import(df_i, "Excel banque"):
                st.session_state["last_import_df"] = df_i
                st.markdown("#### Toutes les opérations importées")
                _afficher_ops_banque(df_i, "after_xls")
        except Exception as e:
            st.error(str(e))
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)

    st.markdown("##### Coller le texte du relevé banque")
    txt = st.text_area("Texte", height=100, key="txt_bank")
    if st.button("Enregistrer depuis texte"):
        df_t = br.parser_texte_releve(txt, banque=banque)
        if _save_bank_import(df_t, "texte"):
            st.session_state["last_import_df"] = df_t
            st.markdown("#### Toutes les opérations importées")
            _afficher_ops_banque(df_t, "after_txt")

    st.markdown("---")
    st.markdown("#### Opérations banque déjà enregistrées")
    col_b1, col_b2 = st.columns([2, 1])
    with col_b1:
        btn_show = st.button("📋 Afficher toutes les opérations banque", type="primary", key="btn_show_ops")
    with col_b2:
        btn_clear = st.button("🗑️ Effacer l’import relevé banque", key="btn_clear_imp")

    if btn_clear:
        br.effacer_releve_import(banque)
        st.session_state.pop("last_import_df", None)
        st.session_state.pop("releve_import", None)
        st.session_state["show_ops_banque"] = False
        st.success("Import relevé banque effacé.")
        st.rerun()

    # Source d'affichage : dernier import session > fichier disque
    df_show = None
    if st.session_state.get("last_import_df") is not None and st.session_state.get("releve_import_banque") == banque:
        df_show = st.session_state["last_import_df"]
    elif df_pdf_saved is not None and not df_pdf_saved.empty:
        df_show = df_pdf_saved
    elif st.session_state.get("releve_import") is not None and st.session_state.get("releve_import_banque") == banque:
        df_show = st.session_state["releve_import"]

    if btn_show or st.session_state.get("show_ops_banque"):
        if df_show is not None and not df_show.empty:
            st.session_state["show_ops_banque"] = True
            _afficher_ops_banque(df_show, "saved")
        else:
            st.warning("Aucune opération enregistrée pour cette banque. Analysez d’abord un PDF/Excel.")

# ========== TAB 3 : rapprochement ==========
with tab3:
    st.subheader("Étape 3 — Comparer STATEMENT (chez nous) ↔ relevé banque (chez eux)")
    st.caption(
        f"STATEMENT local : {path_statement or '—'} · "
        f"Relevé banque : {src_banque}"
    )

    # Chez nous = débits (Suivi) + crédits (base crédits) — déjà dans journal
    use_journal = journal
    use_releve = releve_banque
    st.caption(
        f"Comparaison : **{n_deb} débits + {n_cred} crédits** (chez nous) ↔ "
        f"**{len(use_releve) if use_releve is not None else 0} lignes** relevé banque."
    )

    if (use_releve is None or use_releve.empty) and (use_journal is None or use_journal.empty):
        st.warning("Il faut le Suivi/crédits (chez nous) et le relevé banque importé (étape 2).")
    else:
        if st.button("▶️ Lancer le rapprochement", type="primary"):
            st.session_state["match_result"] = br.rapprocher_complet(
                use_releve if use_releve is not None else pd.DataFrame(),
                use_journal if use_journal is not None else pd.DataFrame(),
                tolerance_jours=int(tol_j),
                tolerance_montant=float(tol_m),
            )

        res = st.session_state.get("match_result")
        if res is None:
            st.caption("Cliquez sur **Lancer le rapprochement**.")
        else:
            ok, banque_seul, journal_seul = res
            m1, m2, m3 = st.columns(3)
            m1.metric("✅ Rapprochés", len(ok))
            m2.metric("🏦 Sur le relevé banque seulement", len(banque_seul))
            m3.metric("📓 Chez nous seulement", len(journal_seul))

            if not banque_seul.empty:
                st.error("Mouvements **à la banque** absents chez nous (ni débit Suivi, ni crédit) — à justifier ou saisir.")
                a = banque_seul.copy()
                a["Date"] = pd.to_datetime(a["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
                st.dataframe(a, hide_index=True, width="stretch", height=260)
            if not journal_seul.empty:
                st.warning("Mouvements **chez nous** (débit ou crédit) absents du relevé banque (période).")
                a = journal_seul.copy()
                a["Date"] = pd.to_datetime(a["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
                st.dataframe(a, hide_index=True, width="stretch", height=260)
            if banque_seul.empty and journal_seul.empty and not ok.empty:
                st.success("Écarts nuls sur la période.")

            with st.expander("Détail rapproché"):
                if not ok.empty:
                    a = ok.copy()
                    for c in ("Date_Releve", "Date_Journal"):
                        if c in a.columns:
                            a[c] = pd.to_datetime(a[c], errors="coerce").dt.strftime("%d/%m/%Y")
                    st.dataframe(a, hide_index=True, width="stretch")

            st.markdown("#### Rapport & impression")
            cexp1, cexp2 = st.columns(2)
            with cexp1:
                st.download_button(
                    "⬇️ Export Excel du rapprochement",
                    data=br.to_excel_bytes({"Rapproches": ok, "Banque_seul": banque_seul, "Suivi_seul": journal_seul}),
                    file_name=f"Rapprochement_{banque}_{d_deb}_{d_fin}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_rappro_xlsx",
                )
            with cexp2:
                try:
                    pdf_bytes = br.generer_pdf_rapprochement_a3(
                        banque, d_deb, d_fin, ok, banque_seul, journal_seul
                    )
                    st.download_button(
                        "🖨️ Rapport PDF A3 (impression)",
                        data=pdf_bytes,
                        file_name=f"Rapprochement_{banque}_{d_deb}_{d_fin}_A3.pdf",
                        mime="application/pdf",
                        type="primary",
                        key="dl_rappro_a3",
                    )
                except Exception as e:
                    st.error(f"PDF A3 indisponible : {e}")
                    st.caption("Vérifiez : pip install reportlab")

# ========== TAB 4 manuel ==========
with tab4:
    st.subheader("Écriture manuelle (frais, régularisations)")
    with st.form("man"):
        d1, d2 = st.columns(2)
        with d1:
            e_date = st.date_input("Date", value=d_fin)
            e_desc = st.text_input("Description *")
            e_chq = st.text_input("N° chèque")
        with d2:
            e_credit = st.number_input("Crédit", min_value=0.0, step=1000.0)
            e_debit = st.number_input("Débit", min_value=0.0, step=1000.0)
            e_motif = st.text_input("Motif", value="FRAIS BANCAIRE")
        if st.form_submit_button("Enregistrer"):
            ok, msg = br.ajouter_manuel({
                "Banque": banque, "Date": e_date, "Description": e_desc,
                "Credit": e_credit, "Debit": e_debit, "Num_Cheque": e_chq,
                "Motif": e_motif, "Categorie": e_motif,
            }, user=name)
            st.success(msg) if ok else st.error(msg)

st.markdown("---")
st.caption(
    "NAP SARL — ① Suivi → STATEMENT local · ② PDF banque → import · ③ Comparaison des écarts"
)
