# -*- coding: utf-8 -*-
"""
NAP SARL — Situation & soldes des Transitaires
Croise dossiers export + paiements (Suivi de paiement.xlsx, Type EXPORT / FRET).
Les paiements se font via le formulaire existant de la page Paiements.
"""
import streamlit as st
import pandas as pd
import os
import importlib

import export_dossiers_data as ed
import transitaires_data as td
importlib.reload(ed)
importlib.reload(td)

st.set_page_config(page_title="Soldes Transitaires", page_icon="💰", layout="wide")

if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"
FEUILLE_DEPENSES = "DEPENSES"


def charger_journal_export():
    """Charge les paiements liés à l'export / transitaires depuis Suivi de paiement.xlsx."""
    if not os.path.exists(FICHIER_PAIEMENTS):
        return pd.DataFrame()
    try:
        xl = pd.ExcelFile(FICHIER_PAIEMENTS)
        sheet = FEUILLE_DEPENSES if FEUILLE_DEPENSES in xl.sheet_names else xl.sheet_names[0]
        df = None
        for hdr in (2, 1, 0, 3):
            try:
                tmp = pd.read_excel(FICHIER_PAIEMENTS, sheet_name=sheet, header=hdr)
                cols = " ".join(str(c).lower() for c in tmp.columns)
                if "prestataire" in cols or "transaction" in cols or "montant" in cols:
                    df = tmp
                    break
            except Exception:
                continue
        if df is None:
            return pd.DataFrame()
        # normaliser colonnes clés
        rename = {}
        for c in df.columns:
            k = str(c).strip().lower()
            if "prestataire" in k or "fournisseur" in k:
                rename[c] = "Prestataire"
            elif k in ("montant",):
                rename[c] = "Montant"
            elif "type" in k and "op" in k:
                rename[c] = "Type"
            elif k == "produit":
                rename[c] = "Produit"
            elif "no de transaction" in k or "n° de transaction" in k or "transaction" == k:
                rename[c] = "No_Transaction"
            elif k == "date":
                rename[c] = "Date"
            elif "description" in k or "objet" in k:
                rename[c] = "Description"
            elif "observation" in k:
                rename[c] = "Observation"
        df = df.rename(columns=rename)
        for c in ["Prestataire", "Montant", "Type", "Produit", "No_Transaction", "Date", "Description", "Observation"]:
            if c not in df.columns:
                df[c] = ""
        df["Montant"] = pd.to_numeric(df["Montant"], errors="coerce").fillna(0)
        df["Prestataire"] = df["Prestataire"].astype(str).str.strip()
        # Filtrer export / fret si possible, sinon tout (rapprochement par nom transitaire)
        mask_export = (
            df["Type"].astype(str).str.upper().str.contains("EXPORT", na=False)
            | df["Produit"].astype(str).str.upper().str.contains("FRET|EXPORT", na=False, regex=True)
        )
        # on garde aussi les lignes dont le prestataire est un transitaire connu
        return df
    except Exception as e:
        st.error(f"Erreur lecture paiements : {e}")
        return pd.DataFrame()


st.title("💰 Situation & soldes Transitaires — STE NAP SARL")
st.caption(
    "Montant des dossiers export − paiements enregistrés dans **Paiements** "
    "(Type = EXPORT ou Produit = FRET EXPORT). "
    "Utilisez le formulaire existant de la page **Paiements** pour payer un transitaire."
)

dossiers = ed.charger_dossiers()
journal = charger_journal_export()
trs = td.charger_transitaires(actifs_seulement=False)

# ---------- Situation globale ----------
situation = ed.situation_par_transitaire(dossiers, journal, tous_transitaires=trs)

if trs.empty and dossiers.empty:
    st.info("Aucun transitaire et aucun dossier. "
            "Créez d'abord des **Transitaires**, puis des **dossiers export**.")
elif situation.empty:
    st.info("Aucune situation à afficher. Vérifiez le fichier `transitaires.xlsx` à côté de l'application.")
else:
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Transitaires", len(situation))
    k2.metric("Montant dossiers", ed.fmt(situation["Montant_Total"].sum(), "FCFA"))
    k3.metric("Total payé", ed.fmt(situation["Montant_Paye"].sum(), "FCFA"))
    k4.metric("Solde global", ed.fmt(situation["Solde"].sum(), "FCFA"))

    st.markdown("---")
    st.subheader("📊 Situation par transitaire")
    aff = situation.copy()
    st.dataframe(
        aff,
        hide_index=True, width="stretch",
        column_config={
            "Code_Transitaire": "Code",
            "Transitaire": "Transitaire",
            "Nb_Dossiers": "Nb dossiers",
            "Montant_Total": st.column_config.NumberColumn("Montant dossiers", format="%d"),
            "Montant_Paye": st.column_config.NumberColumn("Payé", format="%d"),
            "Solde": st.column_config.NumberColumn("Solde", format="%d"),
            "Nb_En_Cours": "En cours",
            "Nb_Charge": "Chargé",
            "Nb_Embarque": "Embarqué",
        }
    )
    st.download_button(
        "⬇️ Télécharger situation Excel",
        data=ed.to_excel_bytes({"Situation_Transitaires": situation, "Dossiers": dossiers}),
        file_name="situation_transitaires_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------- Détail d'un transitaire ----------
st.markdown("---")
st.subheader("🔎 Détail d'un transitaire")
noms = sorted(situation["Transitaire"].tolist()) if not situation.empty else []
if not noms and not trs.empty:
    noms = sorted(trs["Nom"].tolist())

choix = st.selectbox("Transitaire", [""] + noms, format_func=lambda x: "— Choisir —" if x == "" else x)
if choix:
    sit = situation[situation["Transitaire"] == choix] if not situation.empty else pd.DataFrame()
    if not sit.empty:
        r = sit.iloc[0]
        c1, c2, c3 = st.columns(3)
        c1.metric("Montant dossiers", ed.fmt(r["Montant_Total"], "FCFA"))
        c2.metric("Payé", ed.fmt(r["Montant_Paye"], "FCFA"))
        c3.metric("Solde", ed.fmt(r["Solde"], "FCFA"))

    # Dossiers de ce transitaire
    dos = dossiers[dossiers["Transitaire"] == choix] if not dossiers.empty else pd.DataFrame()
    st.markdown("**Dossiers export**")
    if dos.empty:
        st.caption("Aucun dossier.")
    else:
        st.dataframe(
            dos[["ID", "Date", "Contrat", "Produit", "Numero_Booking", "Numero_BL",
                 "Conteneur", "Montant", "Statut"]].sort_values("ID", ascending=False),
            hide_index=True, width="stretch",
            column_config={
                "Date": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
                "Montant": st.column_config.NumberColumn("Montant", format="%d"),
                "Numero_Booking": "Booking",
                "Numero_BL": "BL",
            }
        )

    # Paiements liés (journal)
    st.markdown("**Paiements enregistrés (page Paiements)**")
    if journal.empty:
        st.caption("Aucun paiement dans Suivi de paiement.xlsx — "
                   "enregistrez un paiement avec Type=EXPORT et le nom du transitaire comme prestataire.")
    else:
        pay = journal[journal["Prestataire"].str.casefold().str.contains(choix.casefold(), na=False)]
        # aussi filtre type export
        if pay.empty:
            pay = journal[
                journal["Prestataire"].map(lambda p: choix.casefold() in str(p).casefold() or str(p).casefold() in choix.casefold())
            ]
        if pay.empty:
            st.caption(f"Aucun paiement trouvé pour « {choix} ». "
                       "Dans **Paiements**, créez une ligne avec Prestataire = ce transitaire, Type = EXPORT.")
        else:
            cols_aff = [c for c in ["Date", "No_Transaction", "Type", "Produit", "Prestataire",
                                     "Description", "Montant", "Observation"] if c in pay.columns]
            st.dataframe(pay[cols_aff], hide_index=True, width="stretch",
                         column_config={"Montant": st.column_config.NumberColumn("Montant", format="%d")})
            st.caption(f"Total payé affiché : {ed.fmt(pay['Montant'].sum(), 'FCFA')}")

st.markdown("---")
st.info(
    "**Comment payer un transitaire ?**  \n"
    "1. Allez sur la page **Paiements** (formulaire existant)  \n"
    "2. Type d'opération = **EXPORT**  \n"
    "3. Produit = **FRET EXPORT** (ou le produit concerné)  \n"
    "4. Prestataire = **nom exact du transitaire**  \n"
    "5. Description / Réf. Doc = N° BL ou Booking si possible  \n"
    "Le solde se met à jour automatiquement ici."
)

st.caption("NAP SARL — Soldes Transitaires · données locales Excel exploitables")
