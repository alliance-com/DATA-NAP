# -*- coding: utf-8 -*-
"""
NAP SARL — Impression A3 (PDF)
Listes : déchargements, bons de commande, paiements, camions non déchargés, synthèse fournisseurs.
"""
import streamlit as st
import pandas as pd
from datetime import date, datetime
import os
import re

st.set_page_config(page_title="Impression A3", page_icon="🖨️", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

import dechargement_data as dd
from impression_pdf import generer_pdf_liste, df_to_rows, fmt_num

FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"
os.makedirs("impressions_pdf", exist_ok=True)

# ---------- Produit ----------
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

# ---------- Langue ----------
lang = st.sidebar.radio(
    "Language / Langue",
    ["Français", "English"],
    horizontal=True,
    help="The generated A3 PDF will use this language for titles and column headers.",
)
EN = lang == "English"

def T(fr, en):
    return en if EN else fr

st.title(T("🖨️ Impression A3 — listes professionnelles", "🖨️ A3 Print — professional lists"))
st.caption(
    T(
        "Format **A3 paysage** · PDF prêt à imprimer · toutes les colonnes utiles · pied de page NAP SARL",
        "Format **A3 landscape** · print-ready PDF · all useful columns · NAP SARL footer",
    )
)

TYPE_OPTS = {
    T("Déchargements", "Unloadings"): "dech",
    T("Bons de commande", "Purchase orders"): "bons",
    T("Paiements (Suivi de paiement)", "Payments (payment tracker)"): "paiements",
    T("Camions non déchargés (1 ligne = 1 camion)", "Trucks not unloaded (1 row = 1 truck)"): "camions",
    T("Synthèse fournisseurs (bons / livré / reste)", "Supplier summary (orders / delivered / balance)"): "synthese",
}

type_label = st.selectbox(T("Document à imprimer", "Document to print"), list(TYPE_OPTS.keys()))
type_code = TYPE_OPTS[type_label]

c1, c2, c3 = st.columns(3)
with c1:
    debut = st.date_input(T("Du", "From"), value=date(2025, 1, 1), min_value=date(2025, 1, 1))
with c2:
    fin = st.date_input(T("Au", "To"), value=date.today(), min_value=date(2025, 1, 1))
with c3:
    st.write("")
    st.write("")
    inclure_tout = st.checkbox(T("Ignorer les dates (tout exporter)", "Ignore dates (export all)"), value=False)

# ---------- Chargement selon type ----------
def charger_paiements():
    if not os.path.exists(FICHIER_PAIEMENTS):
        return pd.DataFrame()
    try:
        xl = pd.ExcelFile(FICHIER_PAIEMENTS)
        sheet = "DEPENSES" if "DEPENSES" in xl.sheet_names else xl.sheet_names[0]
        df = None
        for hdr in (2, 1, 0, 3):
            try:
                tmp = pd.read_excel(FICHIER_PAIEMENTS, sheet_name=sheet, header=hdr)
                cols = " ".join(str(c).lower() for c in tmp.columns)
                if "transaction" in cols or "prestataire" in cols:
                    df = tmp
                    break
            except Exception:
                continue
        if df is None:
            return pd.DataFrame()
        alias = {
            "date": "Date", "no de transaction": "No_transaction", "n° de transaction": "No_transaction",
            "type d'operation": "Type", "type d'opération": "Type", "produit": "Produit",
            "prestataires/fournisseurs": "Fournisseur", "prestataire": "Fournisseur", "fournisseur": "Fournisseur",
            "nom du bénéficiaire": "Beneficiaire", "ref doc": "Ref_doc",
            "description/objet du paiement": "Description", "description": "Description",
            "montant": "Montant", "observation": "Observation", "mode paiement": "Mode",
            "numero de cheque": "Num_cheque", "banque": "Banque", "n_bc": "N_BC",
        }
        rename = {}
        for c in df.columns:
            k = str(c).strip().lower().replace("  ", " ")
            if k in alias:
                rename[c] = alias[k]
        df = df.rename(columns=rename)
        for c in ["Date", "No_transaction", "Type", "Produit", "Fournisseur", "Beneficiaire",
                  "Ref_doc", "Description", "Montant", "Observation", "Mode", "Num_cheque", "Banque", "N_BC"]:
            if c not in df.columns:
                df[c] = ""
        df["Montant"] = pd.to_numeric(df["Montant"], errors="coerce").fillna(0)
        df["Date_dt"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)
        return df
    except Exception as e:
        st.error(str(e))
        return pd.DataFrame()


def filtre_dates(df, col_date):
    if inclure_tout or df.empty or col_date not in df.columns:
        return df
    s = pd.to_datetime(df[col_date], errors="coerce", dayfirst=True)
    mask = (s.dt.date >= debut) & (s.dt.date <= fin)
    return df.loc[mask.fillna(False)].copy()


df_dech, _ = dd.charger_complet()
bons = dd.charger_bons_commande()
suivi = dd.suivi_par_bc(df_dech, bons)

headers, rows, titre, sous, kpis = [], [], "", "", []

if type_code == "dech":
    d = filtre_dates(df_dech, "Date_Dechargement")
    cols = [
        "Transaction", "Date_Dechargement", "Ref_BC", "Fournisseur", "Fournisseur_BC",
        "Affectation_BC", "N_Camion", "Camion_BC", "Weighment_Ticket",
        "Poids_Net_KG", "Nb_Sacs", "PU_FCFA_KG", "Montant_Du", "Lieu_Stockage",
        "Statut_Livraison", "Statut_Paiement", "Remarque",
    ]
    cols = [c for c in cols if c in d.columns]
    labels = {
        "Transaction": T("N° Trans.", "Txn No."),
        "Date_Dechargement": T("Date", "Date"),
        "Ref_BC": T("N° BC", "PO No."),
        "Fournisseur": T("Fournisseur ticket", "Ticket supplier"),
        "Fournisseur_BC": T("Fournisseur BC", "PO supplier"),
        "Affectation_BC": T("Échange", "Swap"),
        "N_Camion": T("Camion livré", "Truck delivered"),
        "Camion_BC": T("Camion BC", "PO truck"),
        "Weighment_Ticket": T("Ticket", "Weigh ticket"),
        "Poids_Net_KG": T("Poids net", "Net weight"),
        "Nb_Sacs": T("Sacs", "Bags"),
        "PU_FCFA_KG": T("PU", "Unit price"),
        "Montant_Du": T("Montant dû", "Amount due"),
        "Lieu_Stockage": T("Stockage", "Storage"),
        "Statut_Livraison": T("Livraison", "Delivery"),
        "Statut_Paiement": T("Paiement", "Payment"),
        "Remarque": T("Remarque", "Remark"),
    }
    headers = [labels.get(c, c) for c in cols]
    fmt = {
        "Poids_Net_KG": lambda v: fmt_num(v),
        "PU_FCFA_KG": lambda v: fmt_num(v),
        "Montant_Du": lambda v: fmt_num(v),
        "Nb_Sacs": lambda v: fmt_num(v),
    }
    rows = df_to_rows(d, cols, fmt)
    titre = T(
        f"LISTE DES DÉCHARGEMENTS — {dd.libelle_produit(produit)}",
        f"UNLOADING LIST — {dd.libelle_produit(produit)}",
    )
    sous = (
        (T(f"Période : {debut.strftime('%d/%m/%Y')} → {fin.strftime('%d/%m/%Y')}",
           f"Period: {debut.strftime('%d/%m/%Y')} → {fin.strftime('%d/%m/%Y')}"))
        if not inclure_tout else T("Toutes périodes", "All periods")
    )
    kpis = [
        T(f"{len(d)} lignes", f"{len(d)} rows"),
        T(
            f"Poids {fmt_num(d['Poids_Net_KG'].sum()/1000 if not d.empty and 'Poids_Net_KG' in d.columns else 0, 1)} T",
            f"Weight {fmt_num(d['Poids_Net_KG'].sum()/1000 if not d.empty and 'Poids_Net_KG' in d.columns else 0, 1)} T",
        ),
        T(
            f"Dû {fmt_num(d['Montant_Du'].sum() if not d.empty and 'Montant_Du' in d.columns else 0)} FCFA",
            f"Due {fmt_num(d['Montant_Du'].sum() if not d.empty and 'Montant_Du' in d.columns else 0)} FCFA",
        ),
    ]
    apercu = d

elif type_code == "bons":
    d = bons.copy()
    if not inclure_tout and "Date_Emission" in d.columns and not d.empty:
        s = pd.to_datetime(d["Date_Emission"], errors="coerce", dayfirst=True)
        d = d.loc[((s.dt.date >= debut) & (s.dt.date <= fin)).fillna(False)].copy()
    cols = [c for c in [
        "Reference", "Date_Emission", "Nom_Fournisseur", "Code_Fournisseur", "Quantite",
        "Numero_Tracteur", "Chauffeur", "Produit",
    ] if c in d.columns]
    # enrich from full file if more cols
    labels = {
        "Reference": T("N° BC", "PO No."),
        "Date_Emission": T("Date", "Date"),
        "Nom_Fournisseur": T("Fournisseur", "Supplier"),
        "Code_Fournisseur": T("Code", "Code"),
        "Quantite": T("Quantité", "Quantity"),
        "Numero_Tracteur": T("Camion", "Truck"),
        "Chauffeur": T("Chauffeur", "Driver"),
        "Produit": T("Produit", "Product"),
    }
    headers = [labels.get(c, c) for c in cols]
    rows = df_to_rows(d, cols)
    titre = T(
        f"LISTE DES BONS DE COMMANDE — {dd.libelle_produit(produit)}",
        f"PURCHASE ORDER LIST — {dd.libelle_produit(produit)}",
    )
    sous = (
        T(f"Période : {debut.strftime('%d/%m/%Y')} → {fin.strftime('%d/%m/%Y')}",
          f"Period: {debut.strftime('%d/%m/%Y')} → {fin.strftime('%d/%m/%Y')}")
        if not inclure_tout else T("Toutes périodes", "All periods")
    )
    kpis = [T(f"{len(d)} bons", f"{len(d)} POs")]
    apercu = d

elif type_code == "paiements":
    d = charger_paiements()
    d = filtre_dates(d, "Date_dt")
    # filtre produit si colonne
    if "Produit" in d.columns and not d.empty:
        # optionnel : ne garder que le produit actif si renseigné
        pass
    cols = [
        "Date", "No_transaction", "Type", "Produit", "Fournisseur", "Beneficiaire",
        "Description", "Ref_doc", "Montant", "Mode", "Num_cheque", "Banque", "Observation", "N_BC",
    ]
    cols = [c for c in cols if c in d.columns]
    labels = {
        "Date": T("Date", "Date"),
        "No_transaction": T("N° transaction", "Txn No."),
        "Type": T("Type", "Type"),
        "Produit": T("Produit", "Product"),
        "Fournisseur": T("Fournisseur", "Supplier"),
        "Beneficiaire": T("Bénéficiaire", "Beneficiary"),
        "Description": T("Description", "Description"),
        "Ref_doc": T("Réf. doc", "Doc ref."),
        "Montant": T("Montant", "Amount"),
        "Mode": T("Mode", "Payment mode"),
        "Num_cheque": T("N° chèque", "Cheque No."),
        "Banque": T("Banque", "Bank"),
        "Observation": T("Observation", "Remark"),
        "N_BC": T("N° BC", "PO No."),
    }
    headers = [labels.get(c, c) for c in cols]
    fmt = {"Montant": lambda v: fmt_num(v)}
    rows = df_to_rows(d, cols, fmt)
    titre = T("LISTE DES PAIEMENTS — SUIVI NAP SARL", "PAYMENT LIST — NAP SARL TRACKER")
    sous = (
        T(f"Période : {debut.strftime('%d/%m/%Y')} → {fin.strftime('%d/%m/%Y')}",
          f"Period: {debut.strftime('%d/%m/%Y')} → {fin.strftime('%d/%m/%Y')}")
        if not inclure_tout else T("Toutes périodes", "All periods")
    )
    kpis = [
        T(f"{len(d)} opérations", f"{len(d)} transactions"),
        T(f"Total {fmt_num(d['Montant'].sum() if not d.empty else 0)} FCFA",
          f"Total {fmt_num(d['Montant'].sum() if not d.empty else 0)} FCFA"),
    ]
    apercu = d

elif type_code == "camions":
    # liste plate camions non déchargés
    if not suivi.empty and not bons.empty and "Numero_Tracteur" in bons.columns:
        map_c = bons.drop_duplicates("Reference").set_index("Reference")["Numero_Tracteur"].map(dd._s)
        suivi = suivi.copy()
        suivi["Camion_BC"] = suivi["Reference"].map(map_c).fillna("")
    else:
        suivi = suivi.copy() if not suivi.empty else pd.DataFrame()
        if not suivi.empty:
            suivi["Camion_BC"] = ""
    sous_bc = suivi[suivi["Statut_Livraison"].isin(["EN ATTENTE", "EN COURS"])].copy() if not suivi.empty else pd.DataFrame()
    if not sous_bc.empty:
        sous_bc["Reste_KG"] = sous_bc["Qte_Commandee_KG"] - sous_bc["Poids_Livre_KG"]
    cols = ["Fournisseur", "Camion_BC", "Reference", "Qte_Commandee_KG", "Poids_Livre_KG", "Reste_KG", "Statut_Livraison"]
    cols = [c for c in cols if c in sous_bc.columns] if not sous_bc.empty else cols
    labels = {
        "Fournisseur": T("Fournisseur", "Supplier"),
        "Camion_BC": T("N° Camion", "Truck No."),
        "Reference": T("N° BC", "PO No."),
        "Qte_Commandee_KG": T("Commandé kg", "Ordered kg"),
        "Poids_Livre_KG": T("Livré kg", "Delivered kg"),
        "Reste_KG": T("Reste kg", "Balance kg"),
        "Statut_Livraison": T("Statut", "Status"),
    }
    headers = [labels.get(c, c) for c in cols]
    fmt = {c: lambda v: fmt_num(v) for c in ("Qte_Commandee_KG", "Poids_Livre_KG", "Reste_KG")}
    rows = df_to_rows(sous_bc, cols, fmt) if not sous_bc.empty else []
    titre = T(
        f"CAMIONS NON DÉCHARGÉS — {dd.libelle_produit(produit)}",
        f"TRUCKS NOT UNLOADED — {dd.libelle_produit(produit)}",
    )
    sous = T("1 ligne = 1 bon / camion encore attendu", "1 row = 1 PO / truck still pending")
    kpis = [T(f"{len(rows)} lignes", f"{len(rows)} rows")]
    apercu = sous_bc

else:  # synthese
    if suivi is None or suivi.empty:
        apercu = pd.DataFrame()
        headers, rows = [], []
    else:
        s = suivi.copy()
        s["Fournisseur"] = s["Fournisseur"].map(dd._s)
        rows_syn = []
        for frs, g in s.groupby("Fournisseur", sort=True):
            qte_cmd = float(g["Qte_Commandee_KG"].sum())
            qte_liv = float(g["Poids_Livre_KG"].sum())
            rows_syn.append({
                "Fournisseur": frs,
                "Nb_Bons": len(g),
                "Nb_Attente": int((g["Statut_Livraison"] == "EN ATTENTE").sum()),
                "Qte_Commandee_KG": qte_cmd,
                "Qte_Livree_KG": qte_liv,
                "Reste_KG": qte_cmd - qte_liv,
                "Montant_Du": float(g["Montant_Du"].sum()),
                "Montant_Paye": float(g["Montant_Paye"].sum()),
                "Solde": float(g["Montant_Du"].sum()) - float(g["Montant_Paye"].sum()),
            })
        apercu = pd.DataFrame(rows_syn)
        cols = list(apercu.columns)
        labels = {
            "Fournisseur": T("Fournisseur", "Supplier"),
            "Nb_Bons": T("Bons", "POs"),
            "Nb_Attente": T("En attente", "Pending"),
            "Qte_Commandee_KG": T("Commandé kg", "Ordered kg"),
            "Qte_Livree_KG": T("Livré kg", "Delivered kg"),
            "Reste_KG": T("Reste kg", "Balance kg"),
            "Montant_Du": T("Dû", "Due"),
            "Montant_Paye": T("Payé", "Paid"),
            "Solde": T("Solde", "Balance"),
        }
        headers = [labels.get(c, c) for c in cols]
        fmt = {c: lambda v: fmt_num(v) for c in cols if c != "Fournisseur"}
        rows = df_to_rows(apercu, cols, fmt)
    titre = T(
        f"SYNTHÈSE FOURNISSEURS — {dd.libelle_produit(produit)}",
        f"SUPPLIER SUMMARY — {dd.libelle_produit(produit)}",
    )
    sous = T("Bons · quantités · soldes", "POs · quantities · balances")
    kpis = [T(f"{len(rows)} fournisseurs", f"{len(rows)} suppliers")]

# ---------- Aperçu ----------
st.markdown("---")
st.subheader(T(f"Aperçu — {type_label}", f"Preview — {type_label}"))
m1, m2, m3 = st.columns(3)
m1.metric(T("Lignes", "Rows"), len(rows))
m2.write(" · ".join(kpis) if kpis else "")
m3.write(sous)

if not rows:
    st.warning(T("Aucune donnée pour ce filtre.", "No data for this filter."))
else:
    st.dataframe(apercu, hide_index=True, width="stretch", height=360)

st.markdown("---")
st.subheader(T("📄 Générer le PDF A3", "📄 Generate A3 PDF"))

col_pdf1, col_pdf2 = st.columns(2)
with col_pdf1:
    gen = st.button(
        T("🖨️ Générer PDF A3 (paysage)", "🖨️ Generate A3 PDF (landscape)"),
        type="primary", width="stretch", disabled=not rows,
    )
with col_pdf2:
    st.caption(
        T(
            "Impression : orientation **Paysage**, format **A3**, ajuster « Ajuster à la page » si besoin.",
            "Print: **Landscape** orientation, **A3** size, use Fit to page if needed.",
        )
    )

if gen and rows:
    try:
        pdf_bytes = generer_pdf_liste(titre, sous, headers, rows, kpis=kpis, lang="EN" if EN else "FR")
        lang_tag = "EN" if EN else "FR"
        nom = f"NAP_A3_{type_code}_{produit}_{lang_tag}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        chemin = os.path.join("impressions_pdf", nom)
        with open(chemin, "wb") as f:
            f.write(pdf_bytes)
        st.success(T(f"PDF généré : `{chemin}`", f"PDF generated: `{chemin}`"))
        st.download_button(
            T("⬇️ Télécharger le PDF A3", "⬇️ Download A3 PDF"),
            data=pdf_bytes,
            file_name=nom,
            mime="application/pdf",
            type="primary",
        )
    except Exception as e:
        st.error(T(f"Erreur PDF : {e}", f"PDF error: {e}"))
        import traceback
        st.code(traceback.format_exc())

st.caption(
    "STE NAP SARL — Impression A3 · Déchargements · Bons · Paiements · Camions · Synthèse fournisseurs"
)
