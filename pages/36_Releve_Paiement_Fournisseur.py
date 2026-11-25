# -*- coding: utf-8 -*-
"""
NAP SARL — Relevé fournisseur
Filtre fournisseur + produit + période · fiche A4 Word (style bon de paiement).
Données : Suivi de paiement.xlsx → feuille DEPENSES.
"""
import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta
import os
import re
import glob

st.set_page_config(page_title="Relevé Fournisseur", page_icon="📄", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"
FEUILLE_DEPENSES = "DEPENSES"
DOSSIER_RELEVES = "releves_fournisseurs"
os.makedirs(DOSSIER_RELEVES, exist_ok=True)

RESPONSABLES_NAP = [
    "Alliance BODJRENOU",
    "Manager Logistics",
    "Comptabilité NAP SARL",
    "Direction NAP SARL",
]


def _propre(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, (datetime, date, pd.Timestamp)):
        return v.strftime("%d/%m/%Y")
    t = str(v).strip()
    if t.lower() in ("", "nan", "none", "nat", "-"):
        return ""
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", t)
    if m:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    return t


def charger_depenses():
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
                if "transaction" in cols or "prestataire" in cols:
                    df = tmp
                    break
            except Exception:
                continue
        if df is None:
            return pd.DataFrame()
        df = df.dropna(axis=1, how="all")
        alias = {
            "date": "Date",
            "no de transaction": "No_transaction",
            "n° de transaction": "No_transaction",
            "type d'operation": "Type",
            "type d'opération": "Type",
            "produit": "Produit",
            "prestataires/fournisseurs": "Fournisseur",
            "prestataire": "Fournisseur",
            "fournisseur": "Fournisseur",
            "nom du bénéficiaire": "Beneficiaire",
            "ref doc": "Ref_doc",
            "description/objet du paiement": "Description",
            "description": "Description",
            "montant": "Montant",
            "observation": "Observation",
            "mode paiement": "Mode",
            "numero de cheque": "Num_cheque",
            "banque": "Banque",
            "n_bc": "N_BC",
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
        df["Fournisseur"] = df["Fournisseur"].map(_propre)
        df["Montant"] = pd.to_numeric(df["Montant"], errors="coerce").fillna(0)
        df["Date_dt"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)
        df = df[df["Fournisseur"] != ""]
        df = df[df["No_transaction"].map(_propre) != ""]
        return df.reset_index(drop=True)
    except Exception as e:
        st.error(f"Erreur lecture Excel : {e}")
        return pd.DataFrame()


def meme_nom(a, b):
    """Comparaison souple des noms fournisseurs."""
    def tokens(x):
        t = re.sub(r"[^A-Za-z0-9]+", " ", str(x or "").upper())
        return {w for w in t.split() if len(w) >= 3}
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb:
        return str(a).strip().casefold() == str(b).strip().casefold()
    return len(ta & tb) >= min(2, len(ta), len(tb)) or str(a).strip().casefold() == str(b).strip().casefold()


# ---------- UI ----------
st.title("📄 Relevé fournisseur")
st.caption(
    f"Source : `{FICHIER_PAIEMENTS}` (DEPENSES) · Fiche A4 Word dans `{DOSSIER_RELEVES}/` "
    "· Pied de page NAP SARL · 2 signatures (Fournisseur | Responsable NAP)"
)

df = charger_depenses()
if df.empty:
    st.warning(f"Aucune donnée. Placez `{FICHIER_PAIEMENTS}` à la racine du projet.")
    st.stop()

fournisseurs = sorted({f for f in df["Fournisseur"].tolist() if f})

# Produits disponibles (tous, puis affinés après choix fournisseur)
produits_tous = sorted({_propre(p) for p in df["Produit"].tolist() if _propre(p)})

# Filtres
c1, c2, c3, c4 = st.columns([2.2, 1.4, 1.2, 1.2])
with c1:
    frs = st.selectbox("Fournisseur *", [""] + fournisseurs, format_func=lambda x: "— Choisir —" if x == "" else x)
with c2:
    # Liste produits : si fournisseur choisi, limiter aux produits de ce frs
    if frs:
        mask_tmp = df["Fournisseur"].map(lambda x: meme_nom(x, frs) or x.casefold() == frs.casefold())
        produits_opts = sorted({_propre(p) for p in df.loc[mask_tmp, "Produit"].tolist() if _propre(p)})
        if not produits_opts:
            produits_opts = produits_tous
    else:
        produits_opts = produits_tous
    produit = st.selectbox(
        "Produit",
        ["Tous"] + produits_opts,
        help="Filtrer le relevé sur un produit (SHEA, SOYABEAN…) ou laisser Tous.",
    )
with c3:
    # Période par défaut : depuis le 01/01/2025 jusqu'à aujourd'hui
    defaut_debut = date(2025, 1, 1)
    defaut_fin = date.today()
    debut = st.date_input("Du *", value=defaut_debut, min_value=date(2025, 1, 1))
with c4:
    fin = st.date_input("Au *", value=defaut_fin, min_value=date(2025, 1, 1))

responsable = "NAP SARL"

if not frs:
    st.info("Sélectionnez un fournisseur pour afficher son relevé.")
    st.stop()

if debut > fin:
    st.error("La date de début doit être antérieure à la date de fin.")
    st.stop()

# Filtrer : fournisseur + période + produit
mask_frs = df["Fournisseur"].map(lambda x: meme_nom(x, frs) or x.casefold() == frs.casefold())
sous = df[mask_frs].copy()
sous = sous[(sous["Date_dt"].dt.date >= debut) & (sous["Date_dt"].dt.date <= fin)]
if produit != "Tous":
    sous = sous[sous["Produit"].map(_propre).str.upper() == produit.upper()]
sous = sous.sort_values("Date_dt")

total = float(sous["Montant"].sum())
nb = len(sous)

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Fournisseur", frs[:24] + ("…" if len(frs) > 24 else ""))
k2.metric("Produit", produit)
k3.metric("Opérations", nb)
k4.metric("Total période (FCFA)", f"{total:,.0f}".replace(",", " "))
k5.metric("Période", f"{debut.strftime('%d/%m/%Y')} → {fin.strftime('%d/%m/%Y')}")

st.markdown("---")
st.subheader("📋 Mouvements de la période")

if sous.empty:
    st.warning("Aucun paiement pour ce fournisseur sur la période choisie.")
else:
    aff = sous.copy()
    aff["Date"] = aff["Date_dt"].dt.strftime("%d/%m/%Y")
    cols_show = ["Date", "No_transaction", "Produit", "Description", "Ref_doc", "Montant", "Mode", "Observation", "N_BC"]
    cols_show = [c for c in cols_show if c in aff.columns]
    st.dataframe(
        aff[cols_show],
        hide_index=True,
        width="stretch",
        height=360,
        column_config={
            "No_transaction": "N° transaction",
            "Ref_doc": "Réf. doc",
            "Montant": st.column_config.NumberColumn("Montant", format="%d"),
            "N_BC": "N° BC",
        },
    )

st.markdown("---")
st.subheader("📊 État du compte à la date de fin")
st.caption(
    "Indiquez le **reste à payer** au " + fin.strftime("%d/%m/%Y") +
    ". Mettez **0** si le compte est soldé."
)
solde_restant = st.number_input(
    "Reste à payer / Solde du compte (FCFA)",
    min_value=0.0,
    value=0.0,
    step=1000.0,
    format="%.0f",
    help="0 = compte soldé. Sinon montant encore dû au fournisseur à cette date.",
)
if solde_restant <= 0:
    st.success(f"État au {fin.strftime('%d/%m/%Y')} : **Solde 0 FCFA — Compte soldé**")
else:
    st.warning(
        f"État au {fin.strftime('%d/%m/%Y')} : **Reste à payer "
        f"{solde_restant:,.0f} FCFA**".replace(",", " ")
    )

st.markdown("---")
st.subheader("📄 Générer la fiche relevé A4 (Word)")

st.markdown(
    """
| Élément | Contenu |
|--------|---------|
| **En-tête** | STE NAP SARL + QR |
| **Titre** | RELEVÉ FOURNISSEUR + ligne bleu ciel |
| **Tableau** | Date · N° transaction · Produit · Description · Montant · Mode |
| **Total** | Somme de la période (vert) |
| **État du compte** | À la date de fin : reste à payer ou solde 0 FCFA |
| **Signatures** | **Gauche** : nom du fournisseur · **Droite** : NAP SARL |
| **Pied de page** | Adresse / RCCM / IFU / Tél (comme bon de paiement) |
"""
)

if st.button("📄 Générer le relevé Word", type="primary", width="stretch", disabled=sous.empty):
    lignes = []
    for _, r in sous.iterrows():
        lignes.append({
            "Date": r["Date_dt"].strftime("%d/%m/%Y") if pd.notna(r["Date_dt"]) else _propre(r.get("Date")),
            "No_transaction": _propre(r.get("No_transaction")),
            "Produit": _propre(r.get("Produit")),
            "Description": _propre(r.get("Description")) or _propre(r.get("Ref_doc")),
            "Ref_doc": _propre(r.get("Ref_doc")),
            "Mode": _propre(r.get("Mode")),
            "Montant": float(r.get("Montant") or 0),
            "Observation": _propre(r.get("Observation")),
        })
    data = {
        "Fournisseur": frs,
        "Produit": "" if produit == "Tous" else produit,
        "Periode_Debut": debut.strftime("%d/%m/%Y"),
        "Periode_Fin": fin.strftime("%d/%m/%Y"),
        "Total": total,
        "Nb_operations": nb,
        "Solde_Restant": float(solde_restant),
        "Responsable_NAP": responsable,
        "Utilisateur": name,
        "Lignes": lignes,
    }
    try:
        from generer_releve_fournisseur import generer_releve_fournisseur
        chemin = generer_releve_fournisseur(data)
        st.success(f"Relevé généré : `{chemin}`")
        with open(chemin, "rb") as f:
            st.download_button(
                "⬇️ Télécharger le relevé Word",
                data=f.read(),
                file_name=os.path.basename(chemin),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
    except Exception as e:
        st.error(f"Erreur génération : {e}")
        import traceback
        st.code(traceback.format_exc())

# Liste des relevés déjà générés
st.markdown("---")
st.subheader("📥 Relevés déjà générés")
fichiers = sorted(glob.glob(os.path.join(DOSSIER_RELEVES, "Releve_*.docx")), reverse=True)
if not fichiers:
    st.caption("Aucun relevé dans le dossier pour le moment.")
else:
    for fpath in fichiers[:20]:
        c1, c2 = st.columns([4, 1])
        with c1:
            st.write(f"📄 **{os.path.basename(fpath)}**")
        with c2:
            with open(fpath, "rb") as f:
                st.download_button(
                    "⬇️",
                    data=f.read(),
                    file_name=os.path.basename(fpath),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_{fpath}",
                )

st.caption("NAP SARL — Relevé fournisseur · données Suivi de paiement · fiche A4 style bon de paiement")
