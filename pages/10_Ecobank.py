# -*- coding: utf-8 -*-
"""
Page Streamlit — EcoBank NAP
Source principale : ECO BANK STATMENT SHEET 2026.xlsx → feuille STATEMENT
Colonnes : Date | DESCRIPTION | CREDIT | DEBIT | SOLDE | NUM DE CHEQUE | MOTIF | CATEGORIE
Relié aux paiements CHEQUE EcoBank (Suivi de paiement.xlsx)
Formulaire d'ajout / mise à jour des mouvements banque.
"""
import streamlit as st
import pandas as pd
from datetime import datetime, date
import os
import re

st.set_page_config(page_title="EcoBank", page_icon="🏦", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

# ==================== CONFIG ====================
FICHIER_ECOBANK = "ECO BANK STATMENT SHEET 2026.xlsx"
FEUILLE_STATEMENT = "STATEMENT"
FEUILLE_CHEQUIER = "SUIVI CHEQUIER"
FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"
FEUILLE_DEPENSES = "DEPENSES"

COLONNES_STMT = [
    "Date", "DESCRIPTION", "CREDIT", "DEBIT", "SOLDE",
    "NUM DE CHEQUE", "MOTIF", "CATEGORIE",
]
CATEGORIES = [
    "PURCHASE/ SOYA", "PURCHASE/ SHEA", "PURCHASE/ PP BAGS",
    "SALES/ SHEA BUTER", "SALES/ SOYA", "SERVICES", "SALARY",
    "FRAIS BANCAIRE", "WITHDRAWAL", "EXPORT ACC", "LOYER",
    "TRANSPORT", "AUTRE",
]
# =================================================


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


def num_cheque_norm(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    s = str(v).strip().replace(" ", "").replace(",", "")
    if s.endswith(".0"):
        s = s[:-2]
    if not s or s.lower() in ("nan", "none", "-", "canceled", "cancelled"):
        return ""
    m = re.search(r"(\d{4,})", s)
    return m.group(1) if m else s


def fmt_montant(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        n = float(v)
        if n == 0:
            return ""
        return f"{n:,.0f}".replace(",", " ")
    except Exception:
        return str(v) if v else ""


def charger_statement():
    """Lit STATEMENT : Date, DESCRIPTION, CREDIT, DEBIT, SOLDE, NUM DE CHEQUE, MOTIF, CATEGORIE."""
    if not os.path.exists(FICHIER_ECOBANK):
        return pd.DataFrame(columns=COLONNES_STMT)
    try:
        xl = pd.ExcelFile(FICHIER_ECOBANK)
        sheet = FEUILLE_STATEMENT if FEUILLE_STATEMENT in xl.sheet_names else None
        if sheet is None:
            for s in xl.sheet_names:
                if "STATEMENT" in s.upper():
                    sheet = s
                    break
        if sheet is None:
            st.error(f"Feuille STATEMENT introuvable. Feuilles : {xl.sheet_names}")
            return pd.DataFrame(columns=COLONNES_STMT)

        raw = pd.read_excel(FICHIER_ECOBANK, sheet_name=sheet, header=None)
        hdr = None
        for i, row in raw.iterrows():
            vals = " ".join(str(x).upper() for x in row.tolist() if x is not None)
            if "DATE" in vals and "DESCRIP" in vals and ("CREDIT" in vals or "DEBIT" in vals):
                hdr = i
                break
        if hdr is None:
            hdr = 2

        df = pd.read_excel(FICHIER_ECOBANK, sheet_name=sheet, header=hdr)
        df = df.dropna(axis=1, how="all")
        rename = {}
        for c in df.columns:
            k = str(c).strip().upper().replace("  ", " ")
            if k == "DATE" or k.startswith("DATE"):
                rename[c] = "Date"
            elif "DESCRIP" in k:
                rename[c] = "DESCRIPTION"
            elif "CREDIT" in k:
                rename[c] = "CREDIT"
            elif "DEBIT" in k:
                rename[c] = "DEBIT"
            elif "SOLDE" in k or "BALANCE" in k:
                rename[c] = "SOLDE"
            elif "CHEQUE" in k or "CHÈQUE" in k:
                rename[c] = "NUM DE CHEQUE"
            elif "MOTIF" in k:
                rename[c] = "MOTIF"
            elif "CATEGOR" in k:
                rename[c] = "CATEGORIE"
        df = df.rename(columns=rename)
        for c in COLONNES_STMT:
            if c not in df.columns:
                df[c] = None

        out = df[COLONNES_STMT].copy()
        out["Date"] = out["Date"].map(_propre)
        out["DESCRIPTION"] = out["DESCRIPTION"].map(_propre)
        out["MOTIF"] = out["MOTIF"].map(_propre)
        out["CATEGORIE"] = out["CATEGORIE"].map(_propre)
        out["NUM DE CHEQUE"] = out["NUM DE CHEQUE"].map(num_cheque_norm)
        for col in ("CREDIT", "DEBIT", "SOLDE"):
            out[col] = pd.to_numeric(out[col], errors="coerce")

        # garder lignes utiles
        out = out[
            (out["DESCRIPTION"] != "")
            | (out["NUM DE CHEQUE"] != "")
            | (out["CREDIT"].fillna(0) != 0)
            | (out["DEBIT"].fillna(0) != 0)
        ]
        return out.reset_index(drop=True)
    except Exception as e:
        st.error(f"Erreur lecture STATEMENT : {e}")
        return pd.DataFrame(columns=COLONNES_STMT)


def charger_paiements_ecobank():
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
                if "transaction" in cols or "cheque" in cols:
                    df = tmp
                    break
            except Exception:
                continue
        if df is None:
            return pd.DataFrame()

        df = df.dropna(axis=1, how="all")
        alias = {
            "date": "Date",
            "no de transaction": "N° transaction",
            "n° de transaction": "N° transaction",
            "prestataires/fournisseurs": "Prestataire",
            "nom du bénéficiaire": "Bénéficiaire",
            "montant": "Montant",
            "mode paiement": "Mode",
            "numero de cheque": "N° chèque",
            "banque": "Banque",
            "description/objet du paiement": "Description",
            "observation": "Observation",
            "type d'operation": "Type",
            "produit": "Produit",
        }
        rename = {}
        for c in df.columns:
            k = str(c).strip().lower()
            if k in alias:
                rename[c] = alias[k]
        df = df.rename(columns=rename)
        if "Mode" not in df.columns:
            return pd.DataFrame()

        mask = df["Mode"].astype(str).str.upper().str.contains("CHEQUE", na=False)
        df = df[mask].copy()
        if "Banque" in df.columns:
            eco = df["Banque"].astype(str).str.upper().str.contains("ECOBANK", na=False)
            vide = df["Banque"].map(_propre) == ""
            df = df[eco | vide]

        out = pd.DataFrame()
        n = len(df)
        empty = pd.Series([""] * n)
        out["Date"] = df.get("Date", empty).map(_propre)
        out["N° transaction"] = df.get("N° transaction", empty).map(_propre)
        out["Prestataire"] = df.get("Prestataire", empty).map(_propre)
        out["Bénéficiaire"] = df.get("Bénéficiaire", empty).map(_propre)
        out["Montant"] = pd.to_numeric(df.get("Montant", 0), errors="coerce").fillna(0)
        out["N° chèque"] = df.get("N° chèque", empty).map(num_cheque_norm)
        out["Description"] = df.get("Description", empty).map(_propre)
        out["Observation"] = df.get("Observation", empty).map(_propre)
        out["Type"] = df.get("Type", empty).map(_propre)
        out["Produit"] = df.get("Produit", empty).map(_propre)
        out = out[out["N° chèque"] != ""]
        return out.reset_index(drop=True)
    except Exception as e:
        st.error(f"Erreur paiements : {e}")
        return pd.DataFrame()


def rapprochement(stmt, paiements):
    if stmt is None or stmt.empty:
        stmt = pd.DataFrame(columns=COLONNES_STMT)
    if paiements is None or paiements.empty:
        paiements = pd.DataFrame(columns=["N° chèque", "Montant"])

    s = stmt.copy()
    s["_key"] = s["NUM DE CHEQUE"].map(num_cheque_norm)
    p = paiements.copy()
    p["_key"] = p["N° chèque"].map(num_cheque_norm)

    keys_bank = set(s["_key"]) - {""}
    keys_pay = set(p["_key"]) - {""}

    merged = s.merge(
        p[["_key", "N° transaction", "Prestataire", "Bénéficiaire", "Montant", "Description", "Observation", "Type", "Produit"]],
        on="_key",
        how="outer",
        indicator=True,
    )

    def statut(row):
        k = row.get("_key") or ""
        if not k:
            return "SANS N° CHÈQUE"
        if k in keys_bank and k in keys_pay:
            return "RAPPROCHÉ"
        if k in keys_bank:
            return "BANQUE SEUL"
        return "PAIEMENT SEUL"

    merged["Statut"] = merged.apply(statut, axis=1)

    def ecart(row):
        try:
            d = float(row.get("DEBIT") or 0)
            m = float(row.get("Montant") or 0)
            if d and m:
                return d - m
        except Exception:
            pass
        return None

    merged["Écart"] = merged.apply(ecart, axis=1)
    return merged


def _trouver_header_statement(ws):
    """Retourne (hdr_row, col_map) pour STATEMENT."""
    col_map = {}
    hdr_row = None
    for r in range(1, 20):
        for c in range(1, 15):
            v = str(ws.cell(r, c).value or "").strip().upper()
            if not v:
                continue
            if v.startswith("DATE") and hdr_row is None:
                hdr_row = r
            if hdr_row == r:
                if v.startswith("DATE"):
                    col_map["Date"] = c
                elif "DESCRIP" in v:
                    col_map["DESCRIPTION"] = c
                elif "CREDIT" in v:
                    col_map["CREDIT"] = c
                elif "DEBIT" in v:
                    col_map["DEBIT"] = c
                elif "SOLDE" in v or "BALANCE" in v:
                    col_map["SOLDE"] = c
                elif "CHEQUE" in v:
                    col_map["NUM DE CHEQUE"] = c
                elif "MOTIF" in v:
                    col_map["MOTIF"] = c
                elif "CATEGOR" in v:
                    col_map["CATEGORIE"] = c
        if hdr_row and "Date" in col_map and "DESCRIPTION" in col_map:
            break
    if not hdr_row:
        hdr_row = 3
        col_map = {
            "Date": 1, "DESCRIPTION": 2, "CREDIT": 3, "DEBIT": 4,
            "SOLDE": 5, "NUM DE CHEQUE": 6, "MOTIF": 7, "CATEGORIE": 8,
        }
    return hdr_row, col_map


def ajouter_mouvement_statement(ligne):
    """Ajoute une ligne en bas de STATEMENT."""
    from openpyxl import load_workbook
    from openpyxl.styles import Border, Side, Font, Alignment, numbers

    wb = load_workbook(FICHIER_ECOBANK)
    if FEUILLE_STATEMENT not in wb.sheetnames:
        raise ValueError("Feuille STATEMENT absente")
    ws = wb[FEUILLE_STATEMENT]
    hdr_row, col_map = _trouver_header_statement(ws)

    last = hdr_row
    col_desc = col_map.get("DESCRIPTION", 2)
    for r in range(ws.max_row, hdr_row, -1):
        if ws.cell(r, col_desc).value not in (None, ""):
            last = r
            break
    new_row = last + 1

    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    val_date = ligne.get("Date")
    if isinstance(val_date, str) and "/" in val_date:
        try:
            val_date = datetime.strptime(val_date, "%d/%m/%Y")
        except ValueError:
            pass

    data = {
        "Date": val_date,
        "DESCRIPTION": ligne.get("DESCRIPTION", ""),
        "CREDIT": float(ligne.get("CREDIT") or 0) or None,
        "DEBIT": float(ligne.get("DEBIT") or 0) or None,
        "SOLDE": None,  # recalcul optionnel
        "NUM DE CHEQUE": num_cheque_norm(ligne.get("NUM DE CHEQUE")),
        "MOTIF": ligne.get("MOTIF", ""),
        "CATEGORIE": ligne.get("CATEGORIE", ""),
    }

    # solde = dernier solde - debit + credit
    col_solde = col_map.get("SOLDE")
    last_solde = 0
    if col_solde:
        for r in range(last, hdr_row, -1):
            v = ws.cell(r, col_solde).value
            if v is not None:
                try:
                    last_solde = float(v)
                    break
                except Exception:
                    pass
    credit = float(ligne.get("CREDIT") or 0)
    debit = float(ligne.get("DEBIT") or 0)
    data["SOLDE"] = last_solde + credit - debit

    for key, col in col_map.items():
        cell = ws.cell(new_row, col, data.get(key))
        cell.border = thin
        if key in ("CREDIT", "DEBIT", "SOLDE") and cell.value is not None:
            cell.number_format = "#,##0"

    wb.save(FICHIER_ECOBANK)
    return new_row


def mettre_a_jour_par_cheque(num_cheque, champs):
    """Met à jour la première ligne STATEMENT avec ce N° chèque."""
    from openpyxl import load_workbook
    from openpyxl.styles import Border, Side

    num = num_cheque_norm(num_cheque)
    if not num:
        raise ValueError("N° chèque vide")

    wb = load_workbook(FICHIER_ECOBANK)
    ws = wb[FEUILLE_STATEMENT]
    hdr_row, col_map = _trouver_header_statement(ws)
    col_chq = col_map.get("NUM DE CHEQUE", 6)

    cible = None
    for r in range(hdr_row + 1, ws.max_row + 1):
        if num_cheque_norm(ws.cell(r, col_chq).value) == num:
            cible = r
            break
    if cible is None:
        raise ValueError(f"Chèque {num} introuvable dans STATEMENT")

    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    for key, val in champs.items():
        if key not in col_map or val is None or val == "":
            continue
        col = col_map[key]
        if key == "Date" and isinstance(val, str) and "/" in val:
            try:
                val = datetime.strptime(val, "%d/%m/%Y")
            except ValueError:
                pass
        if key in ("CREDIT", "DEBIT", "SOLDE"):
            try:
                val = float(val)
            except Exception:
                pass
        if key == "NUM DE CHEQUE":
            val = num_cheque_norm(val)
        cell = ws.cell(cible, col, val)
        cell.border = thin
        if key in ("CREDIT", "DEBIT", "SOLDE") and isinstance(val, (int, float)):
            cell.number_format = "#,##0"

    wb.save(FICHIER_ECOBANK)
    return cible


def envoyer_paiements_vers_statement(lignes_pay):
    """Ajoute les paiements absents du STATEMENT (DEBIT = montant)."""
    from openpyxl import load_workbook

    wb = load_workbook(FICHIER_ECOBANK)
    ws = wb[FEUILLE_STATEMENT]
    hdr_row, col_map = _trouver_header_statement(ws)
    col_chq = col_map.get("NUM DE CHEQUE", 6)

    existants = set()
    for r in range(hdr_row + 1, ws.max_row + 1):
        existants.add(num_cheque_norm(ws.cell(r, col_chq).value))

    ajoutes, ignores = 0, 0
    for _, row in lignes_pay.iterrows():
        num = num_cheque_norm(row.get("N° chèque") or row.get("_key"))
        if not num:
            continue
        if num in existants:
            ignores += 1
            continue
        desc = _propre(row.get("Prestataire")) or _propre(row.get("Bénéficiaire")) or "—"
        ligne = {
            "Date": row.get("Date"),
            "DESCRIPTION": desc,
            "CREDIT": 0,
            "DEBIT": float(row.get("Montant") or 0),
            "NUM DE CHEQUE": num,
            "MOTIF": _propre(row.get("Description")) or _propre(row.get("Observation")),
            "CATEGORIE": "",
        }
        # produit → catégorie approximative
        prod = _propre(row.get("Produit")).upper()
        if "SOYA" in prod:
            ligne["CATEGORIE"] = "PURCHASE/ SOYA"
        elif "SHEA" in prod:
            ligne["CATEGORIE"] = "PURCHASE/ SHEA"
        ajouter_mouvement_statement(ligne)
        existants.add(num)
        ajoutes += 1
    return ajoutes, ignores


# ==================== UI ====================
st.title("🏦 EcoBank — Relevé & rapprochement paiements")
st.caption(
    f"**Source banque** : `{FICHIER_ECOBANK}` → feuille **STATEMENT** "
    f"(Date · DESCRIPTION · CREDIT · DEBIT · SOLDE · N° chèque · Motif · Catégorie) · "
    f"**Paiements** : `{FICHIER_PAIEMENTS}`"
)

stmt = charger_statement()
paiements = charger_paiements_ecobank()
rap = rapprochement(stmt, paiements)

# ----- KPI -----
c1, c2, c3, c4, c5 = st.columns(5)
total_credit = float(stmt["CREDIT"].fillna(0).sum()) if not stmt.empty else 0
total_debit = float(stmt["DEBIT"].fillna(0).sum()) if not stmt.empty else 0
n_rap = int((rap["Statut"] == "RAPPROCHÉ").sum()) if not rap.empty else 0
n_pay_seul = int((rap["Statut"] == "PAIEMENT SEUL").sum()) if not rap.empty else 0
dernier_solde = ""
if not stmt.empty and stmt["SOLDE"].notna().any():
    # dernière ligne avec solde
    s_ok = stmt[stmt["SOLDE"].notna()]
    if not s_ok.empty:
        dernier_solde = fmt_montant(s_ok.iloc[-1]["SOLDE"])

c1.metric("Lignes STATEMENT", len(stmt))
c2.metric("Total CRÉDIT (FCFA)", fmt_montant(total_credit) or "0")
c3.metric("Total DÉBIT (FCFA)", fmt_montant(total_debit) or "0")
c4.metric("Rapprochés / Paiements absents", f"{n_rap} / {n_pay_seul}")
c5.metric("Dernier solde", dernier_solde or "—")

st.markdown("---")

# ----- Journal banque lisible -----
st.subheader("📒 Journal EcoBank (STATEMENT)")
st.markdown("Colonnes principales pour la lecture : **Date · Description · Crédit · Débit · Solde · N° chèque · Motif**")

f1, f2, f3, f4 = st.columns(4)
with f1:
    sens = st.selectbox("Sens", ["Tous", "Crédit uniquement", "Débit uniquement", "Avec n° chèque"])
with f2:
    cats = ["Toutes"] + sorted({c for c in stmt["CATEGORIE"].tolist() if c}) if not stmt.empty else ["Toutes"]
    fcat = st.selectbox("Catégorie", cats)
with f3:
    q = st.text_input("Recherche", placeholder="Nom, motif, n° chèque…")
with f4:
    n_aff = st.selectbox("Afficher", [50, 100, 200, 500, "Tout"], index=1)

vue_b = stmt.copy()
if not vue_b.empty:
    if sens == "Crédit uniquement":
        vue_b = vue_b[vue_b["CREDIT"].fillna(0) > 0]
    elif sens == "Débit uniquement":
        vue_b = vue_b[vue_b["DEBIT"].fillna(0) > 0]
    elif sens == "Avec n° chèque":
        vue_b = vue_b[vue_b["NUM DE CHEQUE"] != ""]
    if fcat != "Toutes":
        vue_b = vue_b[vue_b["CATEGORIE"] == fcat]
    if q.strip():
        qq = q.strip().lower()
        vue_b = vue_b[vue_b.apply(lambda r: qq in " ".join(r.astype(str)).lower(), axis=1)]

# affichage formaté
if vue_b.empty:
    st.warning("Aucune ligne banque. Vérifiez le fichier Excel.")
else:
    aff = vue_b.copy()
    aff["CREDIT (FCFA)"] = aff["CREDIT"].map(fmt_montant)
    aff["DEBIT (FCFA)"] = aff["DEBIT"].map(fmt_montant)
    aff["SOLDE (FCFA)"] = aff["SOLDE"].map(fmt_montant)
    cols = ["Date", "DESCRIPTION", "CREDIT (FCFA)", "DEBIT (FCFA)", "SOLDE (FCFA)", "NUM DE CHEQUE", "MOTIF", "CATEGORIE"]
    lim = len(aff) if n_aff == "Tout" else int(n_aff)
    st.dataframe(aff[cols].tail(lim) if n_aff != "Tout" else aff[cols], width="stretch", height=380, hide_index=True)

st.markdown("---")

# ----- Formulaire ajout / maj -----
st.subheader("✏️ Ajouter ou mettre à jour un mouvement banque")

tab_add, tab_upd = st.tabs(["➕ Ajouter une ligne STATEMENT", "🔄 Mettre à jour par N° chèque"])

with tab_add:
    st.caption("Saisie libre d’un mouvement (crédit ou débit) dans le relevé EcoBank.")
    with st.form("form_add_stmt", clear_on_submit=True):
        a1, a2, a3 = st.columns(3)
        with a1:
            d_add = st.date_input("Date *", value=date.today())
            desc_add = st.text_input("Description *", placeholder="Nom bénéficiaire / libellé banque")
            cat_add = st.selectbox("Catégorie", [""] + CATEGORIES)
        with a2:
            credit_add = st.number_input("CRÉDIT (FCFA)", min_value=0.0, step=1000.0, format="%.0f")
            debit_add = st.number_input("DÉBIT (FCFA)", min_value=0.0, step=1000.0, format="%.0f")
            st.caption("Renseignez **soit** Crédit **soit** Débit (pas les deux sauf cas spécial).")
        with a3:
            chq_add = st.text_input("N° chèque", placeholder="2000751 (si applicable)")
            motif_add = st.text_input("Motif", placeholder="PURCHASE OF SHEA / Frais…")
        ok_add = st.form_submit_button("💾 Enregistrer dans STATEMENT", type="primary", width="stretch")
        if ok_add:
            if not desc_add.strip():
                st.error("Description obligatoire.")
            elif credit_add <= 0 and debit_add <= 0:
                st.error("Indiquez un montant Crédit ou Débit.")
            else:
                try:
                    row_n = ajouter_mouvement_statement({
                        "Date": d_add.strftime("%d/%m/%Y"),
                        "DESCRIPTION": desc_add.strip(),
                        "CREDIT": credit_add,
                        "DEBIT": debit_add,
                        "NUM DE CHEQUE": chq_add.strip(),
                        "MOTIF": motif_add.strip(),
                        "CATEGORIE": cat_add,
                    })
                    st.success(f"Ligne ajoutée (ligne Excel ≈ {row_n}). Solde recalculé.")
                    st.rerun()
                except Exception as e:
                    st.error(str(e))

with tab_upd:
    st.caption("Corriger DESCRIPTION, montants, motif… d’un chèque déjà présent.")
    nums_dispo = sorted({n for n in stmt["NUM DE CHEQUE"].tolist() if n}) if not stmt.empty else []
    with st.form("form_upd_stmt"):
        u1, u2 = st.columns(2)
        with u1:
            chq_u = st.selectbox("N° chèque à modifier *", [""] + nums_dispo)
            desc_u = st.text_input("Nouvelle description (vide = inchangé)")
            motif_u = st.text_input("Nouveau motif (vide = inchangé)")
        with u2:
            credit_u = st.text_input("Nouveau CRÉDIT (vide = inchangé)", placeholder="ex. 5000000")
            debit_u = st.text_input("Nouveau DÉBIT (vide = inchangé)", placeholder="ex. 1200000")
            cat_u = st.selectbox("Nouvelle catégorie", ["(inchangé)"] + CATEGORIES)
        ok_u = st.form_submit_button("🔄 Mettre à jour", width="stretch")
        if ok_u:
            if not chq_u:
                st.error("Choisissez un N° chèque.")
            else:
                champs = {}
                if desc_u.strip():
                    champs["DESCRIPTION"] = desc_u.strip()
                if motif_u.strip():
                    champs["MOTIF"] = motif_u.strip()
                if cat_u != "(inchangé)":
                    champs["CATEGORIE"] = cat_u
                if credit_u.strip():
                    try:
                        champs["CREDIT"] = float(credit_u.replace(" ", ""))
                    except ValueError:
                        st.error("Crédit invalide")
                        st.stop()
                if debit_u.strip():
                    try:
                        champs["DEBIT"] = float(debit_u.replace(" ", ""))
                    except ValueError:
                        st.error("Débit invalide")
                        st.stop()
                if not champs:
                    st.warning("Aucun champ à modifier.")
                else:
                    try:
                        r = mettre_a_jour_par_cheque(chq_u, champs)
                        st.success(f"Chèque {chq_u} mis à jour (ligne {r}).")
                        st.rerun()
                    except Exception as e:
                        st.error(str(e))

st.markdown("---")

# ----- Rapprochement -----
st.subheader("🔗 Rapprochement STATEMENT ↔ Paiements EcoBank")
fs1, fs2 = st.columns(2)
with fs1:
    fstat = st.selectbox("Statut", ["Tous", "RAPPROCHÉ", "BANQUE SEUL", "PAIEMENT SEUL", "SANS N° CHÈQUE"])
with fs2:
    q2 = st.text_input("Filtrer rapprochement", placeholder="N° chèque, prestataire…", key="q_rap")

vue_r = rap.copy() if not rap.empty else pd.DataFrame()
if not vue_r.empty:
    if fstat != "Tous":
        vue_r = vue_r[vue_r["Statut"] == fstat]
    if q2.strip():
        qq = q2.strip().lower()
        vue_r = vue_r[vue_r.apply(lambda r: qq in " ".join(r.astype(str)).lower(), axis=1)]

    aff_r = vue_r.copy()
    aff_r["CREDIT (FCFA)"] = aff_r.get("CREDIT", pd.Series(dtype=float)).map(fmt_montant)
    aff_r["DEBIT (FCFA)"] = aff_r.get("DEBIT", pd.Series(dtype=float)).map(fmt_montant)
    aff_r["Montant paiement"] = aff_r.get("Montant", pd.Series(dtype=float)).map(fmt_montant)
    aff_r["Écart"] = aff_r.get("Écart", pd.Series(dtype=float)).map(fmt_montant)
    cols_r = [c for c in [
        "Date", "NUM DE CHEQUE", "DESCRIPTION", "DEBIT (FCFA)", "CREDIT (FCFA)", "MOTIF",
        "N° transaction", "Prestataire", "Montant paiement", "Statut", "Écart",
    ] if c in aff_r.columns]
    st.dataframe(aff_r[cols_r], width="stretch", height=320, hide_index=True)

# Envoi paiements manquants
seul = rap[rap["Statut"] == "PAIEMENT SEUL"] if not rap.empty else pd.DataFrame()
st.markdown("##### 📤 Paiements EcoBank absents du relevé")
if seul.empty:
    st.success("Aucun paiement EcoBank en attente d’envoi vers STATEMENT.")
else:
    st.info(f"**{len(seul)}** chèque(s) dans les paiements mais pas encore dans STATEMENT.")
    prev = seul[["Date", "N° chèque", "Prestataire", "Montant", "Description"]].copy() if "N° chèque" in seul.columns else seul
    if "N° chèque" not in prev.columns and "_key" in seul.columns:
        prev = seul.copy()
        prev["N° chèque"] = prev["_key"]
        prev = prev[["Date", "N° chèque", "Prestataire", "Montant", "Description"]]
    prev["Montant"] = prev["Montant"].map(fmt_montant)
    st.dataframe(prev, width="stretch", hide_index=True)
    if st.button("📤 Envoyer ces paiements vers STATEMENT (en DÉBIT)", type="primary"):
        try:
            # s'assurer colonnes
            df_send = seul.copy()
            if "N° chèque" not in df_send.columns:
                df_send["N° chèque"] = df_send.get("_key", "")
            aj, ig = envoyer_paiements_vers_statement(df_send)
            st.success(f"Ajoutés : **{aj}** · Déjà présents : **{ig}**")
            st.rerun()
        except Exception as e:
            st.error(str(e))
            import traceback
            st.code(traceback.format_exc())

st.markdown("---")
with st.expander("ℹ️ Comment lire cette page"):
    st.markdown(
        """
| Élément | Signification |
|--------|----------------|
| **CRÉDIT** | Argent **entrant** sur le compte EcoBank (ventes, virements reçus…) |
| **DÉBIT** | Argent **sortant** (chèques émis, frais, retraits…) |
| **SOLDE** | Position du compte après la ligne |
| **RAPPROCHÉ** | Même N° chèque dans STATEMENT et dans Suivi de paiement |
| **BANQUE SEUL** | Sur le relevé, pas encore saisi dans les paiements NAP |
| **PAIEMENT SEUL** | Saisi en paiement NAP, pas encore sur le relevé EcoBank |

**Formulaire Ajouter** : pour coller une ligne du relevé papier / extranet banque.  
**Mettre à jour** : corriger un chèque déjà listé (libellé, montant, motif).
"""
    )

st.caption(
    f"NAP SARL · EcoBank · `{FICHIER_ECOBANK}` + `{FICHIER_PAIEMENTS}` · Page `pages/6_Ecobank.py`"
)
