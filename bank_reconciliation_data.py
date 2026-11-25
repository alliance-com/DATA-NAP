# -*- coding: utf-8 -*-
"""
NAP SARL — Rapprochement bancaire (analyste)
Compare Suivi de paiement (DEPENSES) ↔ relevé banque (Excel local ou PDF).
OCR Windows : Tesseract + Poppler pour relevés scannés (ECOBANK, ATLANTIC).
"""
from __future__ import annotations

import os
import re
import sys
from datetime import datetime, date
from io import BytesIO
from pathlib import Path as PathLib

import pandas as pd

FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"
FEUILLE_DEPENSES = "DEPENSES"
FICHIER_MANUEL = "ecritures_bancaires_manuelles.xlsx"
FICHIER_RELEVE_IMPORT_PREFIX = "releve_importe_"  # releve_importe_ECOBANK.xlsx

BANQUES = ["ECOBANK", "ATLANTIC", "BIIC", "BGFI", "UBA", "AUTRE"]

FICHIERS_RELEVES_LOCAUX = {
    "ECOBANK": [
        "ECO BANK STATMENT SHEET 2026.xlsx", "ECO BANK STATEMENT SHEET 2026.xlsx",
        "ECO BANK STATMENT SHEET 2026-2.xlsx", "ECO BANK STATMENT SHEET 2026-3.xlsx",
        "ecobank_statement.xlsx",
    ],
    "ATLANTIC": [
        "ATLANTIC BANK STATMENT SHEET 2026.xlsx", "ATLANTIC BANK STATEMENT SHEET 2026.xlsx",
        "ATLANTIC BANK STATMENT SHEET 2026-2.xlsx", "atlantic_statement.xlsx",
    ],
    "BIIC": [
        "BIIC BANK STATMENT SHEET 2026.xlsx", "BIIC BANK STATEMENT SHEET 2026.xlsx",
        "BIIC BANK STATMENT SHEET 2026_1.xlsx", "BIIC BANK STATMENT SHEET 2026_1-2.xlsx",
        "biic_statement.xlsx",
    ],
    "BGFI": ["BGFI BANK STATEMENT 2026.xlsx", "bgfi_statement.xlsx"],
}


# ---------- OCR Windows ----------
def _configurer_ocr():
    """Configure pytesseract + chemins Windows courants."""
    try:
        import pytesseract
    except ImportError:
        return False, "pytesseract non installé (pip install pytesseract)"

    candidats = [
        os.environ.get("TESSERACT_CMD"),
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"C:\Tesseract-OCR\tesseract.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"),
    ]
    for c in candidats:
        if c and os.path.isfile(c):
            pytesseract.pytesseract.tesseract_cmd = c
            return True, c
    # déjà dans PATH ?
    try:
        pytesseract.get_tesseract_version()
        return True, "PATH"
    except Exception:
        return False, "Tesseract introuvable — installez-le ou définissez TESSERACT_CMD"


def _configurer_poppler():
    """Retourne le dossier bin Poppler si trouvé (Windows)."""
    candidats = [
        os.environ.get("POPPLER_PATH"),
        r"C:\poppler-26.09.0\Library\bin",
        r"C:\poppler\Library\bin",
        r"C:\poppler\bin",
        r"C:\Program Files\poppler\Library\bin",
        r"C:\Program Files\poppler\bin",
        os.path.expandvars(r"%USERPROFILE%\poppler\Library\bin"),
    ]
    for c in candidats:
        if c and os.path.isdir(c):
            return c
    # Recherche C:\poppler-*\Library\bin
    try:
        import glob
        for d in sorted(glob.glob(r"C:\poppler-*\Library\bin"), reverse=True):
            if os.path.isfile(os.path.join(d, "pdftoppm.exe")):
                return d
    except Exception:
        pass
    return None


def _s(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none", "nat") else t


def _n(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        if isinstance(v, str):
            v = v.replace("\xa0", " ").replace(" ", "").replace(",", ".")
            v = re.sub(r"[^\d.\-]", "", v)
        return float(v) if v not in ("", ".", "-") else 0.0
    except Exception:
        return 0.0


def _d(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return pd.NaT
    if isinstance(v, datetime):
        return pd.Timestamp(v.date())
    if isinstance(v, date):
        return pd.Timestamp(v)
    s = str(v).strip()
    m = re.match(r"(\d{1,2})[-/\s]([A-Za-z]{3,})[-/\s](\d{2,4})", s)
    if m:
        months = {
            "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
            "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
            "JANV": 1, "FEV": 2, "AVR": 4, "MAI": 5, "JUIN": 6, "JUIL": 7,
            "AOUT": 8, "SEPT": 9, "OCTO": 10, "NOVE": 11, "DECE": 12,
        }
        key = m.group(2).upper()
        mo = months.get(key[:3]) or months.get(key[:4])
        if mo:
            y = int(m.group(3))
            if y < 100:
                y += 2000
            try:
                return pd.Timestamp(date(y, mo, int(m.group(1))))
            except Exception:
                pass
    return pd.to_datetime(v, errors="coerce", dayfirst=True)


def fmt(n):
    try:
        return f"{float(n):,.0f}".replace(",", " ")
    except Exception:
        return str(n)


def normaliser_banque(x):
    t = _s(x).upper()
    if not t:
        return ""
    if "ECO" in t:
        return "ECOBANK"
    if "ATLAN" in t:
        return "ATLANTIC"
    if "BIIC" in t:
        return "BIIC"
    if "BGFI" in t:
        return "BGFI"
    if "UBA" in t:
        return "UBA"
    if "CAISSE" in t:
        return "CAISSE"
    return t


# ---------- Suivi de paiement ----------
def _trouver_suivi():
    for cand in [
        FICHIER_PAIEMENTS,
        "Suivi de paiement-5.xlsx",
        "Suivi de paiement-4.xlsx",
        "Suivi de paiement-3.xlsx",
        "Suivi de paiement-2.xlsx",
    ]:
        if os.path.exists(cand):
            return cand
    return None


def charger_journal(banque=None, date_debut=None, date_fin=None):
    cols = [
        "Date", "No_Transaction", "Type", "Produit", "Prestataire", "Beneficiaire",
        "Ref_Doc", "Description", "Montant", "Observation", "Mode", "Num_Cheque", "Banque",
    ]
    vide = pd.DataFrame(columns=cols)
    chemin = _trouver_suivi()
    if not chemin:
        return vide
    try:
        xl = pd.ExcelFile(chemin)
        feuille = FEUILLE_DEPENSES if FEUILLE_DEPENSES in xl.sheet_names else xl.sheet_names[0]
        df = None
        for hdr in (0, 1, 2, 3):
            tmp = pd.read_excel(chemin, sheet_name=feuille, header=hdr)
            cols_l = [str(c).strip().lower() for c in tmp.columns]
            if any("transaction" in c for c in cols_l) and any("montant" in c for c in cols_l):
                df = tmp
                break
        if df is None:
            return vide
    except Exception:
        return vide

    alias = {
        "date": "Date",
        "no de transaction": "No_Transaction", "n° de transaction": "No_Transaction",
        "n° transaction": "No_Transaction", "no transaction": "No_Transaction",
        "type d'operation": "Type", "type d’operation": "Type", "type d'opération": "Type",
        "produit": "Produit",
        "prestataires/fournisseurs": "Prestataire", "prestataire": "Prestataire",
        "nom du bénéficiaire": "Beneficiaire", "nom du beneficiaire": "Beneficiaire",
        "ref doc": "Ref_Doc", "réf doc": "Ref_Doc",
        "description/objet du paiement": "Description", "description": "Description",
        "montant": "Montant",
        "observation": "Observation",
        "mode paiement": "Mode", "mode de paiement": "Mode",
        "numero de cheque": "Num_Cheque", "numéro de chèque": "Num_Cheque",
        "n° chèque": "Num_Cheque", "num cheque": "Num_Cheque",
        "banque": "Banque",
    }
    rename = {}
    for c in df.columns:
        k = str(c).strip().lower()
        if k in alias:
            rename[c] = alias[k]
    df = df.rename(columns=rename)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    for c in ["No_Transaction", "Type", "Produit", "Prestataire", "Beneficiaire",
              "Ref_Doc", "Description", "Observation", "Mode", "Num_Cheque"]:
        df[c] = df[c].map(_s)
    df["Montant"] = df["Montant"].map(_n)
    df["Date"] = df["Date"].map(_d)
    df["Banque"] = df["Banque"].map(normaliser_banque)
    df = df[df["Montant"] != 0].copy()

    if banque and banque.upper() not in ("TOUTES", "ALL", ""):
        b = normaliser_banque(banque)
        df = df[df["Banque"] == b]
    if date_debut is not None:
        df = df[df["Date"] >= pd.Timestamp(date_debut)]
    if date_fin is not None:
        df = df[df["Date"] <= pd.Timestamp(date_fin)]
    return df.reset_index(drop=True)


# ---------- Relevé Excel local ----------
def _trouver_fichier_local(banque):
    for nom in FICHIERS_RELEVES_LOCAUX.get(normaliser_banque(banque), []):
        if os.path.exists(nom):
            return nom
    b = normaliser_banque(banque)
    for f in os.listdir("."):
        fu = f.upper()
        if f.lower().endswith((".xlsx", ".xls")) and b[:4] in fu and (
            "STAT" in fu or "RELEV" in fu or "BANK" in fu
        ):
            return f
    return None


def lire_releve_excel(fichier, banque=""):
    try:
        xl = pd.ExcelFile(fichier)
    except Exception as e:
        raise ValueError(f"Impossible d'ouvrir le fichier : {e}")

    feuille = None
    for n in xl.sheet_names:
        if "STATEMENT" in n.upper() or "RELEVE" in n.upper() or "MOUVEMENT" in n.upper():
            feuille = n
            break
    if feuille is None:
        feuille = xl.sheet_names[0]

    brut = pd.read_excel(fichier, sheet_name=feuille, header=None)
    hdr = None
    for i in range(min(25, len(brut))):
        vals = [str(v).strip().upper() for v in brut.iloc[i].tolist()]
        if any(v == "DATE" for v in vals) and any("DESC" in v or "LIBELL" in v for v in vals):
            hdr = i
            break
    if hdr is None:
        raise ValueError("En-tête Date / DESCRIPTION introuvable dans le relevé Excel.")

    df = brut.iloc[hdr + 1:].copy()
    df.columns = [_s(v).upper() for v in brut.iloc[hdr].tolist()]
    colmap = {}
    for c in df.columns:
        cl = c.lower()
        if cl.startswith("date"):
            colmap[c] = "Date"
        elif "desc" in cl or "libell" in cl:
            colmap[c] = "Description"
        elif "credit" in cl or "crédit" in cl:
            colmap[c] = "Credit"
        elif "debit" in cl or "débit" in cl:
            colmap[c] = "Debit"
        elif "solde" in cl or "balance" in cl:
            colmap[c] = "Solde"
        elif "cheque" in cl or "chèque" in cl or "check" in cl:
            colmap[c] = "Num_Cheque"
        elif "motif" in cl:
            colmap[c] = "Motif"
        elif "categ" in cl:
            colmap[c] = "Categorie"
    df = df.rename(columns=colmap)
    for c in ["Date", "Description", "Credit", "Debit", "Solde", "Num_Cheque", "Motif", "Categorie"]:
        if c not in df.columns:
            df[c] = ""

    rows = []
    for _, r in df.iterrows():
        desc = _s(r["Description"])
        if not desc:
            continue
        up = desc.upper()
        if any(x in up for x in ("SOLDE INITIAL", "OP BALANCE", "OPENING", "SOLDE PRÉCÉDENT", "SOLDE PRECEDENT")):
            continue
        credit, debit = _n(r["Credit"]), _n(r["Debit"])
        if credit == 0 and debit == 0:
            continue
        rows.append({
            "Date": _d(r["Date"]),
            "Description": desc,
            "Credit": credit,
            "Debit": debit,
            "Solde": _n(r["Solde"]),
            "Num_Cheque": _s(r["Num_Cheque"]),
            "Motif": _s(r["Motif"]),
            "Categorie": _s(r["Categorie"]),
            "Banque": normaliser_banque(banque),
            "Source": os.path.basename(str(fichier)),
        })
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    return out[out["Date"].notna()].reset_index(drop=True)


def charger_releve_local(banque, date_debut=None, date_fin=None):
    path = _trouver_fichier_local(banque)
    if not path:
        return pd.DataFrame(), None
    df = lire_releve_excel(path, banque=banque)
    if date_debut is not None:
        df = df[df["Date"] >= pd.Timestamp(date_debut)]
    if date_fin is not None:
        df = df[df["Date"] <= pd.Timestamp(date_fin)]
    return df.reset_index(drop=True), path


# ---------- PDF ----------
def _extraire_texte_pdf(path_or_bytes):
    """Texte natif + OCR si pages image."""
    import pdfplumber

    ok_tess, info_tess = _configurer_ocr()
    poppler_path = _configurer_poppler()
    textes = []
    try:
        if isinstance(path_or_bytes, (bytes, bytearray)):
            src = BytesIO(path_or_bytes)
        else:
            src = path_or_bytes
        with pdfplumber.open(src) as pdf:
            for page in pdf.pages:
                t = page.extract_text() or ""
                if len(t.strip()) < 40:
                    if not ok_tess:
                        t = t or f"[OCR indisponible: {info_tess}]"
                    else:
                        try:
                            from pdf2image import convert_from_path, convert_from_bytes
                            import pytesseract

                            kwargs = {
                                "dpi": 300,
                                "first_page": page.page_number,
                                "last_page": page.page_number,
                                "fmt": "png",
                            }
                            if poppler_path:
                                kwargs["poppler_path"] = poppler_path
                            if isinstance(path_or_bytes, (bytes, bytearray)):
                                imgs = convert_from_bytes(path_or_bytes, **kwargs)
                            else:
                                imgs = convert_from_path(path_or_bytes, **kwargs)
                            if imgs:
                                # Plusieurs modes OCR pour relevés bancaires tabulaires
                                configs = [
                                    r"--oem 3 --psm 6",   # bloc uniforme
                                    r"--oem 3 --psm 4",   # colonne
                                    r"--oem 3 --psm 3",   # auto
                                ]
                                best = ""
                                for cfg in configs:
                                    try:
                                        cand = pytesseract.image_to_string(
                                            imgs[0], lang="fra+eng", config=cfg
                                        )
                                        if len(cand.strip()) > len(best.strip()):
                                            best = cand
                                    except Exception:
                                        continue
                                t = best or pytesseract.image_to_string(imgs[0], lang="fra+eng")
                        except Exception as e:
                            t = t or f"[OCR erreur: {e}]"
                textes.append(t)
    except Exception as e:
        return "", str(e)
    return "\n".join(textes), None


def _parse_money_amounts(line):
    """
    Extrait les montants monétaires d'une ligne de relevé.
    Priorité aux formats type 50,653,200.00 ou 5 000,00 ou 5000.00
    Ignore les pures références (ex. 262740020 sans séparateur).
    """
    line = (line or "").replace("\xa0", " ").replace(" ", " ")
    amounts = []
    # Format international banque: 1,234,567.89
    for m in re.finditer(r"(?<![\d])(\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?)(?![\d])", line):
        a = _n(m.group(1).replace(",", ""))
        if a >= 1:
            amounts.append(a)
    # Format FR: 1 234 567,89 ou 1234567,89
    for m in re.finditer(r"(?<![\d])(\d{1,3}(?:[ \u00a0]\d{3})+(?:,\d{1,2})?)(?![\d])", line):
        a = _n(m.group(1))
        if a >= 1:
            amounts.append(a)
    for m in re.finditer(r"(?<![\d])(\d+[.,]\d{2})(?![\d])", line):
        a = _n(m.group(1))
        if a >= 1 and a not in amounts:
            amounts.append(a)
    # Si rien: grands nombres avec espaces type 50 653 200
    if not amounts:
        for m in re.finditer(r"(?<![\d])(\d{1,3}(?:[ \u00a0.]\d{3}){1,})(?![\d.,])", line):
            a = _n(m.group(1))
            if a >= 1000:
                amounts.append(a)
    # Déduplique en gardant l'ordre
    seen, out = set(), []
    for a in amounts:
        k = round(a, 2)
        if k not in seen:
            seen.add(k)
            out.append(a)
    return out


def _parse_amounts_line(line):
    return _parse_money_amounts(line)



def _classify_credit_debit(up_line, amounts):
    """
    Reconstitue Crédit / Débit / Solde à partir des montants OCR.
    Ordre classique relevé FR imprimé : Débit | Crédit | Solde
    Ordre Excel ECOBANK NAP : Crédit | Débit | Solde
    """
    credit = debit = solde = 0.0
    credit_kw = (
        "VIREMENT RECU", "VERSEMENT", "CREDIT", "CRÉDIT", "REMISE DE CHEQUE",
        "REMISE CHQ", "REMISE", "APPROV", "ENCAISS", "SALAIRE", "SALES",
        "EXPORT", "LODERS", "BUNGE", "PAN TRADING", "INCOMING", "TRF FROM",
        "DEPOT", "DÉPOT", "DEPOSIT", "CR ", " C/R", "CLEARED", "RECEIVED",
        "TRANSFER IN", "VIR RECU", "VIR. RECU", "AVOIR",
    )
    debit_kw = (
        "RETRAIT", "PAYMENT", "PAIEMENT", "FRAIS", "AGIOS", "BFU", "GUCE",
        "SEGUB", "PURCHASE", "WITHDRAWAL", "COMMISSION", "ACHAT", "CHEQUE EMIS",
        "CHQ EMIS", "DR ", " D/R", "TRANSFER OUT", "VIR EMIS", "VIR. EMIS",
        "PRET REMB", "REMBOURSEMENT PRET", "COTISATION",
    )
    # Marqueurs explicites CR / DR en fin de montant
    has_cr_mark = bool(re.search(r"\bCR\b|\bCREDIT\b|\bCRÉDIT\b", up_line))
    has_dr_mark = bool(re.search(r"\bDR\b|\bDEBIT\b|\bDÉBIT\b", up_line))
    is_credit_txt = any(k in up_line for k in credit_kw)
    is_debit_txt = any(k in up_line for k in debit_kw)

    n = len(amounts)
    if n >= 3:
        # Cas colonnes Débit + Crédit + Solde (une des deux colonnes peut être 0)
        a, b, c = amounts[0], amounts[1], amounts[2]
        # Si un des deux premiers est ~0, l'autre est le mouvement
        if a > 0 and b == 0:
            debit, credit, solde = a, 0.0, c
        elif b > 0 and a == 0:
            debit, credit, solde = 0.0, b, c
        elif is_credit_txt and not is_debit_txt:
            # Excel-style CREDIT, DEBIT, SOLDE
            credit, debit, solde = a, b, c
        elif is_debit_txt and not is_credit_txt:
            debit, credit, solde = a, b, c
        else:
            # Défaut relevé papier FR : Débit, Crédit, Solde
            debit, credit, solde = a, b, c
            # Si les deux > 0 c'est rare sur une même ligne ; on garde les deux
    elif n == 2:
        mouvement, solde = amounts[0], amounts[1]
        if has_cr_mark or (is_credit_txt and not is_debit_txt):
            credit = mouvement
        elif has_dr_mark or is_debit_txt:
            debit = mouvement
        else:
            # Sans indice : on ne force plus tout en débit —
            # on met le mouvement en débit SEULEMENT si libellé opérationnel sortant,
            # sinon on expose aussi en crédit si solde a augmenté (impossible ici).
            # Règle prudente : débit par défaut SAUF mots crédit.
            if is_credit_txt:
                credit = mouvement
            else:
                debit = mouvement
    elif n == 1:
        mouvement = amounts[0]
        if has_cr_mark or (is_credit_txt and not is_debit_txt):
            credit = mouvement
        else:
            debit = mouvement
    return credit, debit, solde


def _ecobank_parse_line(line, banque):
    """
    Lignes OCR Ecobank :
    - REF // date XOF montant_debit montant_credit solde
    - REF // date XOF montant solde
    Lit TOUJOURS crédit et débit (plus seulement le débit).
    """
    up = line.upper()
    md = re.search(r"(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})", line)
    if not md:
        return None
    d = _d(md.group(1))
    if pd.isna(d):
        return None

    amounts = []
    mx = re.search(r"XOF\s*(.+)$", line, re.I)
    if mx:
        amounts = _parse_money_amounts(mx.group(1))
    if not amounts:
        # Montants à droite de la date
        after = line[md.end():]
        amounts = _parse_money_amounts(after)
    if not amounts:
        amounts = _parse_money_amounts(line)
    amounts = [a for a in amounts if 0 < a < 1e12]
    if not amounts:
        return None

    credit, debit, solde = _classify_credit_debit(up, amounts)
    if credit <= 0 and debit <= 0:
        return None

    desc = line
    desc = re.sub(r"\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}", " ", desc)
    desc = re.sub(r"XOF", " ", desc, flags=re.I)
    desc = re.sub(r"\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?", " ", desc)
    desc = re.sub(r"\d+[.,]\d{2}", " ", desc)
    desc = re.sub(r"[/]{2,}", " ", desc)
    desc = re.sub(r"\s+", " ", desc).strip(" -|/")
    if not desc:
        desc = line[:80]

    chq = ""
    mq = re.search(r"(?:CHQ|CHEQUE|CH[EÉ]QUE)[^0-9]{0,6}([0-9]{5,})", line, re.I)
    if mq:
        chq = mq.group(1)

    return {
        "Date": d,
        "Description": desc[:200],
        "Credit": float(credit or 0),
        "Debit": float(debit or 0),
        "Solde": float(solde or 0),
        "Num_Cheque": chq,
        "Motif": "",
        "Categorie": "",
        "Banque": banque,
        "Source": "PDF",
    }


def parser_texte_releve(texte, banque=""):
    rows = []
    banque = normaliser_banque(banque)
    for raw in (texte or "").splitlines():
        line = raw.strip()
        if not line or len(line) < 8:
            continue
        up = line.upper()
        if any(x in up for x in (
            "SOLDE PRÉCÉDENT", "SOLDE PRECEDENT", "SOLDE INITIAL", "OPENING BALANCE",
            "PAGE ", "EXTRAIT DE COMPTE", "RELEVE DE COMPTE", "RELEVÉ DE COMPTE",
            "TOTAL DES MOUVEMENTS", "NOMBRE DE TRANSACTIONS",
        )):
            continue

        # Ecobank (et OCR similaire) : parser dédié si XOF présent ou date au milieu
        if banque == "ECOBANK" or "XOF" in up:
            row = _ecobank_parse_line(line, banque)
            if row and (row["Credit"] > 0 or row["Debit"] > 0):
                rows.append(row)
            continue

        m = re.match(
            r"^(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}|\d{1,2}[-/\s][A-Za-z]{3,}[-/\s]\d{2,4})\s+(.*)$",
            line,
        )
        if not m:
            m = re.search(
                r"(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}|\d{1,2}[-/\s][A-Za-z]{3,}[-/\s]\d{2,4})\s+(.*)$",
                line,
            )
        if not m:
            continue
        d = _d(m.group(1))
        if pd.isna(d):
            continue
        rest = m.group(2)
        amounts = _parse_money_amounts(rest)
        amounts = [a for a in amounts if 0 < a < 1e12]
        if not amounts:
            continue
        credit, debit, solde = _classify_credit_debit(rest.upper(), amounts)

        chq = ""
        mq = re.search(
            r"(?:CHQ|CHEQUE|CH[EÉ]QUE|N[°º]?\s*CH[EÉ]?Q?[UE]*|CHQ\s*N)\s*[:\s]*([0-9]{5,})",
            rest, re.I,
        )
        if mq:
            chq = mq.group(1)

        desc = re.sub(r"(\d{1,3}(?:[ ,\u00a0.]\d{3})+(?:[.,]\d{2})?|\d+[.,]\d{2})", " ", rest)
        desc = re.sub(r"\s+", " ", desc).strip(" -|/")
        if not desc:
            desc = rest[:80]
        rows.append({
            "Date": d,
            "Description": desc[:200],
            "Credit": credit,
            "Debit": debit,
            "Solde": solde,
            "Num_Cheque": chq,
            "Motif": "",
            "Categorie": "",
            "Banque": banque,
            "Source": "PDF",
        })
    if not rows:
        return pd.DataFrame(columns=[
            "Date", "Description", "Credit", "Debit", "Solde", "Num_Cheque",
            "Motif", "Categorie", "Banque", "Source",
        ])
    df = pd.DataFrame(rows)
    # Correction par variation de solde : si le solde monte, le mouvement est un crédit
    try:
        df = df.sort_values("Date").reset_index(drop=True)
        prev = None
        for i in range(len(df)):
            solde = float(df.at[i, "Solde"] or 0)
            if solde > 0 and prev is not None and prev > 0:
                delta = solde - prev
                cr = float(df.at[i, "Credit"] or 0)
                db = float(df.at[i, "Debit"] or 0)
                if delta > 1 and db > 0 and cr == 0:
                    # Solde a augmenté → c'était un crédit mal classé
                    df.at[i, "Credit"] = db
                    df.at[i, "Debit"] = 0.0
                elif delta < -1 and cr > 0 and db == 0:
                    df.at[i, "Debit"] = cr
                    df.at[i, "Credit"] = 0.0
            if solde > 0:
                prev = solde
    except Exception:
        pass
    return df


def lire_releve_pdf(path_or_bytes, banque=""):
    texte, err = _extraire_texte_pdf(path_or_bytes)
    if err and not texte:
        raise ValueError(f"Lecture PDF impossible : {err}")
    df = parser_texte_releve(texte, banque=banque)
    return df, texte




def chemin_releve_import(banque):
    return f"{FICHIER_RELEVE_IMPORT_PREFIX}{normaliser_banque(banque) or 'AUTRE'}.xlsx"


def sauver_releve_import(df, banque):
    """Enregistre le relevé importé sur disque (persistant entre sessions)."""
    if df is None or df.empty:
        return False
    path = chemin_releve_import(banque)
    out = df.copy()
    out.to_excel(path, index=False)
    return True


def charger_releve_import(banque):
    """Charge le dernier relevé importé pour cette banque (fichier local)."""
    path = chemin_releve_import(banque)
    if not os.path.exists(path):
        return None, None
    try:
        df = pd.read_excel(path)
        if "Date" in df.columns:
            df["Date"] = df["Date"].map(_d)
        for c in ("Credit", "Debit", "Solde"):
            if c in df.columns:
                df[c] = df[c].map(_n)
        for c in ("Description", "Num_Cheque", "Motif", "Categorie", "Banque", "Source"):
            if c in df.columns:
                df[c] = df[c].map(_s)
        df["Banque"] = normaliser_banque(banque)
        return df, path
    except Exception:
        return None, None


def effacer_releve_import(banque):
    path = chemin_releve_import(banque)
    if os.path.exists(path):
        os.remove(path)
        return True
    return False




def alimenter_statement_depuis_suivi(banque, date_debut=None, date_fin=None):
    """
    Remplit / met à jour le fichier STATEMENT local de la banque
    à partir des lignes Suivi de paiement (DEPENSES) de cette banque.
    Retourne (ok, message, df_statement, chemin).
    """
    banque = normaliser_banque(banque)
    journal = charger_journal(banque, date_debut=date_debut, date_fin=date_fin)
    if journal.empty:
        return False, f"Aucune ligne Suivi de paiement pour {banque} sur la période.", pd.DataFrame(), None

    path = _trouver_fichier_local(banque)
    # Si pas de fichier, créer un nom standard
    if not path:
        noms = {
            "ECOBANK": "ECO BANK STATEMENT SHEET 2026.xlsx",
            "ATLANTIC": "ATLANTIC BANK STATEMENT SHEET 2026.xlsx",
            "BIIC": "BIIC BANK STATEMENT SHEET 2026.xlsx",
            "BGFI": "BGFI BANK STATEMENT 2026.xlsx",
        }
        path = noms.get(banque, f"{banque} BANK STATEMENT 2026.xlsx")

    rows = []
    for _, r in journal.iterrows():
        mt = _n(r.get("Montant"))
        mode = _s(r.get("Mode")).upper()
        # DEPENSES = sorties → Débit (sauf si type crédit exceptionnel)
        credit, debit = 0.0, abs(mt)
        desc_parts = [
            _s(r.get("Prestataire")),
            _s(r.get("Description")),
            _s(r.get("No_Transaction")),
        ]
        desc = " | ".join([x for x in desc_parts if x])
        rows.append({
            "Date": r.get("Date"),
            "DESCRIPTION": desc[:200],
            "CREDIT": credit,
            "DEBIT": debit,
            "SOLDE": "",
            "NUM DE CHEQUE": _s(r.get("Num_Cheque")),
            "MOTIF": _s(r.get("Observation")) or _s(r.get("Type")),
            "CATEGORIE": _s(r.get("Produit")) or _s(r.get("Type")),
        })
    df_new = pd.DataFrame(rows)
    if df_new.empty:
        return False, "Rien à écrire.", pd.DataFrame(), path

    # Écrire feuille STATEMENT (remplace le contenu mouvements générés ;
    # conserve une feuille META si besoin)
    try:
        # Charger l'existant pour éventuellement fusionner par n° transaction / chèque+montant+date
        existing = pd.DataFrame()
        if os.path.exists(path):
            try:
                existing = lire_releve_excel(path, banque=banque)
            except Exception:
                existing = pd.DataFrame()

        # Convertir existing au même format colonnes Excel
        out_rows = []
        seen = set()
        for _, r in df_new.iterrows():
            key = (
                str(r["Date"])[:10],
                _n(r["DEBIT"]),
                _n(r["CREDIT"]),
                _s(r["NUM DE CHEQUE"]),
                _s(r["DESCRIPTION"])[:40],
            )
            seen.add(key)
            out_rows.append(r.to_dict())

        # Garder les lignes existing qui ne viennent pas du suivi (ex. frais banque saisis à la main dans STATEMENT)
        if existing is not None and not existing.empty:
            for _, r in existing.iterrows():
                key = (
                    str(r.get("Date"))[:10],
                    _n(r.get("Debit")),
                    _n(r.get("Credit")),
                    _s(r.get("Num_Cheque")),
                    _s(r.get("Description"))[:40],
                )
                if key in seen:
                    continue
                # garder si source n'est pas purement suivi — on garde tout non dupliqué
                out_rows.append({
                    "Date": r.get("Date"),
                    "DESCRIPTION": _s(r.get("Description")),
                    "CREDIT": _n(r.get("Credit")),
                    "DEBIT": _n(r.get("Debit")),
                    "SOLDE": _n(r.get("Solde")) or "",
                    "NUM DE CHEQUE": _s(r.get("Num_Cheque")),
                    "MOTIF": _s(r.get("Motif")),
                    "CATEGORIE": _s(r.get("Categorie")),
                })

        df_out = pd.DataFrame(out_rows)
        # Format dates for Excel
        df_excel = df_out.copy()
        df_excel["Date"] = pd.to_datetime(df_excel["Date"], errors="coerce").dt.strftime("%Y-%m-%d")

        # Préserver autres feuilles
        other = {}
        if os.path.exists(path):
            try:
                xl = pd.ExcelFile(path)
                for sh in xl.sheet_names:
                    if sh.upper() not in ("STATEMENT", "RELEVE", "MOUVEMENT"):
                        other[sh] = pd.read_excel(path, sheet_name=sh)
            except Exception:
                pass

        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df_excel.to_excel(w, sheet_name="STATEMENT", index=False)
            for sh, sdf in other.items():
                try:
                    sdf.to_excel(w, sheet_name=str(sh)[:31], index=False)
                except Exception:
                    pass

        return True, f"{len(df_new)} ligne(s) Suivi → {path} (feuille STATEMENT).", df_out, path
    except Exception as e:
        return False, f"Écriture STATEMENT impossible : {e}", pd.DataFrame(), path


def charger_releve_pour_rapprochement(banque, date_debut=None, date_fin=None):
    """
    Priorité pour le côté 'interne / statement' :
    1) releve_importe (PDF banque déjà analysé) — NON, c'est le côté banque
    Pour le rapprochement :
    - côté journal = Suivi
    - côté relevé = PDF importé OU Excel STATEMENT local
    Cette fonction charge le STATEMENT local (après alimentation Suivi ou saisie).
    """
    return charger_releve_local(banque, date_debut=date_debut, date_fin=date_fin)



def charger_journal_complet(banque=None, date_debut=None, date_fin=None):
    """
    Journal 'chez nous' pour le rapprochement =
    débits (Suivi DEPENSES) + crédits (Credits_Bancaires_NAP.xlsx).
    Format unifié compatible avec rapprocher().
    """
    debits = charger_journal(banque=banque, date_debut=date_debut, date_fin=date_fin)

    # Crédits
    try:
        import credits_bancaires_data as cb
        if hasattr(cb, "charger_credits_tous"):
            credits = cb.charger_credits_tous(banque=banque, date_debut=date_debut, date_fin=date_fin)
        else:
            credits = cb.charger_credits(banque=banque, date_debut=date_debut, date_fin=date_fin)
    except Exception:
        credits = pd.DataFrame()

    rows = []
    if debits is not None and not debits.empty:
        for _, r in debits.iterrows():
            rows.append({
                "Date": r.get("Date"),
                "No_Transaction": _s(r.get("No_Transaction")),
                "Type": _s(r.get("Type")) or "DEBIT",
                "Produit": _s(r.get("Produit")),
                "Prestataire": _s(r.get("Prestataire")),
                "Beneficiaire": _s(r.get("Beneficiaire")),
                "Ref_Doc": _s(r.get("Ref_Doc")),
                "Description": _s(r.get("Description")),
                "Montant": abs(_n(r.get("Montant"))),
                "Observation": _s(r.get("Observation")),
                "Mode": _s(r.get("Mode")),
                "Num_Cheque": _s(r.get("Num_Cheque")),
                "Banque": _s(r.get("Banque")) or normaliser_banque(banque),
                "Sens": "DEBIT",
            })
    if credits is not None and not credits.empty:
        for _, r in credits.iterrows():
            rows.append({
                "Date": r.get("Date"),
                "No_Transaction": _s(r.get("Ref")),
                "Type": _s(r.get("Type")) or "CREDIT",
                "Produit": "",
                "Prestataire": _s(r.get("Tiers")),
                "Beneficiaire": _s(r.get("Tiers")),
                "Ref_Doc": _s(r.get("Lien")),
                "Description": f"{_s(r.get('Type'))} — {_s(r.get('Motif'))}".strip(" —"),
                "Montant": abs(_n(r.get("Montant"))),
                "Observation": _s(r.get("Observation")),
                "Mode": _s(r.get("Mode")),
                "Num_Cheque": _s(r.get("Piece")),
                "Banque": _s(r.get("Banque")) or normaliser_banque(banque),
                "Sens": "CREDIT",
            })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["Date"] = df["Date"].map(_d)
    return df.sort_values("Date").reset_index(drop=True)


def rapprocher_complet(releve, journal_complet, tolerance_jours=3, tolerance_montant=1.0):
    """
    Rapprochement relevé banque ↔ (débits + crédits).
    Matching sur montant + date + n° chèque ; le sens (crédit/débit) est respecté
    quand le relevé a Credit/Debit renseignés.
    """
    rev = releve.copy() if releve is not None else pd.DataFrame()
    jou = journal_complet.copy() if journal_complet is not None else pd.DataFrame()

    if not rev.empty:
        rev = rev.copy()
        rev["_mt"] = rev.apply(
            lambda r: _n(r.get("Credit")) if _n(r.get("Credit")) > 0 else _n(r.get("Debit")),
            axis=1,
        )
        rev["_sens"] = rev.apply(
            lambda r: "CREDIT" if _n(r.get("Credit")) > 0 else "DEBIT",
            axis=1,
        )
        rev = rev[rev["_mt"] > 0].reset_index(drop=True)

    if not jou.empty:
        jou = jou.copy()
        jou["_mt"] = jou["Montant"].map(_n).abs()
        if "Sens" not in jou.columns:
            jou["Sens"] = "DEBIT"
        jou["_sens"] = jou["Sens"].map(lambda x: "CREDIT" if _s(x).upper() == "CREDIT" else "DEBIT")

    used_j = set()
    matches, banque_seul = [], []

    for i, r in rev.iterrows():
        mt, sens, d0 = r["_mt"], r["_sens"], r["Date"]
        chq = _s(r.get("Num_Cheque"))
        best, best_score = None, 999
        for j, jrow in jou.iterrows():
            if j in used_j:
                continue
            if abs(jrow["_mt"] - mt) > tolerance_montant:
                continue
            # préférer même sens
            sens_ok = (jrow["_sens"] == sens)
            jchq = _s(jrow.get("Num_Cheque"))
            if chq and jchq and re.sub(r"\D", "", chq) == re.sub(r"\D", "", jchq):
                best, best_score = j, -1
                break
            if pd.isna(d0) or pd.isna(jrow["Date"]):
                score = 10
            else:
                score = abs((d0 - jrow["Date"]).days)
            if score <= tolerance_jours and score < best_score:
                # bonus si même sens
                adj = score if sens_ok else score + 0.5
                if adj < best_score or (best is not None and not sens_ok and sens_ok):
                    best, best_score = j, score if sens_ok else score + 0.5
            elif score <= tolerance_jours and best is None:
                best, best_score = j, score + (0 if sens_ok else 0.5)
        if best is not None and best_score <= tolerance_jours + 0.5:
            used_j.add(best)
            jrow = jou.loc[best]
            matches.append({
                "Date_Releve": r["Date"],
                "Desc_Releve": r.get("Description", ""),
                "Montant_Releve": mt,
                "Sens": sens,
                "Num_Cheque": chq,
                "Date_Journal": jrow["Date"],
                "Desc_Journal": f"{jrow.get('Prestataire', '')} | {jrow.get('Description', '')}".strip(" |"),
                "Montant_Journal": jrow["_mt"],
                "Sens_Journal": jrow["_sens"],
                "Ref_Journal": _s(jrow.get("No_Transaction")),
                "Ecart_Jours": int(max(best_score, 0)),
                "Statut": "RAPPROCHÉ",
            })
        else:
            banque_seul.append({
                "Date": r["Date"],
                "Description": r.get("Description", ""),
                "Credit": r.get("Credit", 0),
                "Debit": r.get("Debit", 0),
                "Montant": mt,
                "Sens": sens,
                "Num_Cheque": chq,
                "Motif": _s(r.get("Motif")),
                "Categorie": _s(r.get("Categorie")),
                "Statut": "BANQUE SEUL — à vérifier",
            })

    journal_seul = []
    for j, jrow in jou.iterrows():
        if j in used_j:
            continue
        journal_seul.append({
            "Date": jrow["Date"],
            "No_Transaction": _s(jrow.get("No_Transaction")),
            "Prestataire": _s(jrow.get("Prestataire")),
            "Description": _s(jrow.get("Description")),
            "Montant": jrow["_mt"],
            "Sens": jrow["_sens"],
            "Num_Cheque": _s(jrow.get("Num_Cheque")),
            "Mode": _s(jrow.get("Mode")),
            "Banque": _s(jrow.get("Banque")),
            "Statut": "CHEZ NOUS SEUL — pas (encore) au relevé",
        })

    return pd.DataFrame(matches), pd.DataFrame(banque_seul), pd.DataFrame(journal_seul)




def generer_pdf_rapprochement_a3(banque, date_debut, date_fin, matches, banque_seul, journal_seul):
    """Rapport A3 paysage du rapprochement bancaire."""
    from datetime import datetime
    from io import BytesIO
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A3),
        leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm,
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle("t", parent=styles["Heading1"], alignment=TA_CENTER, fontSize=15, spaceAfter=4)
    sub = ParagraphStyle("s", parent=styles["Normal"], alignment=TA_CENTER, fontSize=9, textColor=colors.HexColor("#444"))
    small = ParagraphStyle("sm", parent=styles["Normal"], fontSize=6.5, leading=8)
    h2 = ParagraphStyle("h2", parent=styles["Heading2"], fontSize=11, spaceBefore=6, spaceAfter=4, textColor=colors.HexColor("#1F4E79"))

    story = []
    story.append(Paragraph("NAP SARL — RAPPORT DE RAPPROCHEMENT BANCAIRE", title))
    story.append(Paragraph(
        f"Banque : <b>{banque}</b> &nbsp;|&nbsp; Période : <b>{date_debut.strftime('%d/%m/%Y')} → {date_fin.strftime('%d/%m/%Y')}</b> "
        f"&nbsp;|&nbsp; Édité le {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        sub,
    ))
    story.append(Spacer(1, 4*mm))

    n_ok = len(matches) if matches is not None else 0
    n_b = len(banque_seul) if banque_seul is not None else 0
    n_j = len(journal_seul) if journal_seul is not None else 0
    kpi = [
        ["Rapprochés", "Banque seulement", "Chez nous seulement"],
        [str(n_ok), str(n_b), str(n_j)],
    ]
    tk = Table(kpi, colWidths=[90*mm, 90*mm, 90*mm])
    tk.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BACKGROUND", (0, 1), (0, 1), colors.HexColor("#D5F5E3")),
        ("BACKGROUND", (1, 1), (1, 1), colors.HexColor("#FADBD8")),
        ("BACKGROUND", (2, 1), (2, 1), colors.HexColor("#FCF3CF")),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(tk)
    story.append(Spacer(1, 4*mm))

    def _fmt_d(v):
        try:
            return pd.to_datetime(v).strftime("%d/%m/%Y")
        except Exception:
            return _s(v)[:10]

    def _tbl(title_txt, df, cols, widths):
        story.append(Paragraph(title_txt, h2))
        if df is None or df.empty:
            story.append(Paragraph("Aucune ligne.", small))
            return
        data = [cols]
        for _, r in df.iterrows():
            row = []
            for c in cols:
                v = r.get(c, "")
                if "Date" in c or c == "Date":
                    row.append(_fmt_d(v))
                elif c in ("Montant", "Montant_Releve", "Montant_Journal"):
                    row.append(fmt(v) if _n(v) else "")
                else:
                    row.append(Paragraph(_s(v)[:60], small))
            data.append(row)
        t = Table(data, colWidths=widths, repeatRows=1)
        cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 7),
            ("FONTSIZE", (0, 1), (-1, -1), 6.5),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CCC")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]
        for i in range(1, len(data)):
            if i % 2 == 0:
                cmds.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F8F9F9")))
        t.setStyle(TableStyle(cmds))
        story.append(t)

    # Banque seulement
    cols_b = [c for c in ["Date", "Description", "Credit", "Debit", "Montant", "Sens", "Num_Cheque", "Statut"] if banque_seul is not None and not banque_seul.empty and c in banque_seul.columns]
    if not cols_b:
        cols_b = ["Date", "Description", "Montant", "Sens", "Statut"]
    w_b = [22*mm, 70*mm] + [28*mm] * (len(cols_b) - 2)
    _tbl("1. Mouvements sur le relevé banque seulement", banque_seul, cols_b, w_b[:len(cols_b)])

    cols_j = [c for c in ["Date", "No_Transaction", "Prestataire", "Description", "Montant", "Sens", "Num_Cheque", "Mode", "Statut"] if journal_seul is not None and not journal_seul.empty and c in journal_seul.columns]
    if not cols_j:
        cols_j = ["Date", "Description", "Montant", "Sens", "Statut"]
    w_j = [22*mm, 28*mm, 35*mm, 55*mm] + [22*mm] * max(0, len(cols_j) - 4)
    _tbl("2. Mouvements chez nous seulement (absents du relevé)", journal_seul, cols_j, w_j[:len(cols_j)])

    cols_m = [c for c in ["Date_Releve", "Desc_Releve", "Montant_Releve", "Sens", "Date_Journal", "Desc_Journal", "Montant_Journal", "Ref_Journal", "Ecart_Jours"] if matches is not None and not matches.empty and c in matches.columns]
    if not cols_m:
        cols_m = ["Date_Releve", "Montant_Releve", "Date_Journal", "Montant_Journal"]
    w_m = [22*mm, 50*mm, 25*mm, 18*mm, 22*mm, 50*mm, 25*mm, 25*mm, 15*mm]
    _tbl("3. Lignes rapprochées", matches, cols_m, w_m[:len(cols_m)])

    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(
        "NAP SARL — Rapport interne de rapprochement. Chez nous = Suivi (débits) + crédits (base / VERSEMENTS / relevé). "
        "À conserver pour le contrôle de trésorerie.",
        ParagraphStyle("f", parent=styles["Normal"], fontSize=7, textColor=colors.grey, alignment=TA_CENTER),
    ))
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


def statut_ocr():
    """Info pour l'UI : OCR prêt ou non."""
    ok, info = _configurer_ocr()
    pop = _configurer_poppler()
    return {
        "tesseract_ok": ok,
        "tesseract_info": info,
        "poppler_path": pop or "(non trouvé — pdf2image peut échouer sur Windows)",
    }


def filtrer_periode(df, date_debut=None, date_fin=None):
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    out = df.copy()
    if "Date" in out.columns:
        out["Date"] = out["Date"].map(_d)
        if date_debut is not None:
            out = out[out["Date"] >= pd.Timestamp(date_debut)]
        if date_fin is not None:
            out = out[out["Date"] <= pd.Timestamp(date_fin)]
    return out.reset_index(drop=True)


# ---------- Matching ----------
def rapprocher(releve, journal, tolerance_jours=3, tolerance_montant=1.0):
    rev = releve.copy() if releve is not None else pd.DataFrame()
    jou = journal.copy() if journal is not None else pd.DataFrame()

    if not rev.empty:
        rev = rev.copy()
        rev["_mt"] = rev.apply(
            lambda r: _n(r.get("Credit")) if _n(r.get("Credit")) > 0 else _n(r.get("Debit")),
            axis=1,
        )
        rev["_sens"] = rev.apply(
            lambda r: "CREDIT" if _n(r.get("Credit")) > 0 else "DEBIT",
            axis=1,
        )
        rev = rev[rev["_mt"] > 0].reset_index(drop=True)

    if not jou.empty:
        jou = jou.copy()
        jou["_mt"] = jou["Montant"].map(_n).abs()
        jou["_sens"] = "DEBIT"

    used_j = set()
    matches, banque_seul = [], []

    for i, r in rev.iterrows():
        mt, sens, d0 = r["_mt"], r["_sens"], r["Date"]
        chq = _s(r.get("Num_Cheque"))
        best, best_score = None, 999
        for j, jrow in jou.iterrows():
            if j in used_j:
                continue
            if abs(jrow["_mt"] - mt) > tolerance_montant:
                continue
            jchq = _s(jrow.get("Num_Cheque"))
            if chq and jchq and re.sub(r"\D", "", chq) == re.sub(r"\D", "", jchq):
                best, best_score = j, -1
                break
            if pd.isna(d0) or pd.isna(jrow["Date"]):
                score = 10
            else:
                score = abs((d0 - jrow["Date"]).days)
            if score <= tolerance_jours and score < best_score:
                best, best_score = j, score
        if best is not None:
            used_j.add(best)
            jrow = jou.loc[best]
            matches.append({
                "Date_Releve": r["Date"],
                "Desc_Releve": r["Description"],
                "Montant_Releve": mt,
                "Sens": sens,
                "Num_Cheque": chq,
                "Date_Journal": jrow["Date"],
                "Desc_Journal": f"{jrow.get('Prestataire', '')} | {jrow.get('Description', '')}".strip(" |"),
                "Montant_Journal": jrow["_mt"],
                "Ref_Journal": _s(jrow.get("No_Transaction")),
                "Ecart_Jours": max(best_score, 0),
                "Statut": "RAPPROCHÉ",
            })
        else:
            banque_seul.append({
                "Date": r["Date"],
                "Description": r["Description"],
                "Credit": r.get("Credit", 0),
                "Debit": r.get("Debit", 0),
                "Montant": mt,
                "Sens": sens,
                "Num_Cheque": chq,
                "Motif": _s(r.get("Motif")),
                "Categorie": _s(r.get("Categorie")),
                "Statut": "BANQUE SEUL — à vérifier",
            })

    journal_seul = []
    for j, jrow in jou.iterrows():
        if j in used_j:
            continue
        journal_seul.append({
            "Date": jrow["Date"],
            "No_Transaction": _s(jrow.get("No_Transaction")),
            "Prestataire": _s(jrow.get("Prestataire")),
            "Description": _s(jrow.get("Description")),
            "Montant": jrow["_mt"],
            "Num_Cheque": _s(jrow.get("Num_Cheque")),
            "Mode": _s(jrow.get("Mode")),
            "Banque": _s(jrow.get("Banque")),
            "Statut": "JOURNAL SEUL — pas (encore) au relevé",
        })

    return pd.DataFrame(matches), pd.DataFrame(banque_seul), pd.DataFrame(journal_seul)


# ---------- Manuel ----------
def charger_manuelles(banque=None):
    if not os.path.exists(FICHIER_MANUEL):
        return pd.DataFrame(columns=[
            "ID", "Banque", "Date", "Description", "Credit", "Debit",
            "Num_Cheque", "Motif", "Categorie", "Cree_Par", "Heure",
        ])
    df = pd.read_excel(FICHIER_MANUEL)
    df["Credit"] = df.get("Credit", 0).map(_n) if "Credit" in df.columns else 0
    df["Debit"] = df.get("Debit", 0).map(_n) if "Debit" in df.columns else 0
    df["Date"] = df["Date"].map(_d) if "Date" in df.columns else pd.NaT
    df["Banque"] = df["Banque"].map(normaliser_banque) if "Banque" in df.columns else ""
    if banque:
        df = df[df["Banque"] == normaliser_banque(banque)]
    return df.reset_index(drop=True)


def ajouter_manuel(data, user=""):
    df = charger_manuelles()
    eid = 1 if df.empty else int(pd.to_numeric(df.get("ID", 0), errors="coerce").fillna(0).max()) + 1
    row = {
        "ID": eid,
        "Banque": normaliser_banque(data.get("Banque")),
        "Date": _d(data.get("Date")),
        "Description": _s(data.get("Description")),
        "Credit": _n(data.get("Credit")),
        "Debit": _n(data.get("Debit")),
        "Num_Cheque": _s(data.get("Num_Cheque")),
        "Motif": _s(data.get("Motif")),
        "Categorie": _s(data.get("Categorie")),
        "Cree_Par": _s(user),
        "Heure": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    if row["Credit"] <= 0 and row["Debit"] <= 0:
        return False, "Saisir un montant crédit ou débit."
    if not row["Description"]:
        return False, "Description obligatoire."
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_excel(FICHIER_MANUEL, index=False)
    return True, f"Écriture #{eid} enregistrée."


def to_excel_bytes(feuilles: dict):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nom, df in feuilles.items():
            d = df.copy() if df is not None else pd.DataFrame()
            for c in list(d.columns):
                if "Date" in str(c) or c == "Date":
                    d[c] = pd.to_datetime(d[c], errors="coerce").dt.strftime("%d/%m/%Y")
            d.to_excel(w, sheet_name=str(nom)[:31], index=False)
    buf.seek(0)
    return buf.getvalue()
