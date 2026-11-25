# -*- coding: utf-8 -*-
"""Crédits bancaires + rapport débit/crédit + solde."""
from __future__ import annotations

import os
import re
from datetime import datetime, date
from io import BytesIO

import pandas as pd

FICHIER_CREDITS = "Credits_Bancaires_NAP.xlsx"
FEUILLE_CREDITS = "CREDITS"
FEUILLE_SOLDES = "SOLDES_DEPART"
FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"

BANQUES = ["ECOBANK", "ATLANTIC", "BIIC", "BGFI", "UBA", "CAISSE", "AUTRE"]
TYPES_CREDIT = [
    "VERSEMENT ESPECES", "VIREMENT RECU", "PRET BANCAIRE",
    "REMBOURSEMENT CLIENT", "ENCAISSEMENT CHEQUE", "AUTRE CREDIT",
]


def _s(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
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
    return pd.to_datetime(v, errors="coerce", dayfirst=True)


def fmt(n):
    try:
        return f"{float(n):,.0f}".replace(",", " ")
    except Exception:
        return str(n)


def normaliser_banque(x):
    t = _s(x).upper()
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
    return t or "AUTRE"


def _trouver_credits():
    for c in [FICHIER_CREDITS, "Credits_Bancaires.xlsx", "credits_bancaires.xlsx"]:
        if os.path.exists(c):
            return c
    return FICHIER_CREDITS


def charger_credits(banque=None, date_debut=None, date_fin=None):
    path = _trouver_credits()
    cols = [
        "Date", "Ref", "Banque", "Type", "Tiers", "Motif", "Montant",
        "Mode", "Piece", "Lien", "Observation",
    ]
    if not os.path.exists(path):
        return pd.DataFrame(columns=cols)

    try:
        df = pd.read_excel(path, sheet_name=FEUILLE_CREDITS)
    except Exception:
        df = pd.read_excel(path, sheet_name=0)

    alias = {
        "date": "Date", "n° référence": "Ref", "n° reference": "Ref", "ref": "Ref",
        "banque": "Banque", "type crédit": "Type", "type credit": "Type", "type": "Type",
        "tiers / déposant": "Tiers", "tiers / deposant": "Tiers", "tiers": "Tiers",
        "motif / description": "Motif", "motif": "Motif", "description": "Motif",
        "montant": "Montant", "mode": "Mode",
        "n° pièce / chèque": "Piece", "n° piece": "Piece", "piece": "Piece",
        "lien prêt / dossier": "Lien", "lien": "Lien", "observation": "Observation",
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
    df["Montant"] = df["Montant"].map(_n)
    df["Date"] = df["Date"].map(_d)
    df["Banque"] = df["Banque"].map(normaliser_banque)
    for c in ["Ref", "Type", "Tiers", "Motif", "Mode", "Piece", "Lien", "Observation"]:
        df[c] = df[c].map(_s)
    df = df[df["Montant"] > 0].copy()
    if banque:
        df = df[df["Banque"] == normaliser_banque(banque)]
    if date_debut is not None:
        df = df[df["Date"] >= pd.Timestamp(date_debut)]
    if date_fin is not None:
        df = df[df["Date"] <= pd.Timestamp(date_fin)]
    return df.reset_index(drop=True)


def charger_soldes_depart():
    path = _trouver_credits()
    if not os.path.exists(path):
        return {b: 0.0 for b in BANQUES}
    try:
        df = pd.read_excel(path, sheet_name=FEUILLE_SOLDES)
    except Exception:
        return {b: 0.0 for b in BANQUES}
    out = {b: 0.0 for b in BANQUES}
    if df is None or df.empty:
        return out
    col_b = next((c for c in df.columns if "banque" in str(c).lower()), df.columns[0])
    col_s = next((c for c in df.columns if "solde" in str(c).lower() or "ouvert" in str(c).lower()), df.columns[1] if len(df.columns) > 1 else None)
    if col_s is None:
        return out
    for _, r in df.iterrows():
        b = normaliser_banque(r[col_b])
        out[b] = _n(r[col_s])
    return out


def prochain_ref_credit():
    df = charger_credits()
    year = datetime.now().year
    n = 1
    if not df.empty and "Ref" in df.columns:
        nums = []
        for v in df["Ref"]:
            m = re.search(r"(\d+)", _s(v))
            if m:
                nums.append(int(m.group(1)))
        if nums:
            n = max(nums) + 1
    return f"CR-{n:04d}-{year}"


def ajouter_credit(data):
    path = _trouver_credits()
    cols_fr = [
        "Date", "N° Référence", "Banque", "Type crédit", "Tiers / Déposant",
        "Motif / Description", "Montant", "Mode", "N° Pièce / Chèque",
        "Lien prêt / dossier", "Observation",
    ]
    if os.path.exists(path):
        try:
            df = pd.read_excel(path, sheet_name=FEUILLE_CREDITS)
        except Exception:
            df = pd.DataFrame(columns=cols_fr)
        # autres feuilles
        other = {}
        try:
            xl = pd.ExcelFile(path)
            for sh in xl.sheet_names:
                if sh != FEUILLE_CREDITS:
                    other[sh] = pd.read_excel(path, sheet_name=sh)
        except Exception:
            pass
    else:
        df = pd.DataFrame(columns=cols_fr)
        other = {}

    row = {
        "Date": data.get("Date"),
        "N° Référence": _s(data.get("Ref")) or prochain_ref_credit(),
        "Banque": normaliser_banque(data.get("Banque")),
        "Type crédit": _s(data.get("Type")),
        "Tiers / Déposant": _s(data.get("Tiers")),
        "Motif / Description": _s(data.get("Motif")),
        "Montant": _n(data.get("Montant")),
        "Mode": _s(data.get("Mode")),
        "N° Pièce / Chèque": _s(data.get("Piece")),
        "Lien prêt / dossier": _s(data.get("Lien")),
        "Observation": _s(data.get("Observation")),
    }
    if row["Montant"] <= 0:
        return False, "Montant crédit obligatoire."
    if not row["Banque"]:
        return False, "Banque obligatoire."

    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=FEUILLE_CREDITS, index=False)
        for sh, sdf in other.items():
            try:
                sdf.to_excel(w, sheet_name=str(sh)[:31], index=False)
            except Exception:
                pass
    return True, f"Crédit {row['N° Référence']} enregistré ({fmt(row['Montant'])} FCFA)."




def charger_credits_versements_suivi(banque=None, date_debut=None, date_fin=None):
    """Crédits issus de la feuille VERSEMENTS du Suivi de paiement."""
    path = None
    for c in [FICHIER_PAIEMENTS, "Suivi de paiement-5.xlsx", "Suivi de paiement-4.xlsx", "Suivi de paiement.xlsx"]:
        if os.path.exists(c):
            path = c
            break
    if not path:
        return pd.DataFrame(columns=["Date", "Ref", "Banque", "Type", "Tiers", "Motif", "Montant", "Mode", "Piece", "Lien", "Observation", "Sens"])

    try:
        xl = pd.ExcelFile(path)
        if "VERSEMENTS" not in xl.sheet_names:
            return pd.DataFrame()
        raw = pd.read_excel(path, sheet_name="VERSEMENTS", header=None)
    except Exception:
        return pd.DataFrame()

    hdr = None
    for i in range(min(10, len(raw))):
        vals = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        if any("date" == v or v == "date" for v in vals) and any("montant" in v for v in vals):
            hdr = i
            break
    if hdr is None:
        return pd.DataFrame()

    df = pd.read_excel(path, sheet_name="VERSEMENTS", header=hdr)
    alias = {}
    for c in df.columns:
        k = str(c).strip().lower()
        if k == "date":
            alias[c] = "Date"
        elif "doc" in k and "ref" in k:
            alias[c] = "Ref"
        elif "deposant" in k or "déposant" in k or "nom" in k:
            alias[c] = "Tiers"
        elif "motif" in k:
            alias[c] = "Motif"
        elif "montant" in k:
            alias[c] = "Montant"
        elif "banque" in k:
            alias[c] = "Banque"
    df = df.rename(columns=alias)
    for c in ["Date", "Ref", "Tiers", "Motif", "Montant", "Banque"]:
        if c not in df.columns:
            df[c] = ""
    df["Date"] = df["Date"].map(_d)
    df["Montant"] = df["Montant"].map(_n)
    df["Banque"] = df["Banque"].map(normaliser_banque)
    df["Ref"] = df["Ref"].map(_s)
    df["Tiers"] = df["Tiers"].map(_s)
    df["Motif"] = df["Motif"].map(_s)
    df["Type"] = "VERSEMENT (Suivi)"
    df["Mode"] = ""
    df["Piece"] = df["Ref"]
    df["Lien"] = ""
    df["Observation"] = "Source: Suivi VERSEMENTS"
    df["Sens"] = "CREDIT"
    df = df[df["Montant"] > 0].copy()
    if banque:
        df = df[df["Banque"] == normaliser_banque(banque)]
    if date_debut is not None:
        df = df[df["Date"] >= pd.Timestamp(date_debut)]
    if date_fin is not None:
        df = df[df["Date"] <= pd.Timestamp(date_fin)]
    return df.reset_index(drop=True)




def charger_credits_releve_importe(banque=None, date_debut=None, date_fin=None):
    """
    Crédits détectés sur le relevé banque importé dans le module rapprochement
    (fichier releve_importe_BANQUE.xlsx, colonne Credit > 0).
    """
    banques = [banque] if banque else ["ECOBANK", "ATLANTIC", "BIIC", "BGFI", "UBA", "CAISSE", "AUTRE"]
    rows = []
    for b in banques:
        bnorm = normaliser_banque(b) if b else ""
        path = f"releve_importe_{bnorm or 'AUTRE'}.xlsx"
        if not os.path.exists(path):
            # variante
            for cand in os.listdir(".") if os.path.isdir(".") else []:
                if cand.lower().startswith("releve_importe") and cand.lower().endswith(".xlsx"):
                    if bnorm and bnorm.lower() not in cand.lower():
                        continue
                    path = cand
                    break
            else:
                if not os.path.exists(path):
                    continue
        try:
            df = pd.read_excel(path)
        except Exception:
            continue
        if df is None or df.empty:
            continue
        # normalize columns
        colmap = {}
        for c in df.columns:
            k = str(c).strip().lower()
            if k == "date":
                colmap[c] = "Date"
            elif k in ("credit", "crédit", "montant credit"):
                colmap[c] = "Credit"
            elif k in ("debit", "débit"):
                colmap[c] = "Debit"
            elif "desc" in k or "libelle" in k or "libellé" in k:
                colmap[c] = "Description"
            elif "cheque" in k or "chèque" in k:
                colmap[c] = "Num_Cheque"
            elif k == "banque":
                colmap[c] = "Banque"
            elif k == "solde":
                colmap[c] = "Solde"
        df = df.rename(columns=colmap)
        if "Credit" not in df.columns:
            continue
        df["Credit"] = df["Credit"].map(_n)
        df = df[df["Credit"] > 0].copy()
        if df.empty:
            continue
        if "Date" in df.columns:
            df["Date"] = df["Date"].map(_d)
        else:
            df["Date"] = pd.NaT
        if date_debut is not None:
            df = df[df["Date"] >= pd.Timestamp(date_debut)]
        if date_fin is not None:
            df = df[df["Date"] <= pd.Timestamp(date_fin)]
        for _, r in df.iterrows():
            rows.append({
                "Date": r.get("Date"),
                "Ref": _s(r.get("Num_Cheque")) or "",
                "Banque": normaliser_banque(r.get("Banque") or bnorm),
                "Type": "CREDIT RELEVE BANQUE",
                "Tiers": "",
                "Motif": _s(r.get("Description"))[:200],
                "Montant": abs(_n(r.get("Credit"))),
                "Mode": "",
                "Piece": _s(r.get("Num_Cheque")),
                "Lien": "",
                "Observation": "Source: relevé importé (rapprochement)",
                "Source_Credit": "RELEVE_BANQUE",
            })
    if not rows:
        return pd.DataFrame(columns=[
            "Date", "Ref", "Banque", "Type", "Tiers", "Motif", "Montant",
            "Mode", "Piece", "Lien", "Observation", "Source_Credit",
        ])
    return pd.DataFrame(rows)



def charger_credits_tous(banque=None, date_debut=None, date_fin=None):
    """
    Union des crédits :
    1) Credits_Bancaires_NAP.xlsx (saisie manuelle)
    2) VERSEMENTS du Suivi de paiement
    3) Colonne Credit du relevé banque importé (module rapprochement)
    """
    parts = []
    for loader in (charger_credits, charger_credits_versements_suivi, charger_credits_releve_importe):
        try:
            df = loader(banque, date_debut, date_fin)
            if df is not None and not df.empty:
                parts.append(df)
        except Exception:
            continue
    if not parts:
        return pd.DataFrame(columns=[
            "Date", "Ref", "Banque", "Type", "Tiers", "Motif", "Montant",
            "Mode", "Piece", "Lien", "Observation", "Source_Credit",
        ])
    cols = ["Date", "Ref", "Banque", "Type", "Tiers", "Motif", "Montant", "Mode", "Piece", "Lien", "Observation"]
    aligned = []
    for df in parts:
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        if "Source_Credit" not in df.columns:
            df["Source_Credit"] = ""
        aligned.append(df[cols + ["Source_Credit"]])
    out = pd.concat(aligned, ignore_index=True)
    out["Montant"] = out["Montant"].map(_n)
    out = out[out["Montant"] > 0]
    out["_k"] = (
        out["Date"].astype(str).str[:10] + "|" +
        out["Montant"].map(lambda x: f"{_n(x):.0f}") + "|" +
        out["Banque"].map(_s)
    )
    out = out.drop_duplicates(subset=["_k"], keep="first").drop(columns=["_k"])
    return out.sort_values("Date").reset_index(drop=True)


def charger_debits_suivi(banque=None, date_debut=None, date_fin=None):
    """Débits depuis Suivi de paiement DEPENSES."""
    try:
        import bank_reconciliation_data as br
        return br.charger_journal(banque=banque, date_debut=date_debut, date_fin=date_fin)
    except Exception:
        pass
    # fallback minimal
    path = None
    for c in [FICHIER_PAIEMENTS, "Suivi de paiement-5.xlsx"]:
        if os.path.exists(c):
            path = c
            break
    if not path:
        return pd.DataFrame()
    try:
        xl = pd.ExcelFile(path)
        feuille = "DEPENSES" if "DEPENSES" in xl.sheet_names else xl.sheet_names[0]
        df = pd.read_excel(path, sheet_name=feuille)
    except Exception:
        return pd.DataFrame()
    # light normalize
    colmap = {}
    for c in df.columns:
        k = str(c).lower()
        if k == "date":
            colmap[c] = "Date"
        elif "montant" in k:
            colmap[c] = "Montant"
        elif "banque" in k:
            colmap[c] = "Banque"
        elif "transaction" in k:
            colmap[c] = "No_Transaction"
        elif "description" in k or "objet" in k:
            colmap[c] = "Description"
        elif "prestataire" in k or "fournisseur" in k:
            colmap[c] = "Prestataire"
        elif "cheque" in k or "chèque" in k:
            colmap[c] = "Num_Cheque"
    df = df.rename(columns=colmap)
    if "Montant" not in df.columns:
        return pd.DataFrame()
    df["Montant"] = df["Montant"].map(_n)
    df["Date"] = df["Date"].map(_d) if "Date" in df.columns else pd.NaT
    df["Banque"] = df["Banque"].map(normaliser_banque) if "Banque" in df.columns else ""
    df = df[df["Montant"] != 0]
    if banque:
        df = df[df["Banque"] == normaliser_banque(banque)]
    if date_debut is not None:
        df = df[df["Date"] >= pd.Timestamp(date_debut)]
    if date_fin is not None:
        df = df[df["Date"] <= pd.Timestamp(date_fin)]
    return df.reset_index(drop=True)


def construire_mouvements(banque, date_debut, date_fin):
    """Liste unique débit + crédit pour rapport."""
    debits = charger_debits_suivi(banque, date_debut, date_fin)
    credits = charger_credits_tous(banque, date_debut, date_fin)
    rows = []
    for _, r in debits.iterrows():
        rows.append({
            "Date": r.get("Date"),
            "Sens": "DEBIT",
            "Ref": _s(r.get("No_Transaction")),
            "Tiers": _s(r.get("Prestataire")),
            "Libelle": _s(r.get("Description")),
            "Debit": abs(_n(r.get("Montant"))),
            "Credit": 0.0,
            "Banque": _s(r.get("Banque")) or normaliser_banque(banque),
            "Piece": _s(r.get("Num_Cheque")),
        })
    for _, r in credits.iterrows():
        rows.append({
            "Date": r.get("Date"),
            "Sens": "CREDIT",
            "Ref": _s(r.get("Ref")),
            "Tiers": _s(r.get("Tiers")),
            "Libelle": f"{_s(r.get('Type'))} — {_s(r.get('Motif'))}".strip(" —"),
            "Debit": 0.0,
            "Credit": abs(_n(r.get("Montant"))),
            "Banque": _s(r.get("Banque")),
            "Piece": _s(r.get("Piece")),
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df = df.sort_values("Date").reset_index(drop=True)
    soldes = charger_soldes_depart()
    solde = soldes.get(normaliser_banque(banque), 0.0)
    running = []
    for _, r in df.iterrows():
        solde = solde + _n(r["Credit"]) - _n(r["Debit"])
        running.append(solde)
    df["Solde"] = running
    return df


def resume_solde(banque, date_debut, date_fin):
    soldes = charger_soldes_depart()
    ouverture = soldes.get(normaliser_banque(banque), 0.0)
    debits = charger_debits_suivi(banque, date_debut, date_fin)
    credits = charger_credits_tous(banque, date_debut, date_fin)
    total_d = float(debits["Montant"].abs().sum()) if not debits.empty else 0.0
    total_c = float(credits["Montant"].sum()) if not credits.empty else 0.0
    cloture = ouverture + total_c - total_d
    return {
        "ouverture": ouverture,
        "total_credits": total_c,
        "total_debits": total_d,
        "cloture": cloture,
        "nb_credits": len(credits),
        "nb_debits": len(debits),
    }


def generer_pdf_a3(banque, date_debut, date_fin, mouvements, resume):
    """Rapport A3 paysage débit/crédit/solde."""
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A3),
        leftMargin=12 * mm, rightMargin=12 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle("t", parent=styles["Heading1"], alignment=TA_CENTER, fontSize=16, spaceAfter=4)
    sub = ParagraphStyle("s", parent=styles["Normal"], alignment=TA_CENTER, fontSize=10, textColor=colors.HexColor("#444444"))
    small = ParagraphStyle("sm", parent=styles["Normal"], fontSize=7, leading=9)

    story = []
    story.append(Paragraph("NAP SARL — RAPPORT BANCAIRE DÉBIT / CRÉDIT / SOLDE", title))
    story.append(Paragraph(
        f"Banque : <b>{banque}</b> &nbsp;|&nbsp; Période : <b>{date_debut.strftime('%d/%m/%Y')} → {date_fin.strftime('%d/%m/%Y')}</b> "
        f"&nbsp;|&nbsp; Édité le {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        sub,
    ))
    story.append(Spacer(1, 6 * mm))

    # KPI
    kpi = [
        ["Solde d'ouverture", "Total crédits", "Total débits", "Solde de clôture"],
        [
            fmt(resume["ouverture"]) + " F",
            fmt(resume["total_credits"]) + " F",
            fmt(resume["total_debits"]) + " F",
            fmt(resume["cloture"]) + " F",
        ],
    ]
    tk = Table(kpi, colWidths=[70 * mm, 70 * mm, 70 * mm, 70 * mm])
    tk.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("BACKGROUND", (0, 1), (0, 1), colors.HexColor("#D6EAF8")),
        ("BACKGROUND", (1, 1), (1, 1), colors.HexColor("#D5F5E3")),
        ("BACKGROUND", (2, 1), (2, 1), colors.HexColor("#FADBD8")),
        ("BACKGROUND", (3, 1), (3, 1), colors.HexColor("#FCF3CF")),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(tk)
    story.append(Spacer(1, 5 * mm))

    # Mouvements
    head = ["Date", "Sens", "Réf.", "Tiers", "Libellé", "Débit", "Crédit", "Solde", "Pièce"]
    data = [head]
    for _, r in mouvements.iterrows():
        d = r["Date"]
        ds = d.strftime("%d/%m/%Y") if hasattr(d, "strftime") else _s(d)[:10]
        data.append([
            ds,
            _s(r.get("Sens")),
            Paragraph(_s(r.get("Ref"))[:18], small),
            Paragraph(_s(r.get("Tiers"))[:28], small),
            Paragraph(_s(r.get("Libelle"))[:50], small),
            fmt(r.get("Debit")) if _n(r.get("Debit")) else "",
            fmt(r.get("Credit")) if _n(r.get("Credit")) else "",
            fmt(r.get("Solde")),
            _s(r.get("Piece"))[:12],
        ])
    if len(data) == 1:
        data.append(["—", "", "", "", "Aucun mouvement", "", "", "", ""])

    col_w = [22 * mm, 18 * mm, 28 * mm, 40 * mm, 70 * mm, 28 * mm, 28 * mm, 28 * mm, 22 * mm]
    t = Table(data, colWidths=col_w, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("ALIGN", (5, 0), (7, -1), "RIGHT"),
        ("ALIGN", (1, 0), (1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CCCCCC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for i in range(1, len(data)):
        sens = str(data[i][1]).upper()
        if sens == "CREDIT":
            style_cmds.append(("TEXTCOLOR", (6, i), (6, i), colors.HexColor("#196F3D")))
        elif sens == "DEBIT":
            style_cmds.append(("TEXTCOLOR", (5, i), (5, i), colors.HexColor("#922B21")))
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F8F9F9")))
    t.setStyle(TableStyle(style_cmds))
    story.append(t)
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(
        "NAP SARL — Document interne de pilotage bancaire. Débits = Suivi de paiement · Crédits = base crédits bancaires.",
        ParagraphStyle("f", parent=styles["Normal"], fontSize=7, textColor=colors.grey, alignment=TA_CENTER),
    ))
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


def to_excel_rapport(mouvements, resume, banque, date_debut, date_fin):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        m = mouvements.copy()
        if not m.empty and "Date" in m.columns:
            m["Date"] = pd.to_datetime(m["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
        m.to_excel(w, sheet_name="Mouvements", index=False)
        pd.DataFrame([
            {"Indicateur": "Banque", "Valeur": banque},
            {"Indicateur": "Période début", "Valeur": date_debut.strftime("%d/%m/%Y")},
            {"Indicateur": "Période fin", "Valeur": date_fin.strftime("%d/%m/%Y")},
            {"Indicateur": "Solde ouverture", "Valeur": resume["ouverture"]},
            {"Indicateur": "Total crédits", "Valeur": resume["total_credits"]},
            {"Indicateur": "Total débits", "Valeur": resume["total_debits"]},
            {"Indicateur": "Solde clôture", "Valeur": resume["cloture"]},
        ]).to_excel(w, sheet_name="Resume", index=False)
    buf.seek(0)
    return buf.getvalue()
