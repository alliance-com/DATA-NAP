# -*- coding: utf-8 -*-
"""Impression PDF Export — A4 ou A3 selon largeur du tableau."""
from io import BytesIO
from datetime import datetime

from reportlab.lib.pagesizes import A4, A3, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether,
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# Seuil : au-delà de N colonnes → A3 paysage recommandé
SEUIL_COLONNES_A3 = 10

NAP_BLUE = colors.Color(0.11, 0.31, 0.45)
NAP_SKY = colors.Color(0.53, 0.81, 0.92)
NAP_ROW = colors.Color(0.92, 0.96, 0.99)
NAP_GRID = colors.Color(0.70, 0.80, 0.90)


def _s(v):
    if v is None:
        return "—"
    t = str(v).strip()
    return t if t and t.lower() not in ("nan", "none", "nat", "natype") else "—"


def choisir_format(n_cols, force=None):
    """force: 'A4' | 'A3' | None (auto)."""
    if force in ("A4", "A3"):
        return force
    return "A3" if n_cols >= SEUIL_COLONNES_A3 else "A4"


def page_size(fmt):
    # Toujours paysage pour tableaux opérationnels
    return landscape(A3 if fmt == "A3" else A4)


def generer_pdf_tableau(
    titre,
    sous_titre,
    headers,
    rows,
    format_page=None,
    colonnes_cm=None,
):
    """
    headers: list[str]
    rows: list[list]
    format_page: None | 'A4' | 'A3'
    """
    n_cols = len(headers)
    fmt = choisir_format(n_cols, format_page)
    pagesize = page_size(fmt)
    page_w, page_h = pagesize

    buf = BytesIO()
    left = 10 * mm
    right = 10 * mm
    top = 10 * mm
    bottom = 12 * mm
    usable = page_w - left - right

    doc = SimpleDocTemplate(
        buf,
        pagesize=pagesize,
        leftMargin=left,
        rightMargin=right,
        topMargin=top,
        bottomMargin=bottom,
        title=titre,
        author="STE NAP SARL",
    )

    styles = getSampleStyleSheet()
    st_title = ParagraphStyle(
        "t", parent=styles["Heading1"], fontSize=12 if fmt == "A4" else 14,
        textColor=NAP_BLUE, alignment=TA_CENTER, spaceAfter=2, leading=14,
    )
    st_sub = ParagraphStyle(
        "s", parent=styles["Normal"], fontSize=8 if fmt == "A4" else 9,
        textColor=colors.Color(0.35, 0.35, 0.35), alignment=TA_CENTER, spaceAfter=4,
    )
    st_cell = ParagraphStyle(
        "c", parent=styles["Normal"], fontSize=6.5 if fmt == "A4" else 7.5,
        leading=8 if fmt == "A4" else 9,
    )
    st_head = ParagraphStyle(
        "h", parent=styles["Normal"], fontSize=6.5 if fmt == "A4" else 7.5,
        textColor=NAP_BLUE, fontName="Helvetica-Bold", alignment=TA_CENTER, leading=8,
    )
    st_foot = ParagraphStyle(
        "f", parent=styles["Normal"], fontSize=7, textColor=colors.grey, alignment=TA_CENTER,
    )

    story = []
    story.append(Paragraph("STE NAP SARL — Export", st_title))
    story.append(Paragraph(titre, st_title))
    story.append(Paragraph(
        f"{sous_titre}  ·  Format {fmt} paysage  ·  {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        st_sub,
    ))
    story.append(Spacer(1, 2 * mm))

    data = [[Paragraph(_s(h), st_head) for h in headers]]
    for row in rows:
        data.append([Paragraph(_s(c), st_cell) for c in row])

    if len(data) == 1:
        data.append([Paragraph("—", st_cell)] * n_cols)

    # Largeurs de colonnes
    if colonnes_cm and len(colonnes_cm) == n_cols:
        col_w = [w * 10 * mm / 10 for w in colonnes_cm]  # cm → roughly
        # normalize to usable
        total = sum(colonnes_cm)
        col_w = [usable * (w / total) for w in colonnes_cm]
    else:
        col_w = [usable / n_cols] * n_cols

    tbl = Table(data, colWidths=col_w, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAP_SKY),
        ("TEXTCOLOR", (0, 0), (-1, 0), NAP_BLUE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.3, NAP_GRID),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, NAP_ROW]),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(
        f"STE NAP SARL  ·  Glo-Djigbé  ·  Document généré le {datetime.now().strftime('%d/%m/%Y %H:%M')}  ·  {fmt} paysage",
        st_foot,
    ))

    doc.build(story)
    buf.seek(0)
    return buf.read(), fmt


# ---- Jeux de colonnes par type de rapport ----

RAPPORTS = {
    "bookings": {
        "titre": "LISTE DES BOOKINGS / BL",
        "headers": [
            "Booking", "BL", "Commodity", "Buyer", "Forwarder", "Liner",
            "Ctr B.", "Ctr L.", "Bal.", "Net KG", "Tank", "Stuffing",
            "Vessel", "ETD", "ETA", "Shipping",
        ],
        "keys": [
            "Booking_No", "BL_No", "Commodity", "Buyer", "Forwarder", "Liner",
            "Ctr_Booked", "Ctr_Loaded", "Balance", "Net_Qty_KG", "Storage_Tank",
            "Stuffing_Status", "Vessel", "ETD", "ETA", "Shipping_Status",
        ],
        "format_defaut": "A3",
    },
    "stuffing_pending": {
        "titre": "BOOKINGS EN ATTENTE D'EMPOTAGE",
        "headers": [
            "Booking", "BL", "Commodity", "Buyer", "Forwarder",
            "Ctr B.", "Ctr L.", "Balance", "Tank", "Stuffing", "Vessel", "ETD",
        ],
        "keys": [
            "Booking_No", "BL_No", "Commodity", "Buyer", "Forwarder",
            "Ctr_Booked", "Ctr_Loaded", "Balance", "Storage_Tank",
            "Stuffing_Status", "Vessel", "ETD",
        ],
        "format_defaut": "A4",
    },
    "containers": {
        "titre": "LISTE DES CONTENEURS",
        "headers": [
            "Conteneur", "Booking", "BL", "Type", "Seal", "Poids KG",
            "Status", "Tank", "Date stuffing", "Notes",
        ],
        "keys": [
            "Container_No", "Booking_No", "BL_No", "Type_Ctr", "Seal", "Weight_KG",
            "Status_Load", "Tank", "Date_Stuffing", "Notes",
        ],
        "format_defaut": "A4",
    },
    "contracts": {
        "titre": "LISTE DES CONTRATS",
        "headers": [
            "Contrat", "Date", "Buyer", "Commodity", "Qty MT", "PU", "Devise",
            "POD", "Country", "Actif",
        ],
        "keys": [
            "Contract_No", "Date_Contract", "Buyer", "Commodity", "Qty_Contract_MT",
            "Price_Unit", "Currency", "POD", "Country", "Actif",
        ],
        "format_defaut": "A4",
    },
    "recap": {
        "titre": "RECAP EXPORT — STATUTS",
        "headers": [
            "Booking", "BL", "Commodity", "Buyer", "Forwarder", "POD",
            "Ctr B.", "Ctr L.", "Stuffing", "Vessel", "ETD", "ETA", "Shipping",
        ],
        "keys": [
            "Booking_No", "BL_No", "Commodity", "Buyer", "Forwarder", "POD",
            "Ctr_Booked", "Ctr_Loaded", "Stuffing_Status", "Vessel", "ETD", "ETA",
            "Shipping_Status",
        ],
        "format_defaut": "A3",
    },
}


def _fmt_date(v):
    try:
        import pandas as pd
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "—"
        ts = pd.to_datetime(v, errors="coerce")
        if pd.isna(ts):
            return _s(v)
        return ts.strftime("%d/%m/%Y")
    except Exception:
        return _s(v)


def dataframe_vers_rows(df, keys):
    import pandas as pd
    rows = []
    if df is None or df.empty:
        return rows
    for _, r in df.iterrows():
        line = []
        for k in keys:
            v = r[k] if k in r.index else ""
            if k in ("ETD", "ETA", "Date_Creation", "Date_Contract", "Date_Stuffing"):
                line.append(_fmt_date(v))
            else:
                line.append(_s(v))
        rows.append(line)
    return rows


def generer_rapport(kind, df, format_page=None, filtre_label=""):
    meta = RAPPORTS.get(kind)
    if not meta:
        raise ValueError(f"Type de rapport inconnu : {kind}")
    headers = meta["headers"]
    keys = meta["keys"]
    rows = dataframe_vers_rows(df, keys)
    fmt = format_page or meta.get("format_defaut")
    sous = filtre_label or f"{len(rows)} ligne(s)"
    return generer_pdf_tableau(meta["titre"], sous, headers, rows, format_page=fmt)
