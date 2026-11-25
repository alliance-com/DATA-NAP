# -*- coding: utf-8 -*-
"""Génération PDF A3 paysage — listes NAP SARL (impression professionnelle)."""
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
from io import BytesIO
import os
import re

PAGE = landscape(A3)  # 420 x 297 mm
MARGIN = 12 * mm

BLEU = colors.Color(27 / 255, 79 / 255, 114 / 255)
BLEU_C = colors.Color(135 / 255, 206 / 255, 235 / 255)
BLEU_L = colors.Color(234 / 255, 246 / 255, 253 / 255)
GRIS = colors.Color(0.25, 0.25, 0.25)
VERT = colors.Color(34 / 255, 139 / 255, 34 / 255)


def _styles():
    base = getSampleStyleSheet()
    styles = {
        "title": ParagraphStyle(
            "NAPTitle", parent=base["Heading1"],
            fontSize=16, textColor=BLEU, spaceAfter=2 * mm, alignment=1, fontName="Helvetica-Bold",
        ),
        "sub": ParagraphStyle(
            "NAPSub", parent=base["Normal"],
            fontSize=8, textColor=GRIS, spaceAfter=1 * mm, alignment=1, fontName="Helvetica",
        ),
        "cell": ParagraphStyle(
            "NAPCell", parent=base["Normal"],
            fontSize=6.5, textColor=GRIS, leading=8, fontName="Helvetica",
        ),
        "cell_b": ParagraphStyle(
            "NAPCellB", parent=base["Normal"],
            fontSize=6.5, textColor=GRIS, leading=8, fontName="Helvetica-Bold",
        ),
        "head": ParagraphStyle(
            "NAPHead", parent=base["Normal"],
            fontSize=7, textColor=BLEU, leading=9, fontName="Helvetica-Bold", alignment=1,
        ),
        "foot": ParagraphStyle(
            "NAPFoot", parent=base["Normal"],
            fontSize=7, textColor=GRIS, alignment=1, fontName="Helvetica",
        ),
        "kpi": ParagraphStyle(
            "NAPKpi", parent=base["Normal"],
            fontSize=8, textColor=BLEU, fontName="Helvetica-Bold", alignment=1,
        ),
    }
    return styles


def _p(text, style):
    t = "" if text is None else str(text)
    t = t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return Paragraph(t, style)


def _footer(canvas, doc):
    canvas.saveState()
    w, h = PAGE
    canvas.setStrokeColor(colors.Color(0.36, 0.68, 0.89))
    canvas.setLineWidth(1.2)
    canvas.line(MARGIN, 10 * mm, w - MARGIN, 10 * mm)
    canvas.setFont("Helvetica", 6.5)
    canvas.setFillColor(GRIS)
    txt = (
        "STE NAP SARL  •  PLOT NO 14 GLODJIGBE ECONOMIC ZONE — BENIN  •  "
        "RCCM RB/COT/21 B 30710  •  IFU 3202113334014  •  "
        "TEL +229 01 94 94 04 14 / +229 01 97 97 45 06"
    )
    canvas.drawCentredString(w / 2, 6 * mm, txt)
    canvas.drawRightString(w - MARGIN, 6 * mm, f"Page {doc.page}")
    canvas.restoreState()


def _header_block(styles, titre, sous_titre, kpis):
    elems = []
    elems.append(_p("STE NAP SARL", styles["title"]))
    elems.append(_p(titre, styles["title"]))
    if sous_titre:
        elems.append(_p(sous_titre, styles["sub"]))
    if kpis:
        elems.append(_p("  |  ".join(kpis), styles["kpi"]))
    elems.append(Spacer(1, 3 * mm))
    return elems


def _table(headers, rows, styles, col_widths=None):
    data = [[_p(h, styles["head"]) for h in headers]]
    for r in rows:
        data.append([_p(c, styles["cell"]) for c in r])
    t = Table(data, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), BLEU_C),
        ("TEXTCOLOR", (0, 0), (-1, 0), BLEU),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.5),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.Color(0.7, 0.8, 0.9)),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, BLEU_L]),
    ]
    t.setStyle(TableStyle(style_cmds))
    return t


def _usable_width():
    return PAGE[0] - 2 * MARGIN


def generer_pdf_liste(titre, sous_titre, headers, rows, kpis=None, fichier=None, lang="FR"):
    """Génère un PDF A3 paysage. Retourne bytes si fichier is None, sinon chemin. lang=FR|EN"""
    styles = _styles()
    buf = BytesIO() if fichier is None else None
    target = buf if buf is not None else fichier
    doc = SimpleDocTemplate(
        target,
        pagesize=PAGE,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=12 * mm,
        bottomMargin=16 * mm,
        title=titre,
        author="STE NAP SARL",
    )
    story = _header_block(styles, titre, sous_titre, kpis or [])
    n = max(len(headers), 1)
    usable = _usable_width()
    col_w = [usable / n] * n
    story.append(_table(headers, rows, styles, col_widths=col_w))
    story.append(Spacer(1, 4 * mm))
    if str(lang).upper().startswith("EN"):
        foot_txt = (
            f"Document generated on {datetime.now().strftime('%d/%m/%Y at %H:%M')} "
            f"— A3 landscape print — NAP SARL"
        )
    else:
        foot_txt = (
            f"Document généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')} "
            f"— Impression A3 paysage — NAP SARL"
        )
    story.append(_p(foot_txt, styles["foot"]))
    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    if buf is not None:
        buf.seek(0)
        return buf.read()
    return fichier


def fmt_num(v, dec=0):
    try:
        x = float(v)
        if dec == 0:
            return f"{x:,.0f}".replace(",", " ")
        return f"{x:,.{dec}f}".replace(",", " ")
    except Exception:
        return str(v or "")


def df_to_rows(df, cols, formatters=None):
    formatters = formatters or {}
    rows = []
    for _, r in df.iterrows():
        line = []
        for c in cols:
            v = r.get(c, "")
            if c in formatters:
                line.append(formatters[c](v))
            else:
                if hasattr(v, "strftime"):
                    try:
                        line.append(v.strftime("%d/%m/%Y"))
                        continue
                    except Exception:
                        pass
                line.append("" if v is None or (isinstance(v, float) and str(v) == "nan") else str(v))
        rows.append(line)
    return rows
