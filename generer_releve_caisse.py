# -*- coding: utf-8 -*-
"""Fiche relevé de caisse boutique — Word A4 + PDF A4."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
from io import BytesIO
import os

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.enums import TA_CENTER

DOSSIER = "releves_caisse"
os.makedirs(DOSSIER, exist_ok=True)

BLEU = RGBColor(27, 79, 114)
BLEU_M = RGBColor(46, 134, 193)
GRIS = RGBColor(55, 55, 55)
GRIS_C = RGBColor(110, 110, 110)
VERT = RGBColor(34, 139, 34)
ROUGE = RGBColor(192, 57, 43)
CIEL = "5DADE2"


def set_run(run, size=9, bold=False, color=None, name="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name
    if color:
        run.font.color.rgb = color


def shade(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def borders(table, color=CIEL, size="8"):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), size)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        b.append(el)
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    tblPr.append(b)


def fmt(n):
    try:
        return f"{float(n):,.0f}".replace(",", " ")
    except Exception:
        return str(n)


def generer_releve_caisse_word(data):
    doc = Document()
    for s in doc.sections:
        s.page_width = Cm(21)
        s.page_height = Cm(29.7)
        s.top_margin = Cm(1.2)
        s.bottom_margin = Cm(1.5)
        s.left_margin = Cm(1.4)
        s.right_margin = Cm(1.4)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run("STE NAP SARL"), 14, True, BLEU)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p2.add_run("Caisse Boutique — Relevé de caisse"), 12, True, BLEU_M)

    ligne = doc.add_paragraph()
    pPr = ligne._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "18")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), CIEL)
    pBdr.append(bot)
    pPr.append(pBdr)

    d1 = data.get("Periode_Debut", "")
    d2 = data.get("Periode_Fin", "")
    mode = data.get("Mode_Filtre") or "Tous"

    info = doc.add_paragraph()
    set_run(info.add_run("Période : " + d1 + " → " + d2 + "    |    Mode : " + str(mode)), 9, False, GRIS)

    t = doc.add_table(rows=2, cols=4)
    borders(t)
    headers = ["Solde départ", "Total crédits", "Total débits", "Solde final"]
    vals = [
        fmt(data.get("Solde_Depart")),
        fmt(data.get("Total_Credit")),
        fmt(data.get("Total_Debit")),
        fmt(data.get("Solde_Final")),
    ]
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        shade(cell, "87CEEB")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(cell.paragraphs[0].add_run(h), 8, True, BLEU)
    for i, v in enumerate(vals):
        cell = t.rows[1].cells[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        col = ROUGE if i == 2 else VERT
        set_run(cell.paragraphs[0].add_run(v + " FCFA"), 10, True, col)

    sec = doc.add_paragraph()
    set_run(sec.add_run("Détail des mouvements"), 10, True, BLEU_M)

    lignes = data.get("Lignes") or []
    n = max(len(lignes), 1)
    det = doc.add_table(rows=1 + n, cols=7)
    borders(det, size="6")
    heads = ["Date", "Libellé", "Mode", "Crédit", "Débit", "Solde", "Réf."]
    for i, h in enumerate(heads):
        cell = det.rows[0].cells[i]
        shade(cell, "87CEEB")
        set_run(cell.paragraphs[0].add_run(h), 7, True, BLEU)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    if not lignes:
        for i in range(7):
            set_run(det.rows[1].cells[i].paragraphs[0].add_run("—"), 7, False, GRIS_C)
    else:
        for ri, L in enumerate(lignes, start=1):
            c_val = float(L.get("Credit") or 0)
            d_val = float(L.get("Debit") or 0)
            vals_row = [
                L.get("Date", "—"),
                str(L.get("Libelle") or "—")[:40],
                L.get("Mode") or "—",
                fmt(c_val) if c_val else "—",
                fmt(d_val) if d_val else "—",
                fmt(L.get("Solde") or 0),
                str(L.get("Reference") or "—")[:16],
            ]
            fill = "EAF6FD" if ri % 2 == 0 else "FFFFFF"
            for i, v in enumerate(vals_row):
                cell = det.rows[ri].cells[i]
                shade(cell, fill)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i != 1 else WD_ALIGN_PARAGRAPH.LEFT
                col = GRIS
                if i == 3 and v != "—":
                    col = VERT
                if i == 4 and v != "—":
                    col = ROUGE
                set_run(p.add_run(str(v)), 7, i in (3, 4, 5), col)

    note = doc.add_paragraph()
    set_run(
        note.add_run(
            "NB : Relevé de caisse boutique. Modes : Espèce et Mobile Money. "
            "Tout écart doit être signalé immédiatement."
        ),
        7, True, ROUGE,
    )
    foot = doc.add_paragraph()
    foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(foot.add_run("STE NAP SARL  •  Glo-Djigbé  •  TEL +229 01 94 94 04 14"), 7, False, GRIS_C)

    nom = "Releve_Caisse_" + datetime.now().strftime("%Y%m%d_%H%M") + ".docx"
    chemin = os.path.join(DOSSIER, nom)
    doc.save(chemin)
    return chemin


def generer_releve_caisse_pdf(data):
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=12 * mm, bottomMargin=14 * mm,
        title="Relevé de caisse", author="STE NAP SARL",
    )
    styles = getSampleStyleSheet()
    st_title = ParagraphStyle(
        "t", parent=styles["Heading1"], fontSize=13,
        textColor=colors.Color(0.11, 0.31, 0.45), alignment=TA_CENTER, spaceAfter=4,
    )
    st_cell = ParagraphStyle("c", parent=styles["Normal"], fontSize=7, leading=9)
    st_head = ParagraphStyle(
        "h", parent=styles["Normal"], fontSize=7,
        textColor=colors.Color(0.11, 0.31, 0.45), fontName="Helvetica-Bold", alignment=TA_CENTER,
    )
    st_foot = ParagraphStyle("f", parent=styles["Normal"], fontSize=7, textColor=colors.grey, alignment=TA_CENTER)

    story = []
    story.append(Paragraph("STE NAP SARL — Caisse Boutique", st_title))
    story.append(Paragraph("RELEVÉ DE CAISSE", st_title))
    story.append(Paragraph(
        "Période : " + str(data.get("Periode_Debut")) + " → " + str(data.get("Periode_Fin"))
        + " &nbsp;|&nbsp; Mode : " + str(data.get("Mode_Filtre") or "Tous"),
        ParagraphStyle("s", parent=styles["Normal"], fontSize=9, alignment=TA_CENTER, textColor=colors.Color(0.3, 0.3, 0.3)),
    ))
    story.append(Spacer(1, 3 * mm))

    synth = [[
        Paragraph("<b>Solde départ</b>", st_head),
        Paragraph("<b>Total crédits</b>", st_head),
        Paragraph("<b>Total débits</b>", st_head),
        Paragraph("<b>Solde final</b>", st_head),
    ], [
        Paragraph(fmt(data.get("Solde_Depart")) + " FCFA", st_cell),
        Paragraph(fmt(data.get("Total_Credit")) + " FCFA", st_cell),
        Paragraph(fmt(data.get("Total_Debit")) + " FCFA", st_cell),
        Paragraph(fmt(data.get("Solde_Final")) + " FCFA", st_cell),
    ]]
    usable = A4[0] - 28 * mm
    ts = Table(synth, colWidths=[usable / 4] * 4)
    ts.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.Color(0.53, 0.81, 0.92)),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.Color(0.7, 0.8, 0.9)),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(ts)
    story.append(Spacer(1, 4 * mm))

    heads = ["Date", "Libellé", "Mode", "Crédit", "Débit", "Solde", "Réf."]
    data_tbl = [[Paragraph(h, st_head) for h in heads]]
    for L in data.get("Lignes") or []:
        c_val = float(L.get("Credit") or 0)
        d_val = float(L.get("Debit") or 0)
        data_tbl.append([
            Paragraph(str(L.get("Date") or "—"), st_cell),
            Paragraph(str(L.get("Libelle") or "—")[:40], st_cell),
            Paragraph(str(L.get("Mode") or "—"), st_cell),
            Paragraph(fmt(c_val) if c_val else "—", st_cell),
            Paragraph(fmt(d_val) if d_val else "—", st_cell),
            Paragraph(fmt(L.get("Solde")), st_cell),
            Paragraph(str(L.get("Reference") or "—")[:14], st_cell),
        ])
    if len(data_tbl) == 1:
        data_tbl.append([Paragraph("—", st_cell)] * 7)

    widths = [usable * x for x in (0.11, 0.28, 0.12, 0.12, 0.12, 0.13, 0.12)]
    t = Table(data_tbl, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.Color(0.53, 0.81, 0.92)),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.Color(0.7, 0.8, 0.9)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.Color(0.92, 0.96, 0.99)]),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(t)
    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph(
        "NB : Relevé de caisse boutique (Espèce / Mobile Money). Signalement immédiat en cas d'écart.",
        ParagraphStyle("n", parent=styles["Normal"], fontSize=7, textColor=colors.Color(0.75, 0.22, 0.17), fontName="Helvetica-Bold"),
    ))
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph("STE NAP SARL  •  Glo-Djigbé  •  TEL +229 01 94 94 04 14", st_foot))
    doc.build(story)
    buf.seek(0)
    return buf.read()
