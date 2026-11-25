# -*- coding: utf-8 -*-
"""
Commercial Invoice NAP — calé sur le modèle projet (corrections layout / gras).
"""
from __future__ import annotations

import os
import re
from datetime import datetime
from io import BytesIO

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    import qrcode
except ImportError:
    qrcode = None

_BASE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_BASE)
DOSSIER = os.path.join(_ROOT, "commercial_invoices_generes")
os.makedirs(DOSSIER, exist_ok=True)

BLEU = RGBColor(27, 79, 114)
BLEU_M = RGBColor(46, 134, 193)
CIEL = "5DADE2"
CIEL_H = "87CEEB"
NOIR = RGBColor(20, 20, 20)
GRIS = RGBColor(70, 70, 70)
GRIS_C = RGBColor(120, 120, 120)
FONT = "Times New Roman"


def _s(v):
    if v is None:
        return ""
    return str(v).strip()


def fmt_amt(n, dec=2):
    try:
        return f"{float(n):,.{dec}f}"
    except Exception:
        return str(n)


def fmt_qty(n):
    try:
        v = float(n)
        return f"{v:,.3f}".rstrip("0").rstrip(".")
    except Exception:
        return str(n)


def set_run(run, size=10, bold=False, underline=False, color=None, name=FONT):
    run.font.size = Pt(size)
    run.font.bold = bool(bold)
    run.font.underline = bool(underline)
    run.font.name = name
    try:
        run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    except Exception:
        pass
    if color is not None:
        run.font.color.rgb = color


def tight(p, before=0, after=0, line=1.0):
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = line


def shade(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_margins(cell, top=60, bottom=60, left=60, right=60):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for edge, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    # hauteur min cellule description
    tcPr.append(tcMar)


def set_row_height(row, height_twips=1400):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(height_twips))
    trHeight.set(qn("w:hRule"), "atLeast")
    trPr.append(trHeight)


def borders(table, color="000000", size="6"):
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


def qr_bytes(texte):
    if not qrcode:
        return None
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=6,
        border=1,
    )
    qr.add_data(texte)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def _date_en(d):
    if d is None:
        d = datetime.now()
    if not hasattr(d, "strftime"):
        try:
            import pandas as pd
            d = pd.to_datetime(d, errors="coerce", dayfirst=True)
            if pd.isna(d):
                d = datetime.now()
        except Exception:
            d = datetime.now()
    day = int(d.day)
    if 10 <= day % 100 <= 20:
        suf = "TH"
    else:
        suf = {1: "ST", 2: "ND", 3: "RD"}.get(day % 10, "TH")
    return f"{day}{suf} {d.strftime('%B')}  {d.year}"


def pied_page(doc):
    footer = doc.sections[0].footer
    footer.is_linked_to_previous = False
    p0 = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(p0, 0, 2)
    pPr = p0._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "8")
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), "AAAAAA")
    pBdr.append(top)
    pPr.append(pBdr)
    set_run(
        p0.add_run(
            "STE NAP SARL  ·  PLOT NO 14 GLODJIGBE ECONOMIC ZONE  ·  "
            "RCCM RB/COT/21 B 30710  ·  IFU 3202113334014  ·  TEL +229 94 94 04 14"
        ),
        7, False, False, GRIS_C,
    )


# Libellés description : le LIBELLÉ peut être distingué, le CONTENU après n'est pas gras
_LABELS_DESC = (
    "CONTRACT NO:",
    "CONTRACT NO :",
    "Prices are",
    "Bill of Lading:",
    "Bill of Lading :",
    "Shipping Line:",
    "Shipping Line :",
    "ORIGIN OF GOODS",
    "FINAL DESTINATION",
)


def _split_desc_label_value(line: str):
    """Retourne (label avec :, value). Libellé gras, valeur non grasse."""
    t = line.strip()
    # ORIGIN / FINAL DESTINATION (souvent sans structure identique)
    for key in ("ORIGIN OF GOODS", "FINAL DESTINATION"):
        if t.upper().startswith(key):
            # tout avant un séparateur long reste label si présent
            if "–" in t or " - " in t or ":" in t:
                for sep in ("–", " - ", ":"):
                    if sep in t:
                        i = t.index(sep)
                        return t[: i + len(sep)].strip(), t[i + len(sep) :].strip()
            return t, ""
    if ":" in t[:48]:
        idx = t.index(":")
        label = t[: idx + 1].strip()
        value = t[idx + 1 :].strip()
        # ne traiter que les libellés connus
        up = label.upper()
        if any(
            x in up
            for x in (
                "CONTRACT",
                "BILL OF LADING",
                "SHIPPING LINE",
                "PRICES ARE",
                "ORIGIN",
                "FINAL DESTINATION",
            )
        ):
            return label, value
    return None, t


def generer_commercial_invoice(data, lang="EN"):
    doc = Document()
    for s in doc.sections:
        s.page_width = Cm(21.0)
        s.page_height = Cm(29.7)
        s.top_margin = Cm(1.2)
        s.bottom_margin = Cm(1.8)
        s.left_margin = Cm(1.6)
        s.right_margin = Cm(1.6)
        s.header_distance = Cm(0.4)
        s.footer_distance = Cm(0.5)

    pied_page(doc)

    num = _s(data.get("Invoice_No"))
    buyer = _s(data.get("Buyer"))
    currency = _s(data.get("Currency") or "EUR")
    net = data.get("Net_Amount") or 0
    qty = data.get("Qty_MT") or 0
    price = data.get("Price_Per_MT") or 0
    gross = data.get("Gross_Amount") or 0
    advance = float(data.get("Advance") or 0)
    pi_pct = float(data.get("Less_PI_Pct") or 0)
    pi_amt = float(data.get("Less_PI_Amount") or 0)

    # ===== En-tête + QR =====
    t_top = doc.add_table(rows=1, cols=2)
    borders(t_top, color="CCCCCC", size="4")
    c0, c1 = t_top.rows[0].cells
    shade(c0, CIEL_H)
    set_cell_margins(c0, 50, 50, 80, 40)
    set_cell_margins(c1, 30, 30, 20, 40)

    p = c0.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(p, 0, 0)
    set_run(p.add_run("STE NAP SARL"), 13, True, False, BLEU)

    for line in [
        "PLOT NO 14 GLODJIGBE ECONOMIC ZONE, REPUBLIC OF BENIN",
        "RCCM RB / COT / 21 B 30710, IFU 3202113334014",
        "TEL: +229 94940414, 97974506",
    ]:
        pp = c0.add_paragraph()
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tight(pp, 0, 0)
        set_run(pp.add_run(line), 8, True, False, BLEU)

    c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    texte_qr = f"STE NAP SARL|CI|{num}|{buyer}|{fmt_amt(net)} {currency}"
    buf = qr_bytes(texte_qr)
    if buf:
        try:
            c1.paragraphs[0].add_run().add_picture(buf, width=Cm(1.85))
        except Exception:
            pass

    # Ligne simple sous en-tête (pas bleu ciel)
    sep = doc.add_paragraph()
    tight(sep, 4, 6)
    pPr = sep._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "12")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), "000000")
    pBdr.append(bot)
    pPr.append(pBdr)

    # Titre
    titre = doc.add_paragraph()
    titre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(titre, 2, 10)
    set_run(titre.add_run("COMMERCIAL INVOICE"), 14, True, False, NOIR)

    # INVOICE NUMBER
    p = doc.add_paragraph()
    tight(p, 2, 4)
    set_run(p.add_run("INVOICE NUMBER  : "), 11, True, False, NOIR)
    set_run(p.add_run("                                                 "), 11, False, False, NOIR)
    set_run(p.add_run(num), 11, True, False, NOIR)

    # DATE — libellé non gras, date gras
    dtxt = _date_en(data.get("Date_Invoice"))
    p = doc.add_paragraph()
    tight(p, 2, 8)
    set_run(p.add_run("DATE:     : "), 11, False, False, NOIR)
    set_run(p.add_run("                                                                       "), 11, False, False, NOIR)
    set_run(p.add_run(dtxt), 11, True, False, NOIR)

    # BUYER — libellé gras + nom sur la même ligne (non gras)
    # Adresses EN DESSOUS, alignées sous le bloc (pas collées au libellé BUYER)
    p = doc.add_paragraph()
    tight(p, 4, 2)
    set_run(p.add_run("BUYER    :  "), 11, True, False, NOIR)
    set_run(p.add_run(buyer), 11, False, False, NOIR)

    addr = _s(data.get("Buyer_Address"))
    if addr:
        for ln in addr.split("\n"):
            if ln.strip():
                pa = doc.add_paragraph()
                tight(pa, 0, 0)
                # retrait pour placer l'adresse sous le nom (pas sous BUYER)
                pa.paragraph_format.left_indent = Cm(2.2)
                set_run(pa.add_run(ln.strip()), 10, False, False, NOIR)

    sp = doc.add_paragraph()
    tight(sp, 10, 6)

    # ===== Tableau (hauteur augmentée) =====
    price_txt = f"{currency} {fmt_amt(price, 0) if float(price) == int(float(price)) else fmt_amt(price)}"
    price_lines = [price_txt]
    amt_lines = [fmt_amt(gross)]
    if advance > 0:
        price_lines.append("Less Advance")
        amt_lines.append(fmt_amt(advance))
    if pi_pct > 0:
        price_lines.append(f"Less Pi {fmt_amt(pi_pct, 0)} %")
        amt_lines.append(fmt_amt(pi_amt))
    if advance > 0 or pi_pct > 0:
        price_lines.extend(["Final Amount", "Due"])
        amt_lines.append(fmt_amt(net))

    col_amt = (
        f"NET AMOUNT ({currency})"
        if advance == 0 and pi_pct == 0
        else f"TOTAL INV VALUE ({currency})"
    )

    tbl = doc.add_table(rows=2, cols=4)
    borders(tbl, color="000000", size="6")
    headers = ["DESCRIPTION", "QUANTITY (MT)", "PRICE PER MT", col_amt]
    for i, h in enumerate(headers):
        cell = tbl.rows[0].cells[i]
        shade(cell, "D9E2F3")
        set_cell_margins(cell, 40, 40, 40, 40)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p.add_run(h), 9, True, False, NOIR)

    # Ligne données : hauteur min plus grande
    set_row_height(tbl.rows[1], 2200)

    desc = _s(data.get("Description"))
    c0 = tbl.rows[1].cells[0]
    set_cell_margins(c0, 80, 120, 60, 50)
    parts = desc.split("\n") if desc else [""]

    # 1re ligne produit — non gras
    p = c0.paragraphs[0]
    set_run(p.add_run(parts[0]), 8, False, False, NOIR)

    for part in parts[1:]:
        pp = c0.add_paragraph()
        tight(pp, 2, 0)
        label, value = _split_desc_label_value(part)
        if label is not None:
            # Libellé avant ":" en gras ; contenu après non gras
            set_run(pp.add_run(label + " "), 8, True, False, NOIR)
            set_run(pp.add_run(value), 8, False, False, NOIR)
        else:
            set_run(pp.add_run(part), 8, False, False, NOIR)

    c1 = tbl.rows[1].cells[1]
    set_cell_margins(c1, 80, 80, 30, 30)
    c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(c1.paragraphs[0].add_run(f"{fmt_qty(qty)} MT"), 10, True, False, NOIR)

    c2 = tbl.rows[1].cells[2]
    set_cell_margins(c2, 80, 80, 30, 30)
    p = c2.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run(price_lines[0]), 9, True, False, NOIR)
    for ln in price_lines[1:]:
        pp = c2.add_paragraph()
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tight(pp, 0, 0)
        set_run(pp.add_run(ln), 8, False, False, NOIR)

    c3 = tbl.rows[1].cells[3]
    set_cell_margins(c3, 80, 80, 30, 40)
    p = c3.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_run(p.add_run(amt_lines[0]), 9, True, False, NOIR)
    for ln in amt_lines[1:]:
        pp = c3.add_paragraph()
        pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        tight(pp, 0, 0)
        set_run(pp.add_run(ln), 9, True, False, NOIR)

    # IN WORDS: — libellé gras ; phrase montants NON gras
    words = _s(data.get("Amount_Words"))
    if not words:
        try:
            from export.ci_amount_words import amount_in_words
            words = amount_in_words(net if (advance > 0 or pi_pct > 0) else gross, currency)
        except Exception:
            try:
                from ci_amount_words import amount_in_words
                words = amount_in_words(net if (advance > 0 or pi_pct > 0) else gross, currency)
            except Exception:
                words = ""
    p = doc.add_paragraph()
    tight(p, 14, 6)
    set_run(p.add_run("IN WORDS: "), 10, True, False, NOIR)
    set_run(p.add_run(words), 10, False, False, NOIR)

    # FOR STE NAP SARL + grand espace signature
    p = doc.add_paragraph()
    tight(p, 18, 2)
    set_run(p.add_run("FOR STE NAP SARL"), 10, True, False, NOIR)

    # espace pour signature manuscrite
    for _ in range(3):
        blank = doc.add_paragraph()
        tight(blank, 8, 8)
        set_run(blank.add_run(" "), 12, False, False, NOIR)

    p = doc.add_paragraph()
    tight(p, 4, 2)
    set_run(p.add_run("AUTHORIZED SIGNATORY"), 10, True, False, NOIR)

    # BANK DETAILS — gras
    bank = _s(data.get("Bank_Details"))
    p = doc.add_paragraph()
    tight(p, 14, 6)
    set_run(p.add_run("BANK DETAILS: "), 9, True, False, NOIR)
    set_run(p.add_run(bank), 9, True, False, NOIR)

    # INSPECTION / INSURANCE — libellés gras + soulignés ; texte suite normal
    clauses = [
        (
            "INSPECTION:",
            " Inspection of the Goods shall be performed by the buyer according to the export regulations of the country",
        ),
        (
            "INSURANCE:",
            " In case of CFR / C & F terms, the buyer will take complete responsibility of insuring the goods",
        ),
    ]
    for lab, txt in clauses:
        p = doc.add_paragraph()
        tight(p, 6, 3)
        set_run(p.add_run(lab), 8, True, True, NOIR)  # gras + souligné
        set_run(p.add_run(txt), 8, False, False, GRIS)

    safe = re.sub(r'[\\/*?:"<>|]', "", num).replace("/", "-")
    nom = f"Commercial_Invoice_{safe}.docx"
    chemin = os.path.join(DOSSIER, nom)
    doc.save(chemin)
    return chemin
