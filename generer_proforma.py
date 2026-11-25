# -*- coding: utf-8 -*-
"""Génération facture proforma NAP — Word moderne (FR / EN)."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
from io import BytesIO
import os
import re

try:
    import qrcode
except ImportError:
    qrcode = None

DOSSIER = "proformas_generes"
os.makedirs(DOSSIER, exist_ok=True)

BLEU = RGBColor(27, 79, 114)
BLEU_M = RGBColor(46, 134, 193)
CIEL = "5DADE2"
CIEL_H = "87CEEB"
VERT = RGBColor(34, 139, 34)
GRIS = RGBColor(55, 55, 55)
GRIS_C = RGBColor(120, 120, 120)
ROUGE = RGBColor(192, 57, 43)

I18N = {
    "FR": {
        "title": "FACTURE PROFORMA",
        "date_prefix": "Cotonou le ",
        "no": "N°",
        "client": "CLIENT",
        "address": "ADRESSE",
        "ifu": "IFU",
        "country": "PAYS",
        "product": "PRODUIT",
        "desc": "DESCRIPTION DU PRODUIT",
        "qty": "QUANTITÉ (Tonne)",
        "price": "PRIX (FCFA)",
        "amount": "MONTANT",
        "incoterm": "INCOTERM",
        "total_ht": "TOTAL HT",
        "aib": "AIB {t}%",
        "tva": "TVA {t} %",
        "ttc": "MONTANT TTC",
        "in_words": "MONTANT EN LETTRES",
        "pay_terms": "Condition de paiement",
        "validity": "Validité de l'offre",
        "delivery": "Livraison",
        "pay_mode": "Mode de paiement",
        "account": "Numéro de compte",
        "other": "Autres conditions",
        "seller": "Vendeur",
        "buyer": "Acheteur",
        "fcfa": "FCFA",
    },
    "EN": {
        "title": "PROFORMA INVOICE",
        "date_prefix": "Cotonou, ",
        "no": "No.",
        "client": "CLIENT",
        "address": "ADDRESS",
        "ifu": "TAX ID",
        "country": "COUNTRY",
        "product": "PRODUCT",
        "desc": "PRODUCT DESCRIPTION",
        "qty": "QUANTITY (MT)",
        "price": "PRICE (FCFA)",
        "amount": "AMOUNT",
        "incoterm": "INCOTERM",
        "total_ht": "SUBTOTAL",
        "aib": "AIB {t}%",
        "tva": "VAT {t} %",
        "ttc": "TOTAL AMOUNT",
        "in_words": "AMOUNT IN WORDS",
        "pay_terms": "Payment terms",
        "validity": "Offer validity",
        "delivery": "Delivery",
        "pay_mode": "Payment method",
        "account": "Bank account",
        "other": "Other conditions",
        "seller": "Seller",
        "buyer": "Buyer",
        "fcfa": "FCFA",
    },
}


def _t(lang, key, **kw):
    lang = "EN" if str(lang).upper().startswith("EN") else "FR"
    s = I18N[lang].get(key, I18N["FR"].get(key, key))
    return s.format(**kw) if kw else s


def _s(v):
    if v is None:
        return ""
    return str(v).strip()


def fmt_m(n):
    try:
        return f"{float(n):,.0f}".replace(",", " ")
    except Exception:
        return str(n)


def set_run(run, size=10, bold=False, color=None, name="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
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


def tight(p, before=2, after=2):
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.1


def qr_bytes(texte):
    if not qrcode:
        return None
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=6, border=1)
    qr.add_data(texte)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def generer_proforma_word(data, lang="FR"):
    lang = "EN" if str(lang).upper().startswith("EN") else "FR"
    doc = Document()
    for s in doc.sections:
        s.page_width = Cm(21)
        s.page_height = Cm(29.7)
        s.top_margin = Cm(1.2)
        s.bottom_margin = Cm(1.5)
        s.left_margin = Cm(1.4)
        s.right_margin = Cm(1.4)

    # En-tête + QR
    t_top = doc.add_table(rows=1, cols=2)
    borders(t_top, color=CIEL, size="4")
    c0, c1 = t_top.rows[0].cells
    shade(c0, CIEL_H)
    p = c0.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(p, 4, 0)
    set_run(p.add_run("STE NAP SARL"), 14, True, BLEU)
    p2 = c0.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(p2, 0, 2)
    set_run(
        p2.add_run(
            "PLOT NO 14 GLODJIGBE ECONOMIC ZONE, REPUBLIC OF BENIN\n"
            "RCCM RB/COT/21 B 30710  ·  IFU 3202113334014  ·  TEL +229 94 94 04 14 / 97 97 45 06"
        ),
        7, False, BLEU,
    )
    c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    num = _s(data.get("Numero"))
    client = _s(data.get("Client"))
    ttc = data.get("Montant_TTC") or 0
    texte_qr = f"STE NAP SARL|PROFORMA|{num}|{client}|{fmt_m(ttc)}|{lang}"
    buf = qr_bytes(texte_qr)
    if buf:
        try:
            c1.paragraphs[0].add_run().add_picture(buf, width=Cm(1.8))
        except Exception:
            pass

    # Titre
    titre = doc.add_paragraph()
    titre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(titre, 8, 2)
    set_run(titre.add_run(_t(lang, "title")), 14, True, BLEU_M)

    # Ligne bleu ciel
    ligne = doc.add_paragraph()
    tight(ligne, 0, 6)
    pPr = ligne._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "18")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), CIEL)
    pBdr.append(bot)
    pPr.append(pBdr)

    d = data.get("Date_Facture")
    if hasattr(d, "strftime"):
        date_txt = d.strftime("%d/%m/%Y")
    else:
        try:
            import pandas as pd
            date_txt = pd.to_datetime(d, errors="coerce", dayfirst=True).strftime("%d/%m/%Y")
        except Exception:
            date_txt = datetime.now().strftime("%d/%m/%Y")

    meta = doc.add_paragraph()
    tight(meta, 2, 4)
    set_run(meta.add_run(f"{_t(lang, 'date_prefix')}{date_txt}    ·    {_t(lang, 'no')} : {num}"), 10, True, GRIS)

    # Bloc client
    t_cli = doc.add_table(rows=1, cols=1)
    borders(t_cli)
    cell = t_cli.rows[0].cells[0]
    shade(cell, "EAF6FD")
    pc = cell.paragraphs[0]
    tight(pc, 3, 1)
    set_run(pc.add_run(f"{_t(lang, 'client')} : "), 9, True, BLEU)
    set_run(pc.add_run(_s(data.get("Client"))), 10, True, GRIS)
    for lab, key in [
        (_t(lang, "address"), "Adresse"),
        (_t(lang, "ifu"), "IFU_Client"),
        (_t(lang, "country"), "Pays"),
    ]:
        val = _s(data.get(key))
        if val:
            pp = cell.add_paragraph()
            tight(pp, 0, 1)
            set_run(pp.add_run(f"{lab} : "), 9, True, BLEU)
            set_run(pp.add_run(val), 9, False, GRIS)

    prod = doc.add_paragraph()
    tight(prod, 8, 4)
    set_run(prod.add_run(f"{_t(lang, 'product')} : "), 10, True, BLEU)
    set_run(prod.add_run(_s(data.get("Produit"))), 10, True, GRIS)

    # Tableau produit
    desc_lines = [_s(data.get("Description")) or _s(data.get("Produit"))]
    inc = _s(data.get("Incoterm"))
    cond = _s(data.get("Conditionnement"))
    if inc:
        desc_lines.append(f"{_t(lang, 'incoterm')} : {inc}")
    if cond:
        desc_lines.append(cond)
    desc = "\n".join([x for x in desc_lines if x])

    qte = data.get("Quantite_T") or 0
    pu = data.get("Prix_Unitaire") or 0
    ht = data.get("Montant_HT") or 0
    app_aib = _s(data.get("Appliquer_AIB")).upper() in ("OUI", "YES", "TRUE", "1")
    app_tva = _s(data.get("Appliquer_TVA")).upper() in ("OUI", "YES", "TRUE", "1")
    taux_aib = data.get("Taux_AIB") or 0
    taux_tva = data.get("Taux_TVA") or 0
    aib = data.get("Montant_AIB") or 0
    tva = data.get("Montant_TVA") or 0
    ttc = data.get("Montant_TTC") or 0

    # rows: header + product + HT + optional AIB + optional TVA + TTC
    extra = 1 + (1 if app_aib else 0) + (1 if app_tva else 0) + 1
    nrows = 1 + 1 + extra
    tbl = doc.add_table(rows=nrows, cols=4)
    borders(tbl)
    headers = [_t(lang, "desc"), _t(lang, "qty"), _t(lang, "price"), _t(lang, "amount")]
    for i, h in enumerate(headers):
        cell = tbl.rows[0].cells[i]
        shade(cell, CIEL_H)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p.add_run(h), 8, True, BLEU)

    vals = [desc, fmt_m(qte), fmt_m(pu), f"{fmt_m(ht)} {_t(lang, 'fcfa')}"]
    for i, v in enumerate(vals):
        cell = tbl.rows[1].cells[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i else WD_ALIGN_PARAGRAPH.LEFT
        # multi-line description
        parts = str(v).split("\n")
        set_run(p.add_run(parts[0]), 9, False, GRIS)
        for part in parts[1:]:
            pp = cell.add_paragraph()
            pp.alignment = p.alignment
            set_run(pp.add_run(part), 8, False, GRIS_C)

    r = 2
    # TOTAL HT
    cell0 = tbl.rows[r].cells[0]
    cell0.merge(tbl.rows[r].cells[2])
    set_run(cell0.paragraphs[0].add_run(_t(lang, "total_ht")), 9, True, BLEU)
    set_run(tbl.rows[r].cells[3].paragraphs[0].add_run(f"{fmt_m(ht)} {_t(lang, 'fcfa')}"), 9, True, GRIS)
    r += 1
    if app_aib:
        cell0 = tbl.rows[r].cells[0]
        cell0.merge(tbl.rows[r].cells[2])
        set_run(cell0.paragraphs[0].add_run(_t(lang, "aib", t=fmt_m(taux_aib))), 9, True, BLEU)
        set_run(tbl.rows[r].cells[3].paragraphs[0].add_run(f"{fmt_m(aib)} {_t(lang, 'fcfa')}"), 9, False, GRIS)
        r += 1
    if app_tva:
        cell0 = tbl.rows[r].cells[0]
        cell0.merge(tbl.rows[r].cells[2])
        set_run(cell0.paragraphs[0].add_run(_t(lang, "tva", t=fmt_m(taux_tva))), 9, True, BLEU)
        set_run(tbl.rows[r].cells[3].paragraphs[0].add_run(f"{fmt_m(tva)} {_t(lang, 'fcfa')}"), 9, False, GRIS)
        r += 1
    # TTC
    cell0 = tbl.rows[r].cells[0]
    cell0.merge(tbl.rows[r].cells[2])
    shade(cell0, "E8F5E9")
    shade(tbl.rows[r].cells[3], "E8F5E9")
    set_run(cell0.paragraphs[0].add_run(_t(lang, "ttc")), 10, True, VERT)
    set_run(tbl.rows[r].cells[3].paragraphs[0].add_run(f"{fmt_m(ttc)} {_t(lang, 'fcfa')}"), 11, True, VERT)

    # Montant en lettres
    lettres = _s(data.get("Montant_Lettres"))
    if lettres:
        pl = doc.add_paragraph()
        tight(pl, 8, 4)
        set_run(pl.add_run(f"{_t(lang, 'in_words')} : "), 9, True, BLEU)
        set_run(pl.add_run(lettres), 9, True, GRIS)

    # Conditions — libellés soulignés, police Courier
    def cond_line(label, value):
        if not _s(value):
            return
        p = doc.add_paragraph()
        tight(p, 2, 2)
        r_lab = p.add_run(f"{label} :")
        set_run(r_lab, 9, True, BLEU, name="Courier New")
        r_lab.font.underline = True
        r_val = p.add_run(" " + _s(value))
        set_run(r_val, 9, False, GRIS, name="Courier New")

    cond_line(_t(lang, "pay_terms"), data.get("Condition_Paiement"))
    cond_line(_t(lang, "validity"), data.get("Validite"))
    cond_line(_t(lang, "delivery"), data.get("Livraison"))
    cond_line(_t(lang, "pay_mode"), data.get("Mode_Paiement"))
    cond_line(_t(lang, "account"), data.get("Compte_Bancaire"))

    # Autres conditions — en rouge, police Courier
    autres = _s(data.get("Autres_Conditions"))
    if autres:
        p = doc.add_paragraph()
        tight(p, 6, 2)
        r = p.add_run(_t(lang, "other") + " :")
        set_run(r, 9, True, ROUGE, name="Courier New")
        r.font.underline = True
        for line in autres.split("\n"):
            if line.strip():
                pp = doc.add_paragraph()
                tight(pp, 1, 1)
                set_run(pp.add_run(line.strip()), 8, False, ROUGE, name="Courier New")

    # Signatures
    sig = doc.add_paragraph()
    tight(sig, 14, 4)
    t_sig = doc.add_table(rows=2, cols=2)
    borders(t_sig, color="CCCCCC", size="4")
    for i, (lab, nom) in enumerate([
        (_t(lang, "seller"), "NAP SARL"),
        (_t(lang, "buyer"), _s(data.get("Client"))),
    ]):
        c = t_sig.rows[0].cells[i]
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p.add_run(lab), 9, True, BLEU)
        c2 = t_sig.rows[1].cells[i]
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p2.add_run(nom + "\n\n\n_______________\nVisa / date"), 9, True, GRIS)

    # Pied
    foot = doc.add_paragraph()
    foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(foot, 10, 0)
    set_run(
        foot.add_run(
            "STE NAP SARL  ·  Glo-Djigbé  ·  RCCM RB/COT/21 B 30710  ·  IFU 3202113334014  ·  +229 94 94 04 14"
        ),
        7, False, GRIS_C,
    )

    safe = re.sub(r'[\\/*?:"<>|]', "", num).replace("/", "-")
    nom = f"Proforma_{safe}_{lang}_{datetime.now().strftime('%Y%m%d')}.docx"
    chemin = os.path.join(DOSSIER, nom)
    doc.save(chemin)
    return chemin


def generer_proforma_pdf(data, lang="FR"):
    """PDF A4 — même contenu que le Word, FR/EN."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image,
    )
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

    lang = "EN" if str(lang).upper().startswith("EN") else "FR"
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
        title=_s(data.get("Numero")), author="STE NAP SARL",
    )
    styles = getSampleStyleSheet()
    st_h = ParagraphStyle("h", parent=styles["Normal"], fontSize=13, alignment=TA_CENTER,
                          textColor=colors.Color(0.11, 0.31, 0.45), fontName="Helvetica-Bold", leading=15)
    st_s = ParagraphStyle("s", parent=styles["Normal"], fontSize=7, alignment=TA_CENTER,
                          textColor=colors.Color(0.11, 0.31, 0.45), leading=9)
    st_title = ParagraphStyle("t", parent=styles["Normal"], fontSize=13, alignment=TA_CENTER,
                              textColor=colors.Color(0.18, 0.53, 0.76), fontName="Helvetica-Bold", spaceBefore=6, spaceAfter=4)
    st_n = ParagraphStyle("n", parent=styles["Normal"], fontSize=9, leading=11, textColor=colors.Color(0.2, 0.2, 0.2))
    st_b = ParagraphStyle("b", parent=styles["Normal"], fontSize=9, leading=11, fontName="Helvetica-Bold",
                          textColor=colors.Color(0.11, 0.31, 0.45))
    st_foot = ParagraphStyle("f", parent=styles["Normal"], fontSize=7, alignment=TA_CENTER,
                             textColor=colors.grey, spaceBefore=8)

    num = _s(data.get("Numero"))
    client = _s(data.get("Client"))
    ttc = data.get("Montant_TTC") or 0
    d = data.get("Date_Facture")
    if hasattr(d, "strftime"):
        date_txt = d.strftime("%d/%m/%Y")
    else:
        try:
            import pandas as pd
            date_txt = pd.to_datetime(d, errors="coerce", dayfirst=True).strftime("%d/%m/%Y")
        except Exception:
            date_txt = datetime.now().strftime("%d/%m/%Y")

    story = []
    # header
    hdr = [
        [Paragraph("<b>STE NAP SARL</b>", st_h), ""],
    ]
    qr_buf = qr_bytes(f"STE NAP SARL|PROFORMA|{num}|{client}|{fmt_m(ttc)}|{lang}")
    if qr_buf:
        try:
            img = Image(qr_buf, width=18 * mm, height=18 * mm)
            hdr = [[
                Paragraph(
                    "<b>STE NAP SARL</b><br/>PLOT NO 14 GLODJIGBE ECONOMIC ZONE, BENIN<br/>"
                    "RCCM RB/COT/21 B 30710 · IFU 3202113334014 · TEL +229 94 94 04 14",
                    st_s,
                ),
                img,
            ]]
        except Exception:
            hdr = [[Paragraph("<b>STE NAP SARL</b><br/>Glo-Djigbé · BENIN", st_s), ""]]
    else:
        hdr = [[Paragraph(
            "<b>STE NAP SARL</b><br/>PLOT NO 14 GLODJIGBE · RCCM RB/COT/21 B 30710 · IFU 3202113334014",
            st_s,
        ), ""]]

    usable = 180 * mm
    t = Table(hdr, colWidths=[usable - 22 * mm, 22 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), colors.Color(0.53, 0.81, 0.92)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.Color(0.36, 0.68, 0.89)),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t)
    story.append(Paragraph(_t(lang, "title"), st_title))
    story.append(Paragraph(
        f"{_t(lang, 'date_prefix')}{date_txt} &nbsp;&nbsp;·&nbsp;&nbsp; {_t(lang, 'no')} : <b>{num}</b>",
        st_n,
    ))
    story.append(Spacer(1, 3 * mm))

    # client box
    cli_lines = [f"<b>{_t(lang, 'client')} :</b> {client}"]
    if _s(data.get("Adresse")):
        cli_lines.append(f"<b>{_t(lang, 'address')} :</b> {_s(data.get('Adresse'))}")
    if _s(data.get("IFU_Client")):
        cli_lines.append(f"<b>{_t(lang, 'ifu')} :</b> {_s(data.get('IFU_Client'))}")
    if _s(data.get("Pays")):
        cli_lines.append(f"<b>{_t(lang, 'country')} :</b> {_s(data.get('Pays'))}")
    tc = Table([[Paragraph("<br/>".join(cli_lines), st_n)]], colWidths=[usable])
    tc.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.Color(0.92, 0.96, 0.99)),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.Color(0.36, 0.68, 0.89)),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(tc)
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph(f"<b>{_t(lang, 'product')} :</b> {_s(data.get('Produit'))}", st_b))
    story.append(Spacer(1, 2 * mm))

    desc_lines = [_s(data.get("Description")) or _s(data.get("Produit"))]
    if _s(data.get("Incoterm")):
        desc_lines.append(f"{_t(lang, 'incoterm')} : {_s(data.get('Incoterm'))}")
    if _s(data.get("Conditionnement")):
        desc_lines.append(_s(data.get("Conditionnement")))
    desc = "<br/>".join(desc_lines)

    qte = data.get("Quantite_T") or 0
    pu = data.get("Prix_Unitaire") or 0
    ht = data.get("Montant_HT") or 0
    app_aib = _s(data.get("Appliquer_AIB")).upper() in ("OUI", "YES", "TRUE", "1")
    app_tva = _s(data.get("Appliquer_TVA")).upper() in ("OUI", "YES", "TRUE", "1")
    taux_aib = data.get("Taux_AIB") or 0
    taux_tva = data.get("Taux_TVA") or 0
    aib = data.get("Montant_AIB") or 0
    tva = data.get("Montant_TVA") or 0

    rows = [[
        Paragraph(f"<b>{_t(lang, 'desc')}</b>", st_n),
        Paragraph(f"<b>{_t(lang, 'qty')}</b>", st_n),
        Paragraph(f"<b>{_t(lang, 'price')}</b>", st_n),
        Paragraph(f"<b>{_t(lang, 'amount')}</b>", st_n),
    ], [
        Paragraph(desc, st_n),
        Paragraph(fmt_m(qte), st_n),
        Paragraph(fmt_m(pu), st_n),
        Paragraph(f"{fmt_m(ht)} {_t(lang, 'fcfa')}", st_n),
    ], [
        Paragraph(f"<b>{_t(lang, 'total_ht')}</b>", st_n), "", "",
        Paragraph(f"<b>{fmt_m(ht)} {_t(lang, 'fcfa')}</b>", st_n),
    ]]
    if app_aib:
        rows.append([
            Paragraph(f"<b>{_t(lang, 'aib', t=fmt_m(taux_aib))}</b>", st_n), "", "",
            Paragraph(f"{fmt_m(aib)} {_t(lang, 'fcfa')}", st_n),
        ])
    if app_tva:
        rows.append([
            Paragraph(f"<b>{_t(lang, 'tva', t=fmt_m(taux_tva))}</b>", st_n), "", "",
            Paragraph(f"{fmt_m(tva)} {_t(lang, 'fcfa')}", st_n),
        ])
    rows.append([
        Paragraph(f"<b>{_t(lang, 'ttc')}</b>", st_n), "", "",
        Paragraph(f"<b><font color='green'>{fmt_m(ttc)} {_t(lang, 'fcfa')}</font></b>", st_n),
    ])

    # merge labels for summary rows
    col_w = [usable * 0.42, usable * 0.16, usable * 0.20, usable * 0.22]
    tp = Table(rows, colWidths=col_w)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.Color(0.53, 0.81, 0.92)),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.Color(0.36, 0.68, 0.89)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("BACKGROUND", (0, -1), (-1, -1), colors.Color(0.91, 0.96, 0.91)),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("SPAN", (0, 2), (2, 2)),
    ]
    ridx = 3
    if app_aib:
        style_cmds.append(("SPAN", (0, ridx), (2, ridx)))
        ridx += 1
    if app_tva:
        style_cmds.append(("SPAN", (0, ridx), (2, ridx)))
        ridx += 1
    style_cmds.append(("SPAN", (0, ridx), (2, ridx)))
    tp.setStyle(TableStyle(style_cmds))
    story.append(tp)

    lettres = _s(data.get("Montant_Lettres"))
    if lettres:
        story.append(Spacer(1, 3 * mm))
        story.append(Paragraph(f"<b>{_t(lang, 'in_words')} :</b> {lettres}", st_n))

    st_cour = ParagraphStyle(
        "cour", parent=styles["Normal"], fontSize=9, leading=11,
        fontName="Courier", textColor=colors.Color(0.2, 0.2, 0.2),
    )
    st_cour_red = ParagraphStyle(
        "cour_red", parent=styles["Normal"], fontSize=8, leading=10,
        fontName="Courier", textColor=colors.Color(0.75, 0.22, 0.17),
    )
    st_cour_red_lab = ParagraphStyle(
        "cour_red_lab", parent=styles["Normal"], fontSize=9, leading=11,
        fontName="Courier-Bold", textColor=colors.Color(0.75, 0.22, 0.17),
    )

    def add_cond(label, value):
        if _s(value):
            story.append(Paragraph(
                f"<u><b>{label} :</b></u> {_s(value)}",
                st_cour,
            ))

    story.append(Spacer(1, 3 * mm))
    add_cond(_t(lang, "pay_terms"), data.get("Condition_Paiement"))
    add_cond(_t(lang, "validity"), data.get("Validite"))
    add_cond(_t(lang, "delivery"), data.get("Livraison"))
    add_cond(_t(lang, "pay_mode"), data.get("Mode_Paiement"))
    add_cond(_t(lang, "account"), data.get("Compte_Bancaire"))
    autres = _s(data.get("Autres_Conditions"))
    if autres:
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(f"<u><b>{_t(lang, 'other')} :</b></u>", st_cour_red_lab))
        for line in autres.split("\n"):
            if line.strip():
                story.append(Paragraph(line.strip(), st_cour_red))

    story.append(Spacer(1, 8 * mm))
    sig = Table([
        [Paragraph(f"<b>{_t(lang, 'seller')}</b>", st_n), Paragraph(f"<b>{_t(lang, 'buyer')}</b>", st_n)],
        [Paragraph("NAP SARL<br/><br/><br/>_______________", st_n),
         Paragraph(f"{client}<br/><br/><br/>_______________", st_n)],
    ], colWidths=[usable / 2, usable / 2])
    sig.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOX", (0, 0), (-1, -1), 0.3, colors.Color(0.8, 0.8, 0.8)),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(sig)
    story.append(Paragraph(
        "STE NAP SARL · Glo-Djigbé · RCCM RB/COT/21 B 30710 · IFU 3202113334014 · +229 94 94 04 14",
        st_foot,
    ))
    doc.build(story)
    buf.seek(0)
    return buf.read()
