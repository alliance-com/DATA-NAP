# -*- coding: utf-8 -*-
"""Génération lettres bancaires NAP — Word + PDF."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime, date
from io import BytesIO
import os
import re

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT

DOSSIER = "lettres_generees"
os.makedirs(DOSSIER, exist_ok=True)

HEADER_LINES = [
    "STE NAP SARL",
    "PLOT NO 14 GLODJIGBE ECONOMIC ZONE, REPUBLIC OF BENIN",
    "RCCM RB / COT / 21 B 30710, IFU 3202113334014",
    "TEL: +229 94940414, 97974506",
]

BLEU = RGBColor(27, 79, 114)
NOIR = RGBColor(20, 20, 20)
GRIS = RGBColor(80, 80, 80)


def _s(v):
    if v is None:
        return ""
    return str(v).strip()


def fmt_montant(n, devise="XOF"):
    try:
        n = float(n)
    except Exception:
        return _s(n)
    if devise in ("XOF", "CFA", "FCFA"):
        return f"{n:,.0f}".replace(",", " ")
    return f"{n:,.2f}".replace(",", " ")


def montant_en_lettres_simple(n, devise="XOF"):
    """Mention courte pour objet / corps (pas un convertisseur exhaustif)."""
    try:
        n = float(n)
    except Exception:
        return ""
    if devise in ("XOF", "CFA", "FCFA"):
        return f"{fmt_montant(n, devise)} francs CFA"
    if devise == "USD":
        return f"USD {fmt_montant(n, devise)}"
    if devise == "EUR":
        return f"EUR {fmt_montant(n, devise)}"
    return f"{fmt_montant(n, devise)} {devise}"


def set_run(run, size=11, bold=False, color=None, name="Times New Roman"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    if color:
        run.font.color.rgb = color


def add_line(doc, text, size=11, bold=False, center=False, space_after=4, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_run(r, size=size, bold=bold, color=color or NOIR)
    return p


def add_header_nap(doc):
    for i, line in enumerate(HEADER_LINES):
        add_line(doc, line, size=11 if i == 0 else 9, bold=(i == 0), center=True, space_after=1)
    # ligne
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "12")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), "1B4F72")
    pBdr.append(bot)
    pPr.append(pBdr)


def corps_selon_type(data):
    typ = _s(data.get("Type"))
    banque = _s(data.get("Banque")) or "Bank Manager"
    montant = data.get("Montant") or 0
    devise = _s(data.get("Devise") or "XOF")
    compte = _s(data.get("Compte_Credit"))
    benef = _s(data.get("Beneficiaire"))
    motif = _s(data.get("Motif"))
    ref_lc = _s(data.get("Ref_LC"))
    ref_fact = _s(data.get("Ref_Facture"))
    ref_tx = _s(data.get("Ref_Transaction"))
    corps_libre = _s(data.get("Corps_Libre"))
    objet = _s(data.get("Objet"))

    if corps_libre:
        return corps_libre

    if typ == "TIRAGE_CFA":
        mtxt = montant_en_lettres_simple(montant, "XOF")
        lines = [
            "Dear Sir,",
            "",
            f"We request you to approve a new Tirage for CFA {fmt_montant(montant, 'XOF')} ({mtxt})"
            + (f" for {motif}." if motif else "."),
        ]
        if benef:
            lines.append(f"Beneficiary / purpose related to: {benef}.")
        if compte:
            lines.append(f"Kindly credit the above mentioned Tirage to our account No {compte}.")
        lines += ["", "Your earliest action will be highly appreciated."]
        return "\n".join(lines)

    if typ == "PAIEMENT_FOURNISSEUR":
        mtxt = montant_en_lettres_simple(montant, devise)
        lines = [
            "Dear Sir,",
            "",
            f"We request you to pay {mtxt}"
            + (f" against Invoice No {ref_fact}" if ref_fact else "")
            + (f" to our supplier {benef}." if benef else "."),
        ]
        if ref_lc:
            lines.append(f"This payment is related to LC Reference Number {ref_lc}.")
        if motif:
            lines.append(motif)
        lines += ["", "Your earliest action will be highly appreciated."]
        return "\n".join(lines)

    if typ == "EXTENSION_LC":
        lines = [
            "Dear Sir,",
            "",
            f"We have issued an LC"
            + (f" of {montant_en_lettres_simple(montant, devise)}" if montant else "")
            + (f" to our supplier {benef}" if benef else "")
            + (f" against LC Reference Number {ref_lc}" if ref_lc else "")
            + ".",
            "",
            motif or "Due to operational constraints, we kindly request an extension of the LC validity.",
            "",
            "Your earliest action will be highly appreciated.",
        ]
        return "\n".join(lines)

    if typ == "ACCEPTATION_ECART":
        lines = [
            "Dear Sir,",
            "",
            "Our supplier has effected a shipment and the documents have reached the advising/negotiating bank.",
            "A discrepancy has been raised on the documents. After review, the discrepancy is acceptable to us.",
            "",
            "We hereby accept the discrepancy and request you to call for the documents and proceed with the settlement process under the Letter of Credit"
            + (f" {ref_lc}." if ref_lc else "."),
            "",
            "Thank you.",
        ]
        return "\n".join(lines)

    if typ == "REPONSE_JUSTIFICATIFS":
        lines = [
            "Monsieur le Directeur Général,",
            "",
            "Faisant suite à votre courrier"
            + (f" relatif à la transaction bancaire {ref_tx}" if ref_tx else "")
            + ", nous avons l'honneur de vous transmettre, en pièces jointes, les documents justificatifs concernés.",
            "",
            "À cet effet, vous trouverez joints à la présente, selon le cas :",
            "- Les Engagements de Change (EC) correspondants ;",
            "- Les déclarations y afférentes ;",
            "- Les BFU dûment payés ;",
            "- Les factures commerciales.",
            "",
            motif or "Ces documents sont transmis afin de justifier la réalisation et le traitement des opérations concernées.",
            "",
            "Nous restons à votre disposition pour toute information complémentaire.",
            "",
            "Veuillez agréer, Monsieur le Directeur Général, l'expression de notre parfaite considération.",
        ]
        return "\n".join(lines)

    # LETTRE_GENERALE
    lines = [
        "Dear Sir,",
        "",
        objet or motif or "Please find below our formal request.",
        "",
    ]
    if montant:
        lines.append(f"Amount: {montant_en_lettres_simple(montant, devise)}")
    if benef:
        lines.append(f"Related party: {benef}")
    if compte:
        lines.append(f"Account: {compte}")
    if motif and objet:
        lines.append(motif)
    lines += ["", "Your earliest action will be highly appreciated."]
    return "\n".join(lines)


def generer_lettre_word(data):
    doc = Document()
    for s in doc.sections:
        s.page_width = Cm(21)
        s.page_height = Cm(29.7)
        s.top_margin = Cm(1.5)
        s.bottom_margin = Cm(1.5)
        s.left_margin = Cm(2.0)
        s.right_margin = Cm(2.0)

    add_header_nap(doc)

    ref = _s(data.get("Reference"))
    d = data.get("Date_Lettre")
    if hasattr(d, "strftime"):
        date_txt = d.strftime("%d/%m/%Y")
    else:
        try:
            date_txt = pd_to_date(d)
        except Exception:
            date_txt = datetime.now().strftime("%d/%m/%Y")

    add_line(doc, f"{ref}          DATED: {date_txt}", size=11, bold=True, space_after=10)

    banque = _s(data.get("Banque"))
    if banque:
        add_line(doc, "To", size=11, bold=True, space_after=2)
        for part in banque.split("\n"):
            add_line(doc, part, size=11, bold=True, space_after=1)
        add_line(doc, "Cotonou, Benin", size=11, space_after=8)

    objet = _s(data.get("Objet"))
    if not objet:
        typ = _s(data.get("Type"))
        if typ == "TIRAGE_CFA":
            objet = f"Tirage for CFA {fmt_montant(data.get('Montant'), 'XOF')}"
        elif typ == "PAIEMENT_FOURNISSEUR":
            objet = "Request of payment to supplier"
        elif typ == "EXTENSION_LC":
            objet = "Extension of the Letter of Credit"
        elif typ == "REPONSE_JUSTIFICATIFS":
            objet = "Transmission des pièces justificatives"
        else:
            objet = "Formal request"

    add_line(doc, f"OBJET: {objet}", size=11, bold=True, space_after=4)
    add_line(doc, "KIND ATTENTION – BANK MANAGER", size=11, bold=True, space_after=10)

    corps = corps_selon_type(data)
    for para in corps.split("\n"):
        if para.strip() == "":
            doc.add_paragraph()
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.2
        r = p.add_run(para)
        set_run(r, size=11)

    add_line(doc, "", space_after=6)
    add_line(doc, "FOR STE NAP SARL", size=11, bold=True, space_after=2)
    sig = _s(data.get("Signataire") or "AUTHORIZED SIGNATORY")
    add_line(doc, sig, size=11, bold=True, space_after=2)

    safe_ref = re.sub(r"[^\w\-]+", "_", ref)[:40]
    nom = f"Lettre_{safe_ref}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    chemin = os.path.join(DOSSIER, nom)
    doc.save(chemin)
    return chemin


def pd_to_date(d):
    import pandas as pd
    ts = pd.to_datetime(d, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return datetime.now().strftime("%d/%m/%Y")
    return ts.strftime("%d/%m/%Y")


def generer_lettre_pdf(data):
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=14 * mm, bottomMargin=14 * mm,
        title=_s(data.get("Reference")), author="STE NAP SARL",
    )
    styles = getSampleStyleSheet()
    st_h = ParagraphStyle("h", parent=styles["Normal"], fontSize=11, alignment=TA_CENTER, fontName="Times-Bold", leading=13)
    st_s = ParagraphStyle("s", parent=styles["Normal"], fontSize=9, alignment=TA_CENTER, leading=11)
    st_b = ParagraphStyle("b", parent=styles["Normal"], fontSize=11, fontName="Times-Bold", leading=14, spaceAfter=3)
    st_n = ParagraphStyle("n", parent=styles["Normal"], fontSize=11, fontName="Times-Roman", leading=14, alignment=TA_JUSTIFY, spaceAfter=3)

    story = []
    story.append(Paragraph("STE NAP SARL", st_h))
    story.append(Paragraph("PLOT NO 14 GLODJIGBE ECONOMIC ZONE, REPUBLIC OF BENIN", st_s))
    story.append(Paragraph("RCCM RB / COT / 21 B 30710, IFU 3202113334014", st_s))
    story.append(Paragraph("TEL: +229 94940414, 97974506", st_s))
    story.append(Spacer(1, 2 * mm))
    # line
    line = Table([[""]], colWidths=[170 * mm])
    line.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, 0), 1, colors.Color(0.11, 0.31, 0.45))]))
    story.append(line)
    story.append(Spacer(1, 4 * mm))

    ref = _s(data.get("Reference"))
    d = data.get("Date_Lettre")
    date_txt = d.strftime("%d/%m/%Y") if hasattr(d, "strftime") else pd_to_date(d)
    story.append(Paragraph(f"{ref}&nbsp;&nbsp;&nbsp;&nbsp;DATED: {date_txt}", st_b))
    story.append(Spacer(1, 2 * mm))

    banque = _s(data.get("Banque"))
    if banque:
        story.append(Paragraph("<b>To</b>", st_b))
        for part in banque.split("\n"):
            story.append(Paragraph(part, st_b))
        story.append(Paragraph("Cotonou, Benin", st_n))
        story.append(Spacer(1, 2 * mm))

    objet = _s(data.get("Objet")) or "Formal request"
    story.append(Paragraph(f"<b>OBJET: {objet}</b>", st_b))
    story.append(Paragraph("<b>KIND ATTENTION – BANK MANAGER</b>", st_b))
    story.append(Spacer(1, 3 * mm))

    for para in corps_selon_type(data).split("\n"):
        if not para.strip():
            story.append(Spacer(1, 2 * mm))
        else:
            story.append(Paragraph(para.replace("&", "&amp;"), st_n))

    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph("<b>FOR STE NAP SARL</b>", st_b))
    story.append(Paragraph(f"<b>{_s(data.get('Signataire') or 'AUTHORIZED SIGNATORY')}</b>", st_b))

    doc.build(story)
    buf.seek(0)
    return buf.read()
