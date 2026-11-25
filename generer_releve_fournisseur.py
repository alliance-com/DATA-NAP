# -*- coding: utf-8 -*-
"""Génération Fiche Relevé des Paiements NAP SARL — Word A4 + PDF A4 (FR/EN)."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree
from datetime import datetime
from io import BytesIO
import os
import re

try:
    import qrcode
except ImportError:
    qrcode = None

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT

DOSSIER_SORTIE = "releves_fournisseurs"
os.makedirs(DOSSIER_SORTIE, exist_ok=True)

BLEU_FONCE = RGBColor(27, 79, 114)
BLEU_MOYEN = RGBColor(46, 134, 193)
GRIS = RGBColor(55, 55, 55)
GRIS_CLAIR = RGBColor(110, 110, 110)
VERT = RGBColor(34, 139, 34)
ROUGE = RGBColor(192, 57, 43)
BLEU_CIEL = "5DADE2"

I18N = {
    "FR": {
        "title": "RELEVÉ DES PAIEMENTS",
        "supplier": "Fournisseur",
        "product": "Produit",
        "period": "Période",
        "period_from_to": "Du {d1} au {d2}",
        "ref_ops": "N° relevé / Nb ops",
        "ops": "{ref}  ·  {nb} opération(s)",
        "detail": "Détail des paiements",
        "headers": ["Date", "N° transaction", "Produit", "Description / Objet", "Montant", "Mode"],
        "total": "Total période : ",
        "signatures": "Signatures",
        "supplier_sig": "Fournisseur",
        "visa": "\n\nVisa / date\n_______________",
        "nb": (
            "NB : Ce relevé reprend les paiements enregistrés dans le suivi NAP pour le fournisseur "
            "et la période indiqués. Toute réclamation doit être signalée sous 8 jours."
        ),
        "scan": "Scan sécurisé NAP",
        "all_products": "Tous",
        "account_status": "État du compte au {date}",
        "balance_zero": "Solde : 0 FCFA — Compte soldé",
        "balance_due": "Reste à payer : {montant} FCFA",
        "total_paid": "Total payé sur la période : ",
    },
    "EN": {
        "title": "PAYMENT STATEMENT",
        "supplier": "Supplier",
        "product": "Product",
        "period": "Period",
        "period_from_to": "From {d1} to {d2}",
        "ref_ops": "Statement No. / Nb ops",
        "ops": "{ref}  ·  {nb} transaction(s)",
        "detail": "Payment details",
        "headers": ["Date", "Txn No.", "Product", "Description / Purpose", "Amount", "Mode"],
        "total": "Period total: ",
        "signatures": "Signatures",
        "supplier_sig": "Supplier",
        "visa": "\n\nVisa / date\n_______________",
        "nb": (
            "NB: This statement lists payments recorded in the NAP tracker for the supplier "
            "and period shown. Any claim must be raised within 8 days."
        ),
        "scan": "NAP secure scan",
        "all_products": "All",
        "account_status": "Account status as at {date}",
        "balance_zero": "Balance: 0 FCFA — Account settled",
        "balance_due": "Amount due: {montant} FCFA",
        "total_paid": "Total paid for the period: ",
    },
}


def _t(lang, key, **kw):
    lang = "EN" if str(lang).upper().startswith("EN") else "FR"
    s = I18N[lang].get(key, I18N["FR"].get(key, key))
    return s.format(**kw) if kw else s


def set_run_font(run, size=9, bold=False, color=None, name="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name
    r = run._element.rPr
    if r is not None:
        r.rFonts.set(qn("w:eastAsia"), name)
    if color:
        run.font.color.rgb = color


def shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_margins(cell, top=30, bottom=30, left=50, right=50):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table, color="5DADE2", size="8"):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), size)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    tblPr.append(borders)


def set_col_widths(table, widths_cm):
    for row in table.rows:
        for idx, w in enumerate(widths_cm):
            if idx < len(row.cells):
                row.cells[idx].width = Cm(w)


def tight_para(p, before=1, after=1):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.0


def generer_qr_png_bytes(texte):
    if qrcode is None:
        raise ImportError("pip install qrcode[pil]")
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=5, border=1)
    qr.add_data(texte)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def ajouter_filigrane(doc, reference):
    date_str = datetime.now().strftime("%d/%m/%Y")
    texte = f"STE NAP SARL  •  {reference}  •  {date_str}"
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False
    for p in header.paragraphs:
        p.clear()
    paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    r = OxmlElement("w:r")
    paragraph._p.append(r)
    rPr = OxmlElement("w:rPr")
    rPr.append(OxmlElement("w:noProof"))
    r.append(rPr)
    pict = OxmlElement("w:pict")
    r.append(pict)
    shape_xml = (
        '<v:shape xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'id="FiligraneNAP" type="#_x0000_t136" '
        'style="position:absolute;margin-left:0;margin-top:0;width:480pt;height:50pt;'
        'z-index:-251658240;mso-position-horizontal:center;mso-position-horizontal-relative:page;'
        'mso-position-vertical:center;mso-position-vertical-relative:page;" '
        'fillcolor="#C8C8C8" stroked="f">'
        '<v:fill opacity=".30"/>'
        '<v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="A0A0A0"/></w:rPr>'
        f'<w:t xml:space="preserve">{texte}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>'
    )
    pict.append(etree.fromstring(shape_xml))


def _label_val(cell, label, value, label_size=8, val_size=9):
    cell.text = ""
    set_cell_margins(cell)
    p = cell.paragraphs[0]
    tight_para(p, 1, 0)
    r = p.add_run(f"{label}\n")
    set_run_font(r, size=label_size, bold=True, color=BLEU_MOYEN)
    r2 = p.add_run(str(value) if value not in (None, "nan", "None", "") else "—")
    set_run_font(r2, size=val_size, color=GRIS)


def pied_page_nap(doc):
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    for p in footer.paragraphs:
        p.clear()
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight_para(p, 2, 0)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "12")
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), BLEU_CIEL)
    pBdr.append(top)
    pPr.append(pBdr)
    set_run_font(
        p.add_run(
            "STE NAP SARL  •  PLOT NO 14 GLODJIGBE ECONOMIC ZONE — BENIN\n"
            "RCCM RB/COT/21 B 30710  •  IFU 3202113334014  •  TEL +229 01 94 94 04 14 / +229 01 97 97 45 06"
        ),
        size=7,
        color=GRIS_CLAIR,
    )


def fmt_m(n):
    try:
        return f"{float(n):,.0f}".replace(",", " ")
    except Exception:
        return str(n)


def generer_releve_fournisseur(data, lang="FR"):
    lang = "EN" if str(lang).upper().startswith("EN") else "FR"
    doc = Document()
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(1.0)
        section.bottom_margin = Cm(1.8)
        section.left_margin = Cm(1.3)
        section.right_margin = Cm(1.3)
        section.header_distance = Cm(0.3)
        section.footer_distance = Cm(0.5)

    frs = str(data.get("Fournisseur") or "").strip()
    d1 = str(data.get("Periode_Debut") or "")
    d2 = str(data.get("Periode_Fin") or "")
    ref = f"REL-{re.sub(r'[^A-Z0-9]', '', frs.upper())[:12]}-{datetime.now().strftime('%Y%m%d')}"
    ajouter_filigrane(doc, ref)
    pied_page_nap(doc)

    total = float(data.get("Total") or 0)
    nb = int(data.get("Nb_operations") or 0)

    t_top = doc.add_table(rows=1, cols=2)
    t_top.autofit = True
    c_left, c_right = t_top.rows[0].cells
    shade_cell(c_left, "87CEEB")
    set_cell_margins(c_left, top=50, bottom=50, left=80, right=80)
    p = c_left.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight_para(p, 0, 0)
    set_run_font(p.add_run("STE NAP SARL"), size=14, bold=True, color=BLEU_FONCE)
    p2 = c_left.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight_para(p2, 0, 0)
    set_run_font(
        p2.add_run(
            "PLOT NO 14 GLODJIGBE ECONOMIC ZONE • BENIN\n"
            "RCCM RB/COT/21 B 30710 • IFU 3202113334014 • TEL +229 01 94 94 04 14"
        ),
        size=7,
        color=BLEU_FONCE,
    )
    c_right.text = ""
    set_cell_margins(c_right, top=20, bottom=20, left=20, right=20)
    pq = c_right.paragraphs[0]
    pq.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    texte_qr = f"STE NAP SARL|RELEVE|{frs}|{d1}|{d2}|{fmt_m(total)}|{nb}|{lang}"
    if qrcode:
        try:
            pq.add_run().add_picture(generer_qr_png_bytes(texte_qr), width=Cm(1.7))
        except Exception:
            pass
    pl = c_right.add_paragraph()
    pl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tight_para(pl, 0, 0)
    set_run_font(pl.add_run(_t(lang, "scan")), size=6, color=GRIS_CLAIR)

    titre = doc.add_paragraph()
    titre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight_para(titre, 6, 2)
    set_run_font(titre.add_run(_t(lang, "title")), size=13, bold=True, color=BLEU_MOYEN)

    ligne = doc.add_paragraph()
    tight_para(ligne, 0, 4)
    pPr = ligne._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "18")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), BLEU_CIEL)
    pBdr.append(bottom)
    pPr.append(pBdr)

    t_info = doc.add_table(rows=2, cols=4)
    t_info.autofit = True
    produit_f = str(data.get("Produit") or "").strip() or _t(lang, "all_products")
    pairs = [
        (0, 0, (_t(lang, "supplier"), frs)),
        (0, 2, (_t(lang, "product"), produit_f)),
        (1, 0, (_t(lang, "period"), _t(lang, "period_from_to", d1=d1, d2=d2))),
        (1, 2, (_t(lang, "ref_ops"), _t(lang, "ops", ref=ref, nb=nb))),
    ]
    for row_i, col_i, (lab, val) in pairs:
        cell = t_info.rows[row_i].cells[col_i]
        cell2 = t_info.rows[row_i].cells[col_i + 1]
        cell.merge(cell2)
        _label_val(cell, lab, val)

    sec = doc.add_paragraph()
    tight_para(sec, 6, 2)
    set_run_font(sec.add_run(_t(lang, "detail")), size=10, bold=True, color=BLEU_MOYEN)

    lignes = data.get("Lignes") or []
    n = len(lignes)
    if n <= 8:
        font_h, font_b, m_tb, m_lr, ls = 8, 8, 28, 35, 1.0
    elif n <= 15:
        font_h, font_b, m_tb, m_lr, ls = 7, 7, 18, 25, 0.95
    elif n <= 25:
        font_h, font_b, m_tb, m_lr, ls = 6.5, 6.5, 10, 18, 0.9
    else:
        font_h, font_b, m_tb, m_lr, ls = 6, 6, 6, 12, 0.85

    t_det = doc.add_table(rows=1 + max(n, 1), cols=6)
    t_det.autofit = False
    set_table_borders(t_det, color=BLEU_CIEL, size="6" if n > 15 else "8")
    set_col_widths(t_det, [2.2, 2.7, 2.0, 6.0, 2.5, 2.0])

    headers = _t(lang, "headers")
    for i, h in enumerate(headers):
        cell = t_det.rows[0].cells[i]
        shade_cell(cell, "87CEEB")
        set_cell_margins(cell, top=m_tb, bottom=m_tb, left=m_lr, right=m_lr)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tight_para(p, 0, 0)
        p.paragraph_format.line_spacing = ls
        set_run_font(p.add_run(h), size=font_h, bold=True, color=BLEU_FONCE)

    if not lignes:
        for i in range(6):
            cell = t_det.rows[1].cells[i]
            set_cell_margins(cell, top=m_tb, bottom=m_tb, left=m_lr, right=m_lr)
            p = cell.paragraphs[0]
            tight_para(p, 0, 0)
            set_run_font(p.add_run("—"), size=font_b, color=GRIS_CLAIR)
    else:
        for r_idx, L in enumerate(lignes, start=1):
            try:
                m_txt = fmt_m(float(L.get("Montant") or 0))
            except Exception:
                m_txt = str(L.get("Montant", "—"))
            desc = str(L.get("Description") or L.get("Ref_doc") or "—")
            if n > 15 and len(desc) > 55:
                desc = desc[:52] + "…"
            elif n > 8 and len(desc) > 70:
                desc = desc[:67] + "…"
            vals = [
                str(L.get("Date") or "—"),
                str(L.get("No_transaction") or "—"),
                str(L.get("Produit") or "—"),
                desc,
                m_txt,
                str(L.get("Mode") or "—"),
            ]
            fill = "EAF6FD" if r_idx % 2 == 0 else "FFFFFF"
            for i, v in enumerate(vals):
                cell = t_det.rows[r_idx].cells[i]
                shade_cell(cell, fill)
                set_cell_margins(cell, top=m_tb, bottom=m_tb, left=m_lr, right=m_lr)
                p = cell.paragraphs[0]
                tight_para(p, 0, 0)
                p.paragraph_format.line_spacing = ls
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in (0, 1, 4, 5) else WD_ALIGN_PARAGRAPH.LEFT
                set_run_font(p.add_run(v), size=font_b, bold=(i == 4), color=(VERT if i == 4 else GRIS))

    tot = doc.add_paragraph()
    tight_para(tot, 6, 2)
    tot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_run_font(tot.add_run(_t(lang, "total")), size=9, bold=True, color=BLEU_FONCE)
    set_run_font(tot.add_run(f"XOF {fmt_m(total)}"), size=12, bold=True, color=VERT)

    # État du compte à la date de fin de période
    try:
        solde = float(data.get("Solde_Restant") if data.get("Solde_Restant") is not None else 0)
    except Exception:
        solde = 0.0
    date_etat = str(data.get("Periode_Fin") or "").strip() or datetime.now().strftime("%d/%m/%Y")
    etat = doc.add_paragraph()
    tight_para(etat, 4, 2)
    etat.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_run_font(etat.add_run(_t(lang, "account_status", date=date_etat) + "  —  "), size=9, bold=True, color=BLEU_FONCE)
    if abs(solde) < 0.5:
        set_run_font(etat.add_run(_t(lang, "balance_zero")), size=10, bold=True, color=VERT)
    else:
        set_run_font(
            etat.add_run(_t(lang, "balance_due", montant=fmt_m(solde))),
            size=10, bold=True, color=ROUGE,
        )

    sig = doc.add_paragraph()
    tight_para(sig, 6 if n > 15 else (10 if n > 8 else 14), 3)
    set_run_font(sig.add_run(_t(lang, "signatures")), size=9, bold=True, color=BLEU_MOYEN)

    t_sig = doc.add_table(rows=3, cols=2)
    t_sig.autofit = True
    for i, (lab, sous) in enumerate([
        (frs or _t(lang, "supplier_sig"), _t(lang, "supplier_sig")),
        ("NAP SARL", "NAP SARL"),
    ]):
        cell0 = t_sig.rows[0].cells[i]
        p0 = cell0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tight_para(p0, 1, 1)
        set_run_font(p0.add_run(sous), size=8, bold=True, color=BLEU_FONCE)

        cell1 = t_sig.rows[1].cells[i]
        p1 = cell1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tight_para(p1, 1, 1)
        set_run_font(p1.add_run(lab), size=9, bold=True, color=GRIS)

        cell2 = t_sig.rows[2].cells[i]
        p2 = cell2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tight_para(p2, 1, 1)
        set_run_font(p2.add_run(_t(lang, "visa")), size=8, color=GRIS_CLAIR)

    note = doc.add_paragraph()
    tight_para(note, 10, 1)
    set_run_font(note.add_run(_t(lang, "nb")), size=7, bold=True, color=ROUGE)

    safe = re.sub(r'[\\/*?:"<>|]', "", frs) or "FOURNISSEUR"
    safe = re.sub(r"\s+", "_", safe)[:40]
    nom = f"Releve_{safe}_{lang}_{datetime.now().strftime('%Y%m%d')}.docx"
    chemin = os.path.join(DOSSIER_SORTIE, nom)
    doc.save(chemin)
    return chemin


def generer_releve_pdf(data, lang="FR"):
    """PDF A4 du même relevé (FR/EN). Retourne bytes."""
    lang = "EN" if str(lang).upper().startswith("EN") else "FR"
    frs = str(data.get("Fournisseur") or "").strip()
    d1 = str(data.get("Periode_Debut") or "")
    d2 = str(data.get("Periode_Fin") or "")
    produit_f = str(data.get("Produit") or "").strip() or _t(lang, "all_products")
    total = float(data.get("Total") or 0)
    nb = int(data.get("Nb_operations") or 0)
    ref = f"REL-{re.sub(r'[^A-Z0-9]', '', frs.upper())[:12]}-{datetime.now().strftime('%Y%m%d')}"
    lignes = data.get("Lignes") or []

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=12 * mm, bottomMargin=18 * mm,
        title=_t(lang, "title"), author="STE NAP SARL",
    )
    styles = getSampleStyleSheet()
    st_title = ParagraphStyle("t", parent=styles["Heading1"], fontSize=14, textColor=colors.Color(0.11, 0.31, 0.45), alignment=TA_CENTER, spaceAfter=4)
    st_sub = ParagraphStyle("s", parent=styles["Normal"], fontSize=8, textColor=colors.Color(0.3, 0.3, 0.3), alignment=TA_CENTER)
    st_cell = ParagraphStyle("c", parent=styles["Normal"], fontSize=7, leading=9)
    st_head = ParagraphStyle("h", parent=styles["Normal"], fontSize=7, textColor=colors.Color(0.11, 0.31, 0.45), fontName="Helvetica-Bold", alignment=TA_CENTER)
    st_foot = ParagraphStyle("f", parent=styles["Normal"], fontSize=6.5, textColor=colors.Color(0.4, 0.4, 0.4), alignment=TA_CENTER)
    st_note = ParagraphStyle("n", parent=styles["Normal"], fontSize=7, textColor=colors.Color(0.75, 0.22, 0.17), fontName="Helvetica-Bold")

    story = []
    story.append(Paragraph("STE NAP SARL", st_title))
    story.append(Paragraph(_t(lang, "title"), st_title))
    story.append(Paragraph(
        f"{_t(lang, 'supplier')}: <b>{frs}</b> &nbsp;|&nbsp; "
        f"{_t(lang, 'product')}: <b>{produit_f}</b><br/>"
        f"{_t(lang, 'period')}: {_t(lang, 'period_from_to', d1=d1, d2=d2)} &nbsp;|&nbsp; "
        f"{_t(lang, 'ops', ref=ref, nb=nb)}",
        st_sub,
    ))
    story.append(Spacer(1, 4 * mm))

    headers = _t(lang, "headers")
    n = len(lignes)
    fs = 7 if n <= 15 else 6
    st_cell.fontSize = fs
    data_tbl = [[Paragraph(h, st_head) for h in headers]]
    for L in lignes:
        try:
            m_txt = fmt_m(float(L.get("Montant") or 0))
        except Exception:
            m_txt = str(L.get("Montant", "—"))
        desc = str(L.get("Description") or L.get("Ref_doc") or "—")
        if len(desc) > 60:
            desc = desc[:57] + "…"
        data_tbl.append([
            Paragraph(str(L.get("Date") or "—"), st_cell),
            Paragraph(str(L.get("No_transaction") or "—"), st_cell),
            Paragraph(str(L.get("Produit") or "—"), st_cell),
            Paragraph(desc, st_cell),
            Paragraph(m_txt, st_cell),
            Paragraph(str(L.get("Mode") or "—"), st_cell),
        ])
    if len(data_tbl) == 1:
        data_tbl.append([Paragraph("—", st_cell)] * 6)

    usable = A4[0] - 28 * mm
    widths = [usable * x for x in (0.12, 0.16, 0.12, 0.34, 0.14, 0.12)]
    bleu_c = colors.Color(0.53, 0.81, 0.92)
    bleu_l = colors.Color(0.92, 0.96, 0.99)
    t = Table(data_tbl, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), bleu_c),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.Color(0.7, 0.8, 0.9)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, bleu_l]),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("ALIGN", (4, 1), (4, -1), "RIGHT"),
    ]))
    story.append(t)
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(
        f"<b>{_t(lang, 'total')}</b> <font color='green'><b>XOF {fmt_m(total)}</b></font>",
        ParagraphStyle("tot", parent=styles["Normal"], fontSize=11, alignment=TA_RIGHT),
    ))
    try:
        solde = float(data.get("Solde_Restant") if data.get("Solde_Restant") is not None else 0)
    except Exception:
        solde = 0.0
    date_etat = str(data.get("Periode_Fin") or "").strip()
    if abs(solde) < 0.5:
        etat_txt = f"<b>{_t(lang, 'account_status', date=date_etat)}</b> — <font color='green'><b>{_t(lang, 'balance_zero')}</b></font>"
    else:
        etat_txt = (
            f"<b>{_t(lang, 'account_status', date=date_etat)}</b> — "
            f"<font color='red'><b>{_t(lang, 'balance_due', montant=fmt_m(solde))}</b></font>"
        )
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph(etat_txt, ParagraphStyle("etat", parent=styles["Normal"], fontSize=10, alignment=TA_LEFT)))
    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph(_t(lang, "signatures"), ParagraphStyle("sig", parent=styles["Normal"], fontSize=9, textColor=colors.Color(0.18, 0.53, 0.76), fontName="Helvetica-Bold")))
    sig_data = [
        [Paragraph(f"<b>{_t(lang, 'supplier_sig')}</b>", st_cell), Paragraph("<b>NAP SARL</b>", st_cell)],
        [Paragraph(frs or "—", st_cell), Paragraph("NAP SARL", st_cell)],
        [Paragraph(_t(lang, "visa").replace("\n", "<br/>"), st_cell), Paragraph(_t(lang, "visa").replace("\n", "<br/>"), st_cell)],
    ]
    ts = Table(sig_data, colWidths=[usable / 2, usable / 2])
    ts.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(ts)
    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph(_t(lang, "nb"), st_note))
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(
        "STE NAP SARL  •  PLOT NO 14 GLODJIGBE ECONOMIC ZONE — BENIN  •  "
        "RCCM RB/COT/21 B 30710  •  IFU 3202113334014  •  TEL +229 01 94 94 04 14",
        st_foot,
    ))
    doc.build(story)
    buf.seek(0)
    return buf.read()
