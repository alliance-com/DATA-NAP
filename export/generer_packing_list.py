# -*- coding: utf-8 -*-
"""
Packing List NAP — strictement calé sur Packing_List_PL-MODEL.docx
Modèle attendu : DATA/export/Packing_List_PL-MODEL.docx
(ou même dossier que ce script)
"""
from __future__ import annotations

import os
import re
from copy import deepcopy
from datetime import date

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Twips, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

_EXPORT_DIR = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_EXPORT_DIR)
DOSSIER = os.path.join(_ROOT, "packing_lists_generes")
os.makedirs(DOSSIER, exist_ok=True)

# Chemins modèle (ordre de recherche)
MODELE_CANDIDATS = [
    os.path.join(_EXPORT_DIR, "Packing_List_PL-MODEL.docx"),
    os.path.join(_ROOT, "export", "Packing_List_PL-MODEL.docx"),
    os.path.join(_ROOT, "Packing_List_PL-MODEL.docx"),
    os.path.join(_EXPORT_DIR, "Packing_List_PL-MODEL.docx".lower()),
]

# Logo NAP — ordre de recherche
LOGO_CANDIDATS = [
    os.path.join(_EXPORT_DIR, "logo_nap.png"),
    os.path.join(_EXPORT_DIR, "logo_nap.jpg"),
    os.path.join(_EXPORT_DIR, "logo.png"),
    os.path.join(_EXPORT_DIR, "logo.jpg"),
    os.path.join(_EXPORT_DIR, "LOGO_NAP.png"),
    os.path.join(_ROOT, "logo_nap.png"),
    os.path.join(_ROOT, "logo_nap.jpg"),
    os.path.join(_ROOT, "logo.png"),
    os.path.join(_ROOT, "assets", "logo_nap.png"),
    os.path.join(_ROOT, "assets", "logo.png"),
    os.path.join(_ROOT, "export", "logo_nap.png"),
]


def find_logo():
    for path in LOGO_CANDIDATS:
        if os.path.isfile(path):
            return path
    return None

NOIR = RGBColor(0, 0, 0)
FONT = "Times New Roman"
BLEU = RGBColor(27, 79, 114)
GRIS_C = RGBColor(120, 120, 120)


def qr_bytes(texte):
    try:
        import qrcode
        from io import BytesIO
        qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=6, border=1)
        qr.add_data(texte)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        buf = BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return buf
    except Exception:
        return None


def pied_page(doc):
    """Pied de page fixe en bas de page."""
    section = doc.sections[0]
    section.footer_distance = Cm(0.6)
    footer = section.footer
    footer.is_linked_to_previous = False
    for para in list(footer.paragraphs):
        para.clear()
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(fp, 0, 0)
    # ligne haute
    pPr = fp._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "8")
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), "AAAAAA")
    pBdr.append(top)
    pPr.append(pBdr)
    set_run(
        fp.add_run(
            "STE NAP SARL  ·  PLOT NO 14 GLODJIGBE ECONOMIC ZONE  ·  "
            "RCCM RB/COT/21 B 30710  ·  IFU 3202113334014  ·  TEL +229 94 94 04 14 / 97 97 45 06"
        ),
        7, False, GRIS_C,
    )
    fp2 = footer.add_paragraph()
    fp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(fp2, 0, 0)
    set_run(fp2.add_run("Packing List — Document authentifié NAP SARL"), 6, False, GRIS_C)



def _s(v):
    if v is None:
        return ""
    try:
        import pandas as pd
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def _fmt_date(d):
    if d is None or d == "":
        return ""
    try:
        import pandas as pd
        ts = pd.to_datetime(d, errors="coerce", dayfirst=True)
        if pd.isna(ts):
            return _s(d)
        return ts.strftime("%d/%m/%Y")
    except Exception:
        return _s(d)


def _fmt_w(n):
    try:
        v = float(n)
        if abs(v - round(v)) < 0.01:
            return f"{int(round(v)):,}".replace(",", " ")
        return f"{v:,.2f}".replace(",", " ")
    except Exception:
        return _s(n)


def set_run(run, size=9, bold=False, color=None, name=FONT):
    run.font.size = Pt(size)
    run.font.bold = bool(bold)
    run.font.name = name
    try:
        run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    except Exception:
        pass
    if color is not None:
        run.font.color.rgb = color


def tight(p, before=0, after=0):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.0


def shade(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


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


def set_cell_margins(cell, top=30, bottom=30, left=40, right=40):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for edge, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)


def set_row_height(row, cm_h=0.55):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(int(cm_h * 567)))  # approx twips
    trHeight.set(qn("w:hRule"), "atLeast")
    trPr.append(trHeight)


def clear_cell(cell):
    for p in cell.paragraphs:
        for r in p.runs:
            r.text = ""
        if p.text:
            p.text = ""


def write_lines(cell, lines, sizes=None, bolds=None, center=False):
    """Écrit plusieurs lignes dans une cellule (1er § réutilisé)."""
    clear_cell(cell)
    sizes = sizes or [8] * len(lines)
    bolds = bolds or [False] * len(lines)
    first = True
    for i, ln in enumerate(lines):
        if first:
            p = cell.paragraphs[0]
            first = False
        else:
            p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
        tight(p, 0, 0)
        set_run(p.add_run(_s(ln)), sizes[i] if i < len(sizes) else 8,
                bolds[i] if i < len(bolds) else False, NOIR)


def find_modele():
    for p in MODELE_CANDIDATS:
        if os.path.isfile(p):
            return p
    return None


def _page_setup(doc):
    for s in doc.sections:
        s.page_width = Cm(21.0)
        s.page_height = Cm(29.7)
        s.top_margin = Cm(1.0)
        s.bottom_margin = Cm(1.8)
        s.left_margin = Cm(1.3)
        s.right_margin = Cm(1.3)


def generer_packing_list(data, containers):
    """
    Génère le Packing List.
    - Si le modèle Word est présent, on reprend ses marges / section.
    - Disposition strictement identique au modèle PL-MODEL.
    - Nombre de lignes conteneurs = len(containers) du booking.
    """
    modele = find_modele()
    if modele:
        # Nouveau document vide mais on pourrait copier styles ; on fixe A4 comme le modèle
        doc = Document()
        try:
            mdoc = Document(modele)
            # reprendre marges du modèle
            ms, ss = mdoc.sections[0], doc.sections[0]
            ss.page_width = ms.page_width
            ss.page_height = ms.page_height
            ss.top_margin = ms.top_margin
            ss.bottom_margin = max(ms.bottom_margin, Cm(1.8))
            ss.left_margin = ms.left_margin
            ss.right_margin = ms.right_margin
        except Exception:
            _page_setup(doc)
    else:
        doc = Document()
        _page_setup(doc)

    ctrs = list(containers or [])
    n_ctr = len(ctrs)

    # ========== PIED DE PAGE ==========
    pied_page(doc)

    # ========== EN-TÊTE : LOGO | SOCIÉTÉ | QR ==========
    logo_path = find_logo()
    t_top = doc.add_table(rows=1, cols=3)
    borders(t_top, color="CCCCCC", size="4")
    c_logo, c_mid, c_right = t_top.rows[0].cells
    set_cell_margins(c_logo, 30, 30, 30, 20)
    set_cell_margins(c_mid, 40, 40, 40, 30)
    set_cell_margins(c_right, 20, 20, 20, 30)
    shade(c_mid, "EAF2F8")

    # Largeurs approximatives
    c_logo.width = Cm(2.8)
    c_mid.width = Cm(12.0)
    c_right.width = Cm(2.8)

    # --- Logo ---
    clear_cell(c_logo)
    plogo = c_logo.paragraphs[0]
    plogo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(plogo, 2, 2)
    if logo_path:
        try:
            plogo.add_run().add_picture(logo_path, width=Cm(2.2))
        except Exception:
            set_run(plogo.add_run("NAP"), 10, True, BLEU)
    else:
        set_run(plogo.add_run("NAP"), 11, True, BLEU)

    # --- Société ---
    write_lines(
        c_mid,
        [
            "STE NAP SARL",
            "PLOT NO 14 GLODJIGBE ECONOMIC ZONE, REPUBLIC OF BENIN",
            "RCCM RB / COT / 21 B 30710, IFU 3202113334014",
            "TEL + 229 94940414, 97974506",
        ],
        sizes=[12, 8, 8, 8],
        bolds=[True, True, True, True],
        center=True,
    )
    for para in c_mid.paragraphs:
        for run in para.runs:
            run.font.color.rgb = BLEU

    # --- QR ---
    clear_cell(c_right)
    pr = c_right.paragraphs[0]
    pr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(pr, 2, 2)
    bl = _s(data.get("BL_No"))
    num = _s(data.get("Packing_No") or bl or "PL")
    payload = (
        f"NAP-PL|{num}|BL:{bl}|{_s(data.get('Consignee'))}|"
        f"{_s(data.get('Vessel_Name'))}|CTR:{n_ctr}"
    )
    buf = qr_bytes(payload)
    if buf:
        try:
            pr.add_run().add_picture(buf, width=Cm(2.0))
        except Exception:
            set_run(pr.add_run("QR"), 8, False, NOIR)
    else:
        set_run(pr.add_run(""), 8, False, NOIR)

    # Titre
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(p, 8, 6)
    set_run(p.add_run("PACKING LIST"), 13, True, NOIR)

    # ========== SHIPPER | NOTIFY ==========
    t = doc.add_table(rows=1, cols=2)
    borders(t)
    t.autofit = True
    c0, c1 = t.rows[0].cells
    set_cell_margins(c0, 40, 40, 50, 40)
    set_cell_margins(c1, 40, 40, 50, 40)
    write_lines(
        c0,
        ["SHIPPER", _s(data.get("Shipper")), _s(data.get("Shipper_Address"))],
        sizes=[8, 8, 8],
        bolds=[True, False, False],
    )
    notify_lines = ["NOTIFY PARTY", _s(data.get("Notify"))]
    na = _s(data.get("Notify_Address"))
    if na:
        notify_lines.extend([x.strip() for x in na.split("\n") if x.strip()])
    write_lines(c1, notify_lines, sizes=[8] * len(notify_lines), bolds=[True] + [False] * (len(notify_lines) - 1))

    # ========== CONSIGNEE ==========
    t = doc.add_table(rows=1, cols=1)
    borders(t)
    c = t.rows[0].cells[0]
    set_cell_margins(c, 40, 40, 50, 40)
    cons = ["CONSIGNEE", _s(data.get("Consignee"))]
    ca = _s(data.get("Consignee_Address"))
    if ca:
        cons.extend([x.strip() for x in ca.split("\n") if x.strip()])
    write_lines(c, cons, sizes=[8] * len(cons), bolds=[True] + [False] * (len(cons) - 1))

    # ========== PORTS ==========
    t = doc.add_table(rows=1, cols=3)
    borders(t)
    ports = [
        ("PORT OF LOADING", _s(data.get("Port_Loading"))),
        ("PORT OF DISCHARGE", _s(data.get("Port_Discharge"))),
        ("COUNTRY OF ORIGIN", _s(data.get("Country_Origin"))),
    ]
    for cell, (lab, val) in zip(t.rows[0].cells, ports):
        set_cell_margins(cell, 30, 30, 40, 30)
        write_lines(cell, [lab, val], sizes=[8, 8], bolds=[True, False])

    # ========== BL | DATED ==========
    t = doc.add_table(rows=1, cols=2)
    borders(t)
    for cell, (lab, val) in zip(
        t.rows[0].cells,
        [("BL NO", _s(data.get("BL_No"))), ("DATED", _fmt_date(data.get("Dated_BL")))],
    ):
        set_cell_margins(cell, 30, 30, 40, 30)
        write_lines(cell, [lab, val], sizes=[8, 8], bolds=[True, False])

    # ========== VESSEL | FINAL DESTINATION ==========
    vessel = _s(data.get("Vessel_Name"))
    voy = _s(data.get("Voyage"))
    if voy and voy not in vessel:
        vessel = f"{vessel}, {voy}" if vessel else voy
    t = doc.add_table(rows=1, cols=2)
    borders(t)
    for cell, (lab, val) in zip(
        t.rows[0].cells,
        [("VESSEL NAME", vessel), ("FINAL DESTINATION", _s(data.get("Final_Destination")))],
    ):
        set_cell_margins(cell, 30, 30, 40, 30)
        write_lines(cell, [lab, val], sizes=[8, 8], bolds=[True, False])

    # ========== DESCRIPTION PRODUIT (bloc SI) ==========
    t = doc.add_table(rows=1, cols=1)
    borders(t)
    c = t.rows[0].cells[0]
    set_cell_margins(c, 40, 40, 50, 40)
    pl = _s(data.get("Product_Lines"))
    if pl:
        lines = [x.strip() for x in pl.replace("\r", "").split("\n") if x.strip()]
    else:
        n = n_ctr or 1
        lines = [
            f"{n} CONTAINER EACH CONTAINER WITH HEATING PAD",
            "RAW CRUDE SHEA BUTTER IN FLEXITANK OF REP OF BENIN",
            "ORIGIN WEST AFRICA MATERIAL DESCRIPTION : FA010C",
            "CRUDE SHB",
            f"HS CODE : {_s(data.get('Hs_Code') or '151590990')}",
        ]
    write_lines(c, lines, sizes=[8] * len(lines), bolds=[False] * len(lines))

    # ========== TABLEAU CONTENEURS (N lignes = N conteneurs booking) ==========
    # Colonnes modèle : S.No | CONTAINER NO | SEAL NO | TAN NO | ROTATION | BAGS | GROSS | NET
    ncols = 8
    headers = [
        "S. No.", "CONTAINER NO.", "SEAL NO", "TAN NO",
        "ROTATION", "BAGS", "GROSS WT [KGS]", "NET WT [KGS]",
    ]
    # +1 header + N data +1 total
    t = doc.add_table(rows=1 + max(n_ctr, 0) + 2, cols=ncols)
    borders(t)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER

    # largeurs relatives (somme ~ page utile)
    # S.No small, container, seal, tan, rotation, bags, gross, net
    widths_cm = [1.2, 3.2, 3.0, 2.2, 1.8, 1.4, 2.6, 2.6]
    for row in t.rows:
        for i, cell in enumerate(row.cells):
            cell.width = Cm(widths_cm[i])

    # header
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        set_cell_margins(cell, 25, 25, 20, 20)
        shade(cell, "D9E2F3")
        write_lines(cell, [h], sizes=[7], bolds=[True], center=True)
    set_row_height(t.rows[0], 0.5)

    total_bags = 0
    total_gross = 0.0
    total_net = 0.0

    for i, ctr in enumerate(ctrs):
        row = t.rows[i + 1]
        set_row_height(row, 0.48)
        bags = int(float(ctr.get("Bags") or 0))
        g = float(ctr.get("Gross_KG") or 0)
        n = float(ctr.get("Net_KG") or 0)
        total_bags += bags
        total_gross += g
        total_net += n
        vals = [
            str(ctr.get("S_No") or (i + 1)),
            _s(ctr.get("Container_No")),
            _s(ctr.get("Seal_No")),
            _s(ctr.get("Tan_No")),
            _s(ctr.get("Rotation")),
            str(bags) if bags else "",
            _fmt_w(g) if g else "",
            _fmt_w(n) if n else "",
        ]
        for j, v in enumerate(vals):
            set_cell_margins(row.cells[j], 20, 20, 15, 15)
            write_lines(row.cells[j], [v], sizes=[8], bolds=[False], center=True)

    # --- Ligne 1 totaux : texte récap (col 0-1) + chiffres alignés sous BAGS / GROSS / NET ---
    last = t.rows[n_ctr + 1]
    set_row_height(last, 1.15)
    for j in range(ncols):
        set_cell_margins(last.cells[j], 25, 25, 15, 15)

    summary = [
        f"Total No. of Containers : {n_ctr:02d}",
        f"Net Weight : {_fmt_w(total_net)} KGS",
        f"Gross Weight : {_fmt_w(total_gross)} KGS",
    ]
    # Libellés à gauche (col 0 vide, col 1 = texte gras)
    write_lines(last.cells[0], [""], sizes=[8], bolds=[True], center=True)
    write_lines(last.cells[1], summary, sizes=[8, 8, 8], bolds=[True, True, True])
    for j in range(2, 5):
        write_lines(last.cells[j], [""], sizes=[8], bolds=[True], center=True)
    # BAGS total sous colonne BAGS (index 5)
    write_lines(last.cells[5], [str(total_bags) if total_bags else ""], sizes=[9], bolds=[True], center=True)
    # GROSS sous colonne GROSS (index 6)
    write_lines(last.cells[6], [_fmt_w(total_gross)], sizes=[9], bolds=[True], center=True)
    # NET sous colonne NET (index 7)
    write_lines(last.cells[7], [_fmt_w(total_net)], sizes=[9], bolds=[True], center=True)

    # --- Ligne 2 TOTAL : même tableau, mêmes colonnes (collée, alignement vertical exact) ---
    tot = t.rows[n_ctr + 2]
    set_row_height(tot, 0.55)
    for j in range(ncols):
        set_cell_margins(tot.cells[j], 25, 25, 15, 15)
        shade(tot.cells[j], "F5F5F5")
    write_lines(tot.cells[0], ["TOTAL"], sizes=[9], bolds=[True], center=True)
    write_lines(tot.cells[1], [str(n_ctr)], sizes=[9], bolds=[True], center=True)
    for j in range(2, 5):
        write_lines(tot.cells[j], [""], sizes=[8], bolds=[True], center=True)
    write_lines(tot.cells[5], [str(total_bags) if total_bags else ""], sizes=[9], bolds=[True], center=True)
    write_lines(tot.cells[6], [f"{_fmt_w(total_gross)} KGS"], sizes=[9], bolds=[True], center=True)
    write_lines(tot.cells[7], [f"{_fmt_w(total_net)} KGS"], sizes=[9], bolds=[True], center=True)

    # Signature
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tight(p, 18, 2)
    set_run(p.add_run("FOR STE NAP SARL"), 10, True, NOIR)

    for _ in range(3):
        blank = doc.add_paragraph()
        tight(blank, 8, 8)
        set_run(blank.add_run(" "), 10, False, NOIR)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tight(p, 2, 2)
    set_run(p.add_run("AUTHORIZED SIGNATORY"), 9, True, NOIR)

    notes = _s(data.get("Notes"))
    if notes:
        p = doc.add_paragraph()
        tight(p, 8, 2)
        set_run(p.add_run("Notes: " + notes), 8, False, NOIR)

    num = _s(data.get("Packing_No") or data.get("BL_No") or "PL")
    safe = re.sub(r'[\\/*?:"<>|]', "", num).replace("/", "-")
    nom = f"Packing_List_{safe}.docx"
    chemin = os.path.join(DOSSIER, nom)
    doc.save(chemin)
    return chemin
