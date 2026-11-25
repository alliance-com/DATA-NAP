# -*- coding: utf-8 -*-
"""
Page Streamlit — Bons de livraison (STE NAP SARL)
-------------------------------------------------
Même principe que les bons de commande : formulaire -> bon Word professionnel -> registre numéroté.

Champs : Date de livraison · Quantité livrée · N° Facture Proforma correspondant · Référence (auto,
par produit) · N° camion  (+ client, chauffeur, lieu, observations).

Sécurité du bon :
  • valeurs sensibles imprimées en POLICE DE SÉCURITÉ (OCR A Extended — installée avec Windows)
  • quantité en toutes lettres entre astérisques (anti-surcharge)
  • CODE DE VÉRIFICATION unique (HMAC) + page « Vérifier un bon » dans l'application
  • micro-texte de bordure, QR code (si la bibliothèque `qrcode` est installée)

Fichiers créés à la racine du projet : bons_livraison.xlsx (registre) · bons_livraison_generes/ (Word)
Dépendances : streamlit, pandas, openpyxl, python-docx   (facultatif : pip install qrcode[pil])
À placer dans DATA/pages/
"""
import hashlib
import hmac
import io
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# ====================== PARAMÈTRES ======================
RACINE = Path(__file__).resolve().parent.parent
FICHIER_BL = RACINE / "bons_livraison.xlsx"
DOSSIER_BL = RACINE / "bons_livraison_generes"

SOCIETE = {
    "nom": "STE NAP SARL", "court": "NAP SARL",
    "adresse": "PLOT NO 14 GLODJIGBE ECONOMIC ZONE, REPUBLIC OF BENIN",
    "legal": "RCCM RB / COT / 21 B 30710 · IFU 3202113334014",
    "tel": "TEL: +229 94940414, 97974506",
}
PRODUITS = ["SHEA", "SOYABEAN", "COTTON"]
UNITES = {"kg": ("kilogramme", "kilogrammes"), "MT": ("tonne", "tonnes"), "sacs": ("sac", "sacs")}
STATUTS = ["Émis", "Livré / reçu", "Annulé"]

POLICE_SECURITE = "OCR A Extended"          # police de sécurité (installée avec Windows) — modifiable
CLE_SECURITE = "NAP-SARL-BL-CHANGER-CETTE-CLE"   # ⚠️ remplace par ta propre clé secrète (garde-la privée)

COLS_BL = ["REF", "Annee", "Numero", "Produit", "Date_Livraison", "Camion", "Chauffeur", "Client", "Lieu", "Quantite",
           "Unite", "Proforma", "Ref_BC", "Observations", "Code", "Statut", "Fichier", "Donnees"]
# Palette bleu ciel : bandeaux, titres, bordures et fonds du bon Word
# Palette sobre, économie d'encre : bleu ciel très clair, reste noir / gris / blanc
NAVY, ACCENT, LIGHT, GRAY = "2C3E50", "5D6D7E", "F4F6F7", "5D6D7E"
# Bleu ciel clair (bandeaux) — peu d'encre à l'impression
BLUE_ACCENT = "D6EAF8"          # fond bleu ciel clair
BLUE_ACCENT_TEXT = "1A5276"     # texte foncé sur fond clair



# ====================== OUTILS ======================
def fmt(n, dec=0) -> str:
    try:
        s = f"{float(n):,.{dec}f}".replace(",", " ")
        return s.rstrip("0").rstrip(".") if dec and "." in s else s
    except Exception:
        return ""


_U = ["zéro", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf", "dix", "onze", "douze",
      "treize", "quatorze", "quinze", "seize"]
_D = {2: "vingt", 3: "trente", 4: "quarante", 5: "cinquante", 6: "soixante"}


def _moins_100(n: int) -> str:
    if n < 17:
        return _U[n]
    if n < 20:
        return "dix-" + _U[n - 10]
    if n < 70:
        d, u = divmod(n, 10)
        return _D[d] if u == 0 else (_D[d] + " et un" if u == 1 else _D[d] + "-" + _U[u])
    if n < 80:
        r = n - 60
        return "soixante et onze" if r == 11 else "soixante-" + _moins_100(r)
    if n == 80:
        return "quatre-vingts"
    return "quatre-vingt-" + _moins_100(n - 80)


def _moins_1000(n: int, final: bool) -> str:
    c, r = divmod(n, 100)
    parts = []
    if c:
        parts.append("cent" if c == 1 else _U[c] + " cent" + ("s" if r == 0 and final else ""))
    if r:
        parts.append("quatre-vingt" if (r == 80 and not final) else _moins_100(r))
    return " ".join(parts)


def fr_mots(n: int) -> str:
    """25000 -> 'vingt-cinq mille'"""
    n = int(n)
    if n == 0:
        return "zéro"
    parts = []
    for val, sing, plur in ((10 ** 9, "milliard", "milliards"), (10 ** 6, "million", "millions")):
        q, n = divmod(n, val)
        if q:
            parts.append(_moins_1000(q, True) + " " + (sing if q == 1 else plur))
    q, n = divmod(n, 1000)
    if q:
        parts.append("mille" if q == 1 else _moins_1000(q, False) + " mille")
    if n:
        parts.append(_moins_1000(n, True))
    return " ".join(parts)


def quantite_en_lettres(q: float, unite: str) -> str:
    """28.5, 'MT' -> 'vingt-huit virgule cinq tonnes'"""
    sing, plur = UNITES.get(unite, (unite, unite))
    ent = int(q)
    dec = f"{q:.3f}".split(".")[1].rstrip("0")
    txt = fr_mots(ent)
    if dec:
        txt += " virgule " + " ".join(_U[int(c)] for c in dec)
    if unite == "MT" and not dec and (txt == "un" or txt.endswith(" et un")):
        txt = txt[:-2] + "une"                                   # « une tonne », « vingt et une tonnes »
    return f"{txt} {plur if q >= 2 else sing}"


def code_securite(d: dict) -> str:
    """Code de vérification unique (HMAC-SHA256) lié aux données clés du bon. Format XXXX-XXXX-XXXX."""
    msg = "|".join([str(d["ref"]), f"{d['date_liv']:%Y-%m-%d}", f"{float(d['quantite']):.3f}", str(d["unite"]),
                    str(d["proforma"]).strip().upper(), str(d["camion"]).strip().upper()])
    h = hmac.new(CLE_SECURITE.encode(), msg.encode(), hashlib.sha256).hexdigest()[:12].upper()
    return f"{h[:4]}-{h[4:8]}-{h[8:]}"


def _jdump(d: dict) -> str:
    import json
    return json.dumps(d, default=lambda o: o.isoformat() if isinstance(o, (date, datetime)) else str(o), ensure_ascii=False)


def _jload(s: str) -> dict:
    import json
    d = json.loads(s)
    for k in ("date_liv", "date_emission"):
        d[k] = pd.to_datetime(d.get(k), errors="coerce").date() if d.get(k) else date.today()
    d["societe_info"] = SOCIETE
    return d


def reference(numero: int, produit: str, annee: int) -> str:
    return f"{numero:04d}-BL-{produit}/NAP-{annee}"


# ====================== WORD ======================

def _filigrane_produit_word(doc, produit: str):
    """Filigrane centré : nom du produit (Arial Narrow, 25 pt), gris clair."""
    from docx.oxml import OxmlElement
    from lxml import etree

    texte = (produit or "").strip().upper()
    if not texte:
        return
    texte_xml = texte.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False
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
        'id="FiligraneProduit" type="#_x0000_t136" '
        'style="position:absolute;margin-left:0;margin-top:0;width:480pt;height:80pt;'
        'z-index:-251658240;mso-position-horizontal:center;'
        'mso-position-horizontal-relative:page;mso-position-vertical:center;'
        'mso-position-vertical-relative:page;" fillcolor="#D0D0D0" stroked="f">'
        '<v:fill opacity=".35"/>'
        '<v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr>'
        '<w:rFonts w:ascii="Arial Narrow" w:hAnsi="Arial Narrow"/>'
        '<w:b/><w:sz w:val="50"/><w:szCs w:val="50"/><w:color w:val="B0B0B0"/>'
        '</w:rPr><w:t xml:space="preserve">'
        + texte_xml +
        '</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>'
    )
    shape = etree.fromstring(shape_xml)
    pict.append(shape)


def construire_docx(d: dict) -> bytes:
    try:
        import docx  # noqa: F401
    except ImportError:
        raise RuntimeError("Installe python-docx :  pip install python-docx")
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.table import WD_ROW_HEIGHT_RULE
    from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt, RGBColor, Twips

    soc = d.get("societe_info") or SOCIETE
    SEC = POLICE_SECURITE
    doc = Document()
    stn = doc.styles["Normal"]
    stn.font.name = "Calibri"
    stn.font.size = Pt(10.5)
    stn.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    doc.core_properties.author = soc["nom"]
    doc.core_properties.title = f"Bon de livraison {d['ref']}"
    s = doc.sections[0]
    s.orientation = WD_ORIENT.PORTRAIT
    s.page_width, s.page_height = Mm(210), Mm(297)
    s.top_margin, s.bottom_margin = Twips(700), Twips(700)
    s.left_margin = s.right_margin = Twips(850)
    W = 11906 - 2 * 850

    def R(p, text, size=10.5, bold=False, italic=False, color=None, font="Calibri"):
        r = p.add_run(text)
        r.font.name = font
        r._element.rPr.rFonts.set(qn("w:eastAsia"), font)
        r.font.size = Pt(size)
        r.bold, r.italic = bold, italic
        if color:
            r.font.color.rgb = RGBColor.from_string(color)
        return r

    def P(container, runs, align=None, before=0, after=0, first=False, bottom=None, line=None):
        use = first and container.paragraphs and not container.paragraphs[0].runs
        p = container.paragraphs[0] if use else container.add_paragraph()
        if bottom:
            pPr = p._p.get_or_add_pPr()
            bd = OxmlElement("w:pBdr")
            e = OxmlElement("w:bottom")
            e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(bottom[0])); e.set(qn("w:space"), str(bottom[2])); e.set(qn("w:color"), bottom[1])
            bd.append(e)
            pPr.append(bd)
        pf = p.paragraph_format
        pf.space_before, pf.space_after = Pt(before), Pt(after)
        if line:
            pf.line_spacing = line
        if align:
            p.alignment = align
        for item in runs:
            t, kw = (item, {}) if isinstance(item, str) else item
            R(p, t, **kw)
        return p

    def C(cell, w=None, fill=None, borders=None, margins=(60, 60, 110, 110), valign="center"):
        if w:
            cell.width = Twips(w)
        tcPr = cell._tc.get_or_add_tcPr()
        tb = OxmlElement("w:tcBorders")
        for e in ("top", "left", "bottom", "right"):
            spec = (borders or {}).get(e)
            el = OxmlElement(f"w:{e}")
            if spec:
                el.set(qn("w:val"), "single"); el.set(qn("w:sz"), str(spec[0])); el.set(qn("w:space"), "0"); el.set(qn("w:color"), spec[1])
            else:
                el.set(qn("w:val"), "nil")
            tb.append(el)
        tcPr.append(tb)
        if fill:
            sh = OxmlElement("w:shd")
            sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), fill)
            tcPr.append(sh)
        mar = OxmlElement("w:tcMar")
        for k, v in zip(("top", "left", "bottom", "right"), (margins[0], margins[2], margins[1], margins[3])):
            m = OxmlElement(f"w:{k}")
            m.set(qn("w:w"), str(v)); m.set(qn("w:type"), "dxa")
            mar.append(m)
        tcPr.append(mar)
        va = OxmlElement("w:vAlign")
        va.set(qn("w:val"), valign)
        tcPr.append(va)

    def table(rows, widths):
        t = doc.add_table(rows=rows, cols=len(widths))
        t.autofit = False
        for i, w in enumerate(widths):
            t.columns[i].width = Twips(w)
        for r in t.rows:
            for i, c in enumerate(r.cells):
                c.width = Twips(widths[i])
        return t

    def box(color, sz=6):
        sp = (sz, color)
        return {"top": sp, "bottom": sp, "left": sp, "right": sp}

    def height(row, tw):
        row.height = Twips(tw)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    def micro():
        txt = (f"{soc['court']} · BON DE LIVRAISON · DOCUMENT AUTHENTIFIÉ · " * 9)
        P(doc, [(txt, dict(size=3, color="9AA7B5", font="Arial"))], AL.CENTER)

    # ---------- bande micro-texte haute ----------
    micro()
    # Filigrane produit (centré, Arial Narrow 25 pt)
    _filigrane_produit_word(doc, d.get("produit") or "")

    # ---------- en-tête : société + référence en police de sécurité ----------
    t = table(1, [6300, W - 6300])
    c0, c1 = t.rows[0].cells
    C(c0, 6300, margins=(60, 60, 0, 100))
    P(c0, [(soc["nom"], dict(font="Arial", bold=True, size=22, color=NAVY))], first=True, after=2)
    P(c0, [(soc["adresse"], dict(size=8, color="444444"))], after=0)
    P(c0, [(soc["legal"], dict(size=8, color="444444"))], after=0)
    P(c0, [(soc["tel"], dict(size=8, color="444444"))])
    C(c1, W - 6300, borders=box("CCCCCC", 8), fill=LIGHT, margins=(90, 90, 140, 140))
    P(c1, [("RÉFÉRENCE DU BON", dict(size=7.5, bold=True, color=GRAY))], AL.CENTER, first=True, after=3)
    P(c1, [(d["ref"], dict(font=SEC, size=12, bold=True, color=NAVY))], AL.CENTER)
    P(doc, [], after=2, bottom=(12, "CCCCCC", 1))

    # ---------- bandeau titre ----------
    t = table(1, [7000, W - 7000])
    a, b = t.rows[0].cells
    C(a, 7000, fill=BLUE_ACCENT, margins=(120, 120, 240, 120))
    P(a, [("BON DE LIVRAISON", dict(font="Arial", bold=True, size=20, color=BLUE_ACCENT_TEXT))], first=True)
    C(b, W - 7000, fill=BLUE_ACCENT, margins=(120, 120, 120, 240))
    P(b, [("ÉMIS LE ", dict(size=8, color=GRAY)), (f"{d['date_emission']:%d/%m/%Y}", dict(font=SEC, size=10, bold=True, color=BLUE_ACCENT_TEXT))],
      AL.RIGHT, first=True)
    P(doc, [], after=6)

    # ---------- informations de livraison ----------
    cw = [2250, 2853, 2250, 2853]
    lines = [
        (("DATE DE LIVRAISON", f"{d['date_liv']:%d/%m/%Y}", True), ("N° CAMION", d["camion"], True)),
        (("N° FACTURE PROFORMA", d["proforma"], True), ("CHAUFFEUR", d.get("chauffeur") or "—", False)),
        (("CLIENT / DESTINATAIRE", d.get("client") or "—", False), ("PRODUIT", d["produit"], False)),
        (("LIEU DE LIVRAISON", d.get("lieu") or "—", False), ("RÉF. BON DE COMMANDE", d.get("ref_bc") or "—", bool(d.get("ref_bc")))),
    ]
    t = table(len(lines), cw)
    for i, (l, r) in enumerate(lines):
        height(t.rows[i], 560)
        for k, (lab, val, sec) in enumerate((l, r)):
            cl, cv = t.rows[i].cells[2 * k], t.rows[i].cells[2 * k + 1]
            C(cl, cw[2 * k], fill=LIGHT, borders=box("CCCCCC", 4))
            P(cl, [(lab, dict(size=7.5, bold=True, color=NAVY))], first=True)
            C(cv, cw[2 * k + 1], borders=box("CCCCCC", 4))
            P(cv, [(str(val), dict(font=SEC, size=10.5, bold=True) if sec else dict(size=11, bold=True))], first=True)
    P(doc, [], after=6)

    # ---------- quantité livrée ----------
    q = float(d["quantite"])
    q_txt = f"{fmt(q, 3).replace('.', ',')} {d['unite']}"
    t = table(1, [4000, W - 4000])
    a, b = t.rows[0].cells
    height(t.rows[0], 1100)
    C(a, 4000, fill=BLUE_ACCENT, margins=(100, 100, 200, 120))
    P(a, [("QUANTITÉ LIVRÉE", dict(size=8, bold=True, color=BLUE_ACCENT_TEXT))], first=True, after=3)
    P(a, [(q_txt, dict(font=SEC, size=19, bold=True, color=BLUE_ACCENT_TEXT))])
    C(b, W - 4000, borders=box("CCCCCC", 8), margins=(100, 100, 200, 200))
    P(b, [("ARRÊTÉ LA QUANTITÉ LIVRÉE À", dict(size=7.5, bold=True, color=GRAY))], first=True, after=3)
    P(b, [(f"*** {quantite_en_lettres(q, d['unite']).upper()} ***", dict(font=SEC, size=9.5, bold=True, color=NAVY))])
    P(doc, [], after=6)

    # ---------- observations ----------
    t = table(1, [W])
    height(t.rows[0], 1000)
    C(t.rows[0].cells[0], W, borders=box("CCCCCC", 4), valign="top", margins=(80, 80, 140, 140))
    P(t.rows[0].cells[0], [("OBSERVATIONS", dict(size=7.5, bold=True, color=NAVY))], first=True, after=3)
    P(t.rows[0].cells[0], [(d.get("observations") or " ", dict(size=10))])
    P(doc, [], after=8)

    # ---------- signatures ----------
    sw = [W // 3, W // 3, W - 2 * (W // 3)]
    t = table(2, sw)
    for i, lab in enumerate(("LE LIVREUR / TRANSPORTEUR", "LE RÉCEPTIONNAIRE", f"POUR {soc['nom']}")):
        C(t.rows[0].cells[i], sw[i], fill=LIGHT, borders=box("CCCCCC", 4))
        P(t.rows[0].cells[i], [(lab, dict(size=8, bold=True, color=NAVY))], AL.CENTER, first=True)
        height(t.rows[1], 1500)
        C(t.rows[1].cells[i], sw[i], borders=box("CCCCCC", 4), valign="bottom")
        P(t.rows[1].cells[i], [("Nom, signature et cachet", dict(size=7.5, italic=True, color=GRAY))], AL.CENTER, first=True)
    P(doc, [], after=8)

    # ---------- bloc sécurité ----------
    qr = None
    try:
        import qrcode                                            # facultatif
        payload = f"NAP-BL|{d['ref']}|{d['date_liv']:%Y-%m-%d}|{fmt(q, 3)}{d['unite']}|{d['camion']}|{d['proforma']}|{d['code']}"
        img = qrcode.make(payload)
        qr = io.BytesIO()
        img.save(qr, format="PNG")
        qr.seek(0)
    except Exception:
        qr = None
    # Code à gauche, QR à droite (dans le même tableau)
    if qr:
        sec_w = [W - 1700, 1700]
    else:
        sec_w = [W]
    t = table(1, sec_w)
    height(t.rows[0], 1250)
    C(t.rows[0].cells[0], sec_w[0], borders=box("CCCCCC", 4), fill="FAFBFD", margins=(80, 80, 180, 140))
    P(t.rows[0].cells[0], [("CODE DE VÉRIFICATION", dict(size=7.5, bold=True, color=GRAY))], first=True, after=2)
    P(t.rows[0].cells[0], [(d["code"], dict(font=SEC, size=14, bold=True, color=NAVY))], after=3)
    P(t.rows[0].cells[0], [("Code unique lié à ce bon de livraison et vérifiable. "
                            "Toute modification du bon est proscrite.",
                            dict(size=7.5, color="444444"))])
    if qr:
        C(t.rows[0].cells[1], sec_w[1], borders=box("CCCCCC", 4), fill="FAFBFD", margins=(40, 40, 60, 60))
        t.rows[0].cells[1].paragraphs[0].alignment = AL.CENTER
        t.rows[0].cells[1].paragraphs[0].add_run().add_picture(qr, width=Mm(26))

    # ---------- VRAI PIED DE PAGE (zone footer, collé en bas de page) ----------
    section = doc.sections[0]
    section.footer_distance = Twips(400)
    footer = section.footer
    footer.is_linked_to_previous = False
    # Vider le footer existant
    for p in list(footer.paragraphs):
        p.clear()
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    # Ligne de séparation haute du pied de page
    pPr = fp._p.get_or_add_pPr()
    bd = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "6")
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), "CCCCCC")
    bd.append(top)
    pPr.append(bd)
    fp.alignment = AL.CENTER
    fp.paragraph_format.space_before = Pt(2)
    fp.paragraph_format.space_after = Pt(0)
    R(fp, f"{soc['nom']}  ·  {soc['adresse']}", size=7.5, color=GRAY)
    fp2 = footer.add_paragraph()
    fp2.alignment = AL.CENTER
    fp2.paragraph_format.space_before = Pt(0)
    fp2.paragraph_format.space_after = Pt(0)
    R(fp2, f"{soc['legal']}  ·  {soc['tel']}", size=7, color=GRAY)
    fp3 = footer.add_paragraph()
    fp3.alignment = AL.CENTER
    fp3.paragraph_format.space_before = Pt(0)
    R(fp3, "Document authentifié — modification proscrite", size=6.5, italic=True, color="999999")

    for z in doc.settings.element.findall(qn("w:zoom")):
        z.set(qn("w:percent"), "100")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()



def construire_pdf(d: dict) -> bytes:
    """Génère le bon de livraison en PDF A4 (reportlab)."""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import mm
        from reportlab.lib import colors
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, KeepTogether,
        )
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    except ImportError as e:
        raise RuntimeError("Installez reportlab :  pip install reportlab") from e

    soc = d.get("societe_info") or SOCIETE
    q = float(d["quantite"])
    q_txt = f"{fmt(q, 3).replace('.', ',')} {d['unite']}"
    lettres = quantite_en_lettres(q, d["unite"]).upper()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=12 * mm, bottomMargin=22 * mm,
    )
    styles = getSampleStyleSheet()
    # Sobre : ~25 % couleur (bandeau titre seulement), reste noir / gris / blanc
    navy = colors.HexColor("#2C3E50")       # texte titres
    accent_text = colors.HexColor("#1A5276")  # texte sur fond bleu clair
    accent = colors.HexColor("#D6EAF8")     # bleu ciel clair (économie encre)
    light = colors.HexColor("#F4F6F7")      # fond gris très clair
    gray = colors.HexColor("#5D6D7E")
    border = colors.HexColor("#CCCCCC")

    s_title = ParagraphStyle("t", parent=styles["Normal"], fontName="Helvetica-Bold",
                             fontSize=16, textColor=colors.white, leading=20)
    s_h = ParagraphStyle("h", parent=styles["Normal"], fontName="Helvetica-Bold",
                         fontSize=14, textColor=navy, leading=18)
    s_lab = ParagraphStyle("lab", parent=styles["Normal"], fontName="Helvetica-Bold",
                           fontSize=7.5, textColor=navy, leading=9)
    s_val = ParagraphStyle("val", parent=styles["Normal"], fontName="Courier-Bold",
                           fontSize=10, textColor=colors.black, leading=12)
    s_val2 = ParagraphStyle("val2", parent=styles["Normal"], fontName="Helvetica-Bold",
                            fontSize=10, textColor=colors.black, leading=12)
    s_small = ParagraphStyle("sm", parent=styles["Normal"], fontName="Helvetica",
                             fontSize=8, textColor=colors.HexColor("#444444"), leading=10)
    s_center = ParagraphStyle("c", parent=styles["Normal"], fontName="Helvetica",
                              fontSize=7.5, textColor=gray, alignment=TA_CENTER, leading=9)
    s_code = ParagraphStyle("code", parent=styles["Normal"], fontName="Courier-Bold",
                            fontSize=12, textColor=navy, leading=14)

    story = []
    # En-tête société + ref
    left = [
        Paragraph(soc["nom"], s_h),
        Paragraph(soc["adresse"], s_small),
        Paragraph(soc["legal"], s_small),
        Paragraph(soc["tel"], s_small),
    ]
    right_data = [
        [Paragraph("RÉFÉRENCE DU BON", ParagraphStyle("r1", parent=s_lab, alignment=TA_CENTER))],
        [Paragraph(str(d["ref"]), ParagraphStyle("r2", parent=s_val, alignment=TA_CENTER, fontSize=11, textColor=navy))],
    ]
    right = Table(right_data, colWidths=[70 * mm])
    right.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), light),
        ("BOX", (0, 0), (-1, -1), 0.8, border),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    head = Table([[left, right]], colWidths=[105 * mm, 75 * mm])
    head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(head)
    story.append(Spacer(1, 3 * mm))
    story.append(HRFlowable(width="100%", thickness=0.8, color=border, spaceBefore=1, spaceAfter=3))

    # Bandeau titre
    s_title_light = ParagraphStyle(
        "t_light", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=16, textColor=accent_text, leading=20,
    )
    bandeau = Table(
        [[Paragraph("BON DE LIVRAISON", s_title_light),
          Paragraph(f"ÉMIS LE  <font face='Courier-Bold' color='#1A5276'>{d['date_emission']:%d/%m/%Y}</font>",
                    ParagraphStyle("em", parent=styles["Normal"], fontSize=9,
                                   textColor=accent_text, alignment=TA_RIGHT))]],
        colWidths=[110 * mm, 70 * mm],
    )
    bandeau.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), accent),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(bandeau)
    story.append(Spacer(1, 4 * mm))

    def cell_lab(txt):
        return Paragraph(txt, s_lab)

    def cell_val(txt, sec=False):
        return Paragraph(str(txt), s_val if sec else s_val2)

    rows_info = [
        [cell_lab("DATE DE LIVRAISON"), cell_val(f"{d['date_liv']:%d/%m/%Y}", True),
         cell_lab("N° CAMION"), cell_val(d.get("camion") or "—", True)],
        [cell_lab("N° FACTURE PROFORMA"), cell_val(d.get("proforma") or "—", True),
         cell_lab("CHAUFFEUR"), cell_val(d.get("chauffeur") or "—")],
        [cell_lab("CLIENT / DESTINATAIRE"), cell_val(d.get("client") or "—"),
         cell_lab("PRODUIT"), cell_val(d.get("produit") or "—")],
        [cell_lab("LIEU DE LIVRAISON"), cell_val(d.get("lieu") or "—"),
         cell_lab("RÉF. BON DE COMMANDE"), cell_val(d.get("ref_bc") or "—", bool(d.get("ref_bc")))],
    ]
    t_info = Table(rows_info, colWidths=[38 * mm, 52 * mm, 38 * mm, 52 * mm])
    t_info.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), light),
        ("BACKGROUND", (2, 0), (2, -1), light),
        ("BOX", (0, 0), (-1, -1), 0.6, border),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, border),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(t_info)
    story.append(Spacer(1, 4 * mm))

    # Quantité
    t_q = Table([
        [Paragraph("QUANTITÉ LIVRÉE", ParagraphStyle("ql", parent=s_lab, textColor=accent_text)),
         Paragraph("ARRÊTÉ LA QUANTITÉ LIVRÉE À", s_lab)],
        [Paragraph(q_txt, ParagraphStyle("qv", parent=styles["Normal"], fontName="Courier-Bold",
                                          fontSize=16, textColor=accent_text, leading=20)),
         Paragraph(f"*** {lettres} ***",
                   ParagraphStyle("ql2", parent=styles["Normal"], fontName="Courier-Bold",
                                  fontSize=9, textColor=navy, leading=12))],
    ], colWidths=[70 * mm, 110 * mm])
    t_q.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), accent),  # bleu ciel comme bandeau titre
        ("BOX", (1, 0), (1, -1), 1, border),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(t_q)
    story.append(Spacer(1, 4 * mm))

    # Observations
    t_obs = Table([
        [Paragraph("OBSERVATIONS", s_lab)],
        [Paragraph(d.get("observations") or "—", s_val2)],
    ], colWidths=[180 * mm])
    t_obs.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.6, border),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("BACKGROUND", (0, 0), (0, 0), light),
    ]))
    story.append(t_obs)
    story.append(Spacer(1, 6 * mm))

    # Signatures
    sig_labs = ["LE LIVREUR / TRANSPORTEUR", "LE RÉCEPTIONNAIRE", f"POUR {soc['nom']}"]
    sig_row1 = [Paragraph(x, ParagraphStyle("sl", parent=s_lab, alignment=TA_CENTER)) for x in sig_labs]
    sig_row2 = [Paragraph("Nom, signature et cachet",
                          ParagraphStyle("ss", parent=s_center, fontSize=7, fontName="Helvetica-Oblique"))
                for _ in sig_labs]
    t_sig = Table([sig_row1, sig_row2], colWidths=[60 * mm, 60 * mm, 60 * mm], rowHeights=[8 * mm, 22 * mm])
    t_sig.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), light),
        ("BOX", (0, 0), (-1, -1), 0.6, border),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, border),
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
        ("VALIGN", (0, 1), (-1, 1), "BOTTOM"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t_sig)
    story.append(Spacer(1, 6 * mm))

    # Code vérification (gauche) + QR (droite) — cellules simples (évite overflow KeepTogether)
    code = d.get("code") or ""
    left_html = (
        f'<font color="#5D6D7E"><b>CODE DE VÉRIFICATION</b></font><br/>'
        f'<font face="Courier-Bold" size="12" color="#2C3E50"><b>{code}</b></font><br/>'
        f'<font size="7" color="#444444">Code unique lié à ce bon de livraison et vérifiable '
        f"dans l'application. Toute modification du document est proscrite.</font>"
    )
    left_para = Paragraph(left_html, ParagraphStyle(
        "sec_left", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=11,
    ))
    qr_flowable = None
    try:
        import qrcode
        from reportlab.platypus import Image as RLImage
        payload = (
            f"NAP-BL|{d['ref']}|{d['date_liv']:%Y-%m-%d}|"
            f"{fmt(q, 3)}{d['unite']}|{d.get('camion')}|{d.get('proforma')}|{code}"
        )
        img = qrcode.make(payload)
        qr_buf = io.BytesIO()
        img.save(qr_buf, format="PNG")
        qr_buf.seek(0)
        qr_flowable = RLImage(qr_buf, width=26 * mm, height=26 * mm)
    except Exception:
        qr_flowable = None

    if qr_flowable is not None:
        t_sec = Table([[left_para, qr_flowable]], colWidths=[148 * mm, 32 * mm])
        t_sec.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.8, border),
            ("BACKGROUND", (0, 0), (-1, -1), colors.white),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (1, 0), (1, 0), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (0, 0), 8),
            ("RIGHTPADDING", (1, 0), (1, 0), 4),
        ]))
    else:
        t_sec = Table([[left_para]], colWidths=[180 * mm])
        t_sec.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.8, border),
            ("BACKGROUND", (0, 0), (-1, -1), colors.white),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ]))
    story.append(t_sec)

    def _pied_page(canvas, doc_):
        canvas.saveState()
        page_w, page_h = A4
        # Filigrane produit centré (Arial Narrow 25 si dispo, sinon Helvetica-Bold 25)
        prod = (d.get("produit") or "").strip().upper()
        if prod:
            canvas.saveState()
            canvas.setFillColorRGB(0.69, 0.69, 0.69)
            font_name = "Helvetica-Bold"
            try:
                import os
                from reportlab.pdfbase import pdfmetrics
                from reportlab.pdfbase.ttfonts import TTFont
                for path in (
                    r"C:\Windows\Fonts\ARIALN.TTF",
                    r"C:\Windows\Fonts\arialn.ttf",
                    r"C:\Windows\Fonts\ARIALNB.TTF",
                ):
                    if os.path.exists(path):
                        if "ArialNarrow" not in pdfmetrics.getRegisteredFontNames():
                            pdfmetrics.registerFont(TTFont("ArialNarrow", path))
                        font_name = "ArialNarrow"
                        break
            except Exception:
                pass
            canvas.setFont(font_name, 25)
            canvas.drawCentredString(page_w / 2.0, page_h / 2.0, prod)
            canvas.restoreState()
        # Pied de page
        y = 12 * mm
        canvas.setStrokeColor(colors.HexColor("#CCCCCC"))
        canvas.setLineWidth(0.6)
        canvas.line(14 * mm, y + 10 * mm, page_w - 14 * mm, y + 10 * mm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#5D6D7E"))
        canvas.drawCentredString(page_w / 2, y + 5.5 * mm, f"{soc['nom']}  ·  {soc['adresse']}")
        canvas.setFont("Helvetica", 6.5)
        canvas.drawCentredString(page_w / 2, y + 2.5 * mm, f"{soc['legal']}  ·  {soc['tel']}")
        canvas.setFont("Helvetica-Oblique", 6)
        canvas.setFillColor(colors.HexColor("#999999"))
        canvas.drawCentredString(page_w / 2, y - 0.5 * mm, "Document authentifié — modification proscrite")
        canvas.restoreState()

    doc.build(story, onFirstPage=_pied_page, onLaterPages=_pied_page)
    buf.seek(0)
    return buf.getvalue()


def nom_fichier_pdf(ref: str) -> str:
    return "BL_" + re.sub(r"[^A-Za-z0-9]+", "_", ref).strip("_") + ".pdf"



# ====================== REGISTRE (Excel) ======================
def _sauver(wb, chemin):
    try:
        wb.save(chemin)
    except PermissionError:
        raise PermissionError(f"Impossible d'écrire « {Path(chemin).name} » : ferme-le dans Excel puis réessaie.")


def _creer_registre(chemin: Path):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Bons"
    ws.append(COLS_BL)
    ws.freeze_panes = "A2"
    for col, w in zip("ABCDEFGHIJKLMNOPQR", (28, 7, 8, 12, 13, 14, 18, 24, 20, 12, 7, 18, 20, 30, 16, 12, 34, 20)):
        ws.column_dimensions[col].width = w
    _sauver(wb, chemin)


def lire_registre(chemin=FICHIER_BL) -> pd.DataFrame:
    chemin = Path(chemin)
    if not chemin.exists():
        return pd.DataFrame(columns=COLS_BL)
    conv = {i: (lambda v: "" if pd.isna(v) else str(v).strip()) for i in (0, 5, 11, 12)}
    df = pd.read_excel(chemin, sheet_name=0, converters=conv)
    for c in COLS_BL:
        if c not in df.columns:
            df[c] = None
    df = df[COLS_BL].dropna(how="all", subset=["REF"]).copy()
    df["Date_Livraison"] = pd.to_datetime(df["Date_Livraison"], errors="coerce")
    for c in ("Annee", "Numero"):
        s = pd.to_numeric(df[c], errors="coerce").fillna(0)
        df[c] = s.astype(int)
    df["Quantite"] = pd.to_numeric(df["Quantite"], errors="coerce").fillna(0.0)
    for c in ("Produit", "Chauffeur", "Client", "Lieu", "Unite", "Observations", "Code", "Statut", "Fichier", "Donnees"):
        df[c] = df[c].fillna("").astype(str)
    return df.reset_index(drop=True)


def prochain_numero(produit: str, annee: int, df: pd.DataFrame) -> int:
    """Prochain N° pour ce produit / année. Ne plante jamais si le filtre est vide (NaN)."""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return 1
    if "Numero" not in df.columns:
        return 1
    m = pd.Series(True, index=df.index)
    if produit and "Produit" in df.columns:
        m &= df["Produit"].astype(str).str.upper().str.strip() == str(produit).upper().strip()
    if annee is not None and "Annee" in df.columns:
        m &= pd.to_numeric(df["Annee"], errors="coerce") == int(annee)
    nums = pd.to_numeric(df.loc[m, "Numero"], errors="coerce").dropna()
    if nums.empty:
        return 1
    try:
        return int(nums.max()) + 1
    except (ValueError, TypeError):
        return 1


def ajouter_bon(rec: dict, chemin=FICHIER_BL):
    chemin = Path(chemin)
    if not chemin.exists():
        _creer_registre(chemin)
    from openpyxl import load_workbook
    wb = load_workbook(chemin)
    ws = wb.worksheets[0]
    r = 2
    while r <= ws.max_row and ws.cell(r, 1).value not in (None, ""):
        r += 1
    for c, k in enumerate(COLS_BL, start=1):
        cell = ws.cell(r, c, rec.get(k))
        if k == "Date_Livraison":
            cell.number_format = "dd/mm/yyyy"
        elif k in ("REF", "Camion", "Proforma", "Ref_BC"):
            cell.number_format = "@"
        elif k == "Quantite":
            cell.number_format = "#,##0.000"
    _sauver(wb, chemin)


def maj_statut(ref: str, statut: str, chemin=FICHIER_BL) -> bool:
    from openpyxl import load_workbook
    wb = load_workbook(chemin)
    ws = wb.worksheets[0]
    col = COLS_BL.index("Statut") + 1
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip() == ref:
            ws.cell(r, col, statut)
            _sauver(wb, chemin)
            return True
    return False


def nom_fichier(ref: str) -> str:
    return "BL_" + re.sub(r"[^A-Za-z0-9]+", "_", ref).strip("_") + ".docx"


def verifier_bon(ref: str, code_saisi: str, df: pd.DataFrame):
    """Retourne (ok, message, ligne|None). Recalcule le code depuis les données du registre."""
    ligne = df[df["REF"] == ref.strip()]
    if ligne.empty:
        return False, "Référence inconnue du registre.", None
    lg = ligne.iloc[0]
    d = {"ref": lg["REF"], "date_liv": lg["Date_Livraison"].date(), "quantite": lg["Quantite"], "unite": lg["Unite"],
         "proforma": lg["Proforma"], "camion": lg["Camion"]}
    attendu = code_securite(d)
    saisi = re.sub(r"[^A-Za-z0-9]", "", code_saisi).upper()
    if saisi == attendu.replace("-", ""):
        return True, "✅ Bon AUTHENTIQUE : le code correspond aux données du registre.", lg
    return False, "❌ Code NON CONFORME : bon modifié, contrefait ou code mal saisi.", lg


# ====================== INTERFACE ======================
if __name__ != "__main__" or True:   # (toujours exécuté par Streamlit)
    import streamlit as st

    st.set_page_config(page_title="Bons de livraison", page_icon="🚚", layout="wide")
    ss = st.session_state
    st.title("🚚 Bons de livraison")
    st.caption("Bon professionnel numéroté · police de sécurité · code de vérification · registre")
    flash = ss.pop("BL_flash", None)
    if flash:
        getattr(st, flash[0])(flash[1])

    reg = lire_registre()
    tab_new, tab_reg, tab_ver = st.tabs(["➕ Nouveau bon", f"📚 Registre ({len(reg)})", "🔐 Vérifier un bon"])

    # =================== NOUVEAU BON ===================
    with tab_new:
        a1, a2, a3 = st.columns(3)
        prod_sel = a1.selectbox("Produit", PRODUITS + ["AUTRE…"], key="BL_prod")
        produit = a1.text_input("Nom du produit", value="", key="BL_prod_autre").strip().upper() if prod_sel == "AUTRE…" else prod_sel
        d_liv = a2.date_input("Date de livraison", value=date.today(), format="DD/MM/YYYY", key="BL_date")
        n_auto = prochain_numero(produit, d_liv.year, reg) if produit else 1
        ref = a3.text_input("Référence (automatique, modifiable)", value=reference(n_auto, produit or "XXX", d_liv.year),
                            key=f"BL_ref_{produit}_{d_liv.year}_{n_auto}")

        b1, b2, b3 = st.columns(3)
        proforma = b1.text_input("N° Facture Proforma correspondant", placeholder="ex. NAP-5-2026", key="BL_pf")
        camion = b2.text_input("N° camion", placeholder="ex. BN 7046 RB", key="BL_cam")
        chauffeur = b3.text_input("Chauffeur (facultatif)", key="BL_chf")

        c1, c2, c3 = st.columns(3)
        client = c1.text_input("Client / destinataire", key="BL_cli")
        lieu = c2.text_input("Lieu de livraison", key="BL_lieu")
        ref_bc = c3.text_input("Réf. bon de commande (facultatif)", key="BL_bc")

        q1, q2 = st.columns([2, 1])
        quantite = q1.number_input("Quantité livrée", min_value=0.0, step=0.5, format="%.3f", key="BL_qte")
        unite = q2.selectbox("Unité", list(UNITES), key="BL_unite")
        obs = st.text_area("Observations (facultatif)", height=80, key="BL_obs")

        donnees = {"ref": ref.strip(), "produit": produit, "date_liv": d_liv, "date_emission": date.today(),
                   "camion": camion.strip(), "chauffeur": chauffeur.strip(), "client": client.strip(), "lieu": lieu.strip(),
                   "quantite": float(quantite), "unite": unite, "proforma": proforma.strip(), "ref_bc": ref_bc.strip(),
                   "observations": obs.strip(), "societe_info": SOCIETE}
        if quantite > 0:
            st.caption(f"En lettres : **{quantite_en_lettres(float(quantite), unite)}**")
        if donnees["ref"] and donnees["camion"] and donnees["proforma"] and quantite > 0:
            donnees["code"] = code_securite(donnees)
            st.caption(f"Code de vérification : `{donnees['code']}`")

        if st.button("📄 Générer le bon de livraison (Word)", type="primary"):
            err = []
            if not produit:
                err.append("Indique le produit.")
            if not donnees["ref"]:
                err.append("La référence est vide.")
            elif donnees["ref"] in set(reg["REF"]):
                err.append(f"La référence {donnees['ref']} existe déjà.")
            if not proforma.strip():
                err.append("Indique le N° de facture proforma correspondant.")
            if not camion.strip():
                err.append("Indique le N° du camion.")
            if quantite <= 0:
                err.append("La quantité livrée doit être supérieure à 0.")
            if err:
                for e in err:
                    st.error(e)
            else:
                try:
                    contenu = construire_docx(donnees)
                    DOSSIER_BL.mkdir(exist_ok=True)
                    fichier = nom_fichier(donnees["ref"])
                    (DOSSIER_BL / fichier).write_bytes(contenu)
                    m = re.match(r"(\d+)-BL-.*?-(\d{4})$", donnees["ref"])
                    ajouter_bon({"REF": donnees["ref"], "Annee": int(m.group(2)) if m else d_liv.year,
                                 "Numero": int(m.group(1)) if m else 0, "Produit": produit, "Date_Livraison": d_liv,
                                 "Camion": donnees["camion"], "Chauffeur": donnees["chauffeur"], "Client": donnees["client"],
                                 "Lieu": donnees["lieu"], "Quantite": donnees["quantite"], "Unite": unite,
                                 "Proforma": donnees["proforma"], "Ref_BC": donnees["ref_bc"], "Observations": donnees["observations"],
                                 "Code": donnees["code"], "Statut": "Émis", "Fichier": fichier, "Donnees": _jdump(donnees)})
                except (RuntimeError, PermissionError) as e:
                    st.error(str(e))
                else:
                    ss["BL_dernier"] = (fichier, contenu, donnees["ref"])
                    ss["BL_dernier_data"] = dict(donnees)
                    ss["BL_flash"] = ("success", f"Bon {donnees['ref']} créé (code {donnees['code']}) et enregistré au registre.")
                    st.rerun()
        if ss.get("BL_dernier"):
            f_, c_, r_ = ss["BL_dernier"]
            d_pdf = ss.get("BL_dernier_data")
            col_w, col_p = st.columns(2)
            with col_w:
                st.download_button(
                    f"⬇️ Télécharger le bon {r_} (Word)",
                    data=c_,
                    file_name=f_,
                    type="primary",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_bl_word",
                )
            with col_p:
                try:
                    pdf_bytes = construire_pdf(d_pdf) if d_pdf else None
                    if pdf_bytes:
                        st.download_button(
                            f"⬇️ Télécharger le bon {r_} (PDF)",
                            data=pdf_bytes,
                            file_name=nom_fichier_pdf(r_),
                            type="primary",
                            mime="application/pdf",
                            key="dl_bl_pdf",
                        )
                except Exception as e:
                    st.caption(f"PDF indisponible : {e}")

    # =================== REGISTRE ===================
    with tab_reg:
        if reg.empty:
            st.info("Aucun bon de livraison enregistré pour l'instant.")
        else:
            k = st.columns(4)
            k[0].metric("Bons émis", int((reg["Statut"] != "Annulé").sum()))
            k[1].metric("Camions distincts", reg["Camion"].nunique())
            k[2].metric("Factures proforma", reg["Proforma"].nunique())
            k[3].metric("Annulés", int((reg["Statut"] == "Annulé").sum()))
            f1, f2, f3, f4 = st.columns(4)
            an = f1.selectbox("Année", ["Toutes"] + [int(a) for a in sorted(reg["Annee"].unique(), reverse=True)])
            pr = f2.selectbox("Produit", ["Tous"] + sorted(reg["Produit"].unique()))
            sts = f3.selectbox("Statut", ["Tous"] + STATUTS)
            q = f4.text_input("Recherche (réf, camion, proforma, client…)")
            vue = reg.copy()
            if an != "Toutes":
                vue = vue[vue["Annee"] == an]
            if pr != "Tous":
                vue = vue[vue["Produit"] == pr]
            if sts != "Tous":
                vue = vue[vue["Statut"] == sts]
            if q.strip():
                txt = vue[["REF", "Camion", "Proforma", "Client", "Chauffeur", "Ref_BC"]].astype(str).agg(" ".join, axis=1).str.lower()
                vue = vue[txt.str.contains(q.strip().lower(), regex=False, na=False)]
            aff = vue.sort_values(["Annee", "Produit", "Numero"], ascending=False)[
                ["REF", "Date_Livraison", "Produit", "Camion", "Proforma", "Client", "Quantite", "Unite", "Statut", "Code"]].copy()
            aff["Date_Livraison"] = aff["Date_Livraison"].dt.strftime("%d/%m/%Y")
            st.dataframe(aff, hide_index=True, width="stretch", height=300,
                         column_config={"Quantite": st.column_config.NumberColumn("Quantité", format="%.3f")})

            with st.expander("📊 Cumul livré par facture proforma"):
                actifs = reg[reg["Statut"] != "Annulé"]
                cum = (actifs.groupby(["Proforma", "Produit", "Unite"], as_index=False)
                       .agg(Camions=("Camion", "nunique"), Bons=("REF", "count"), Quantité_livrée=("Quantite", "sum")))
                st.dataframe(cum.sort_values("Proforma"), hide_index=True, width="stretch")

            st.subheader("🔧 Suivre un bon")
            refs = list(vue.sort_values(["Annee", "Produit", "Numero"], ascending=False)["REF"])
            if refs:
                sel = st.selectbox("Bon", refs, key="BL_sel")
                lg = reg[reg["REF"] == sel].iloc[0]
                s1, s2, s3 = st.columns(3)
                nouv = s1.selectbox("Statut", STATUTS, index=STATUTS.index(lg["Statut"]) if lg["Statut"] in STATUTS else 0,
                                    key=f"BL_st_{sel}")
                if s1.button("💾 Enregistrer le statut"):
                    try:
                        maj_statut(sel, nouv)
                    except PermissionError as e:
                        st.error(str(e))
                    else:
                        ss["BL_flash"] = ("success", f"Statut du bon {sel} mis à jour.")
                        st.rerun()
                contenu = None
                try:
                    p = DOSSIER_BL / lg["Fichier"] if lg["Fichier"] else None
                    contenu = p.read_bytes() if p is not None and p.exists() else (construire_docx(_jload(lg["Donnees"])) if lg["Donnees"] else None)
                except RuntimeError as e:
                    st.error(str(e))
                if contenu:
                    s2.download_button(
                        "⬇️ Télécharger le Word",
                        data=contenu,
                        file_name=lg["Fichier"] or nom_fichier(sel),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_reg_word_{sel}",
                    )
                # PDF depuis Donnees enregistrées
                try:
                    if lg.get("Donnees"):
                        pdf_b = construire_pdf(_jload(lg["Donnees"]))
                        s2.download_button(
                            "⬇️ Télécharger le PDF",
                            data=pdf_b,
                            file_name=nom_fichier_pdf(sel),
                            mime="application/pdf",
                            key=f"dl_reg_pdf_{sel}",
                        )
                except Exception as e:
                    s2.caption(f"PDF : {e}")
                if lg["Donnees"] and s3.button("🔁 Régénérer le Word"):
                    try:
                        DOSSIER_BL.mkdir(exist_ok=True)
                        (DOSSIER_BL / (lg["Fichier"] or nom_fichier(sel))).write_bytes(construire_docx(_jload(lg["Donnees"])))
                        ss["BL_flash"] = ("success", f"Word du bon {sel} régénéré.")
                        st.rerun()
                    except RuntimeError as e:
                        st.error(str(e))

    # =================== VÉRIFICATION ===================
    with tab_ver:
        st.write("Contrôle l'authenticité d'un bon reçu ou présenté : saisis sa référence et le code imprimé en bas du bon.")
        v1, v2 = st.columns(2)
        ref_v = v1.text_input("Référence du bon", key="BL_vref")
        code_v = v2.text_input("Code de vérification", placeholder="XXXX-XXXX-XXXX", key="BL_vcode")
        if st.button("🔐 Vérifier"):
            if not ref_v.strip() or not code_v.strip():
                st.warning("Renseigne la référence et le code.")
            else:
                ok, msg, lg = verifier_bon(ref_v, code_v, reg)
                (st.success if ok else st.error)(msg)
                if lg is not None:
                    st.write(f"**{lg['REF']}** · livraison du {lg['Date_Livraison']:%d/%m/%Y} · camion {lg['Camion']} · "
                             f"{fmt(lg['Quantite'], 3)} {lg['Unite']} · proforma {lg['Proforma']} · statut {lg['Statut']}")
