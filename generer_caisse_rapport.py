# -*- coding: utf-8 -*-
"""Rapports PDF Caisse NAP SARL."""
from __future__ import annotations

from datetime import datetime
from io import BytesIO

import pandas as pd

ENTREPRISE = "STE NAP SARL"
NAVY = "1F3A5F"
LIGHT = "EEF3F8"


def _s(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def fmt(n):
    try:
        return f"{float(n):,.0f}".replace(",", " ")
    except Exception:
        return "0"


def generer_pdf_mouvements(df, infos: dict, kpis: dict) -> bytes:
    """Rapport PDF des mouvements filtrés."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
    )
    from xml.sax.saxutils import escape

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=12 * mm, rightMargin=12 * mm,
        topMargin=12 * mm, bottomMargin=14 * mm,
        title="Rapport Caisse", author=ENTREPRISE,
    )
    W = landscape(A4)[0] - 24 * mm

    st_t = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=14, alignment=TA_CENTER,
                          textColor=colors.HexColor("#" + NAVY), leading=17)
    st_s = ParagraphStyle("s", fontName="Helvetica", fontSize=9, alignment=TA_CENTER,
                          textColor=colors.HexColor("#444444"), leading=11)
    st_h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=11,
                          textColor=colors.HexColor("#" + NAVY), spaceBefore=6, spaceAfter=3)
    st_c = ParagraphStyle("c", fontName="Helvetica", fontSize=8, leading=10)
    st_r = ParagraphStyle("r", fontName="Helvetica", fontSize=8, leading=10, alignment=TA_RIGHT)

    el = [
        Paragraph(ENTREPRISE, st_t),
        Paragraph("RAPPORT DE CAISSE — MOUVEMENTS", st_s),
        Paragraph(
            " · ".join(f"{escape(str(k))} : {escape(str(v))}" for k, v in infos.items()),
            st_s,
        ),
        Spacer(1, 4 * mm),
    ]

    # KPI
    items = [
        ("Mouvements", kpis.get("nb", 0)),
        ("Dépenses", fmt(kpis.get("depenses", 0)) + " F"),
        ("Entrées", fmt(kpis.get("entrees", 0)) + " F"),
        ("Solde", fmt(kpis.get("solde", 0)) + " F"),
    ]
    tk = Table(
        [[Paragraph(f"<b>{escape(str(v))}</b><br/><font size='7' color='#666'>{escape(k)}</font>", st_s)
          for k, v in items]],
        colWidths=[W / 4] * 4,
    )
    tk.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#" + LIGHT)),
        ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#9AA7B5")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    el += [tk, Spacer(1, 3 * mm), Paragraph("Détail des mouvements", st_h)]

    if df is None or df.empty:
        el.append(Paragraph("Aucun mouvement.", st_c))
    else:
        data = [["Date", "Type", "Catégorie", "Libellé", "Tiers", "Montant", "Mode", "Réf."]]
        aff = df.copy()
        if "Date" in aff.columns:
            aff["Date"] = pd.to_datetime(aff["Date"], errors="coerce")
        for _, r in aff.iterrows():
            d = r.get("Date")
            dtxt = d.strftime("%d/%m/%Y") if hasattr(d, "strftime") and pd.notna(d) else ""
            data.append([
                dtxt,
                _s(r.get("Type")),
                Paragraph(escape(_s(r.get("Categorie"))), st_c),
                Paragraph(escape(_s(r.get("Libelle"))[:60]), st_c),
                Paragraph(escape(_s(r.get("Tiers"))[:25]), st_c),
                Paragraph(fmt(r.get("Montant")), st_r),
                _s(r.get("Mode")),
                Paragraph(escape(_s(r.get("Reference"))[:18]), st_c),
            ])
        # total row
        data.append([
            "", "", "", Paragraph("<b>TOTAL DÉPENSES</b>", st_c), "",
            Paragraph(f"<b>{fmt(kpis.get('depenses', 0))}</b>", st_r), "", "",
        ])
        cw = [22*mm, 18*mm, 28*mm, 50*mm, 32*mm, 28*mm, 22*mm, W - 200*mm]
        t = Table(data, colWidths=cw, repeatRows=1)
        sty = [
            ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + NAVY)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONT", (0, 1), (-1, -1), "Helvetica", 8),
            ("ALIGN", (5, 0), (5, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#9AA7B5")),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#" + LIGHT)),
        ]
        for i in range(1, len(data) - 1):
            if i % 2 == 0:
                sty.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F8FAFC")))
            typ = str(data[i][1]).upper()
            if typ == "DEPENSE":
                sty.append(("TEXTCOLOR", (5, i), (5, i), colors.HexColor("#B00020")))
            elif typ == "ENTREE":
                sty.append(("TEXTCOLOR", (5, i), (5, i), colors.HexColor("#1B7A3D")))
        t.setStyle(TableStyle(sty))
        el.append(t)

    def pied(canvas, d_):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#555555"))
        canvas.drawString(
            12 * mm, 8 * mm,
            f"{ENTREPRISE} — Caisse — {datetime.now():%d/%m/%Y %H:%M}",
        )
        canvas.drawRightString(landscape(A4)[0] - 12 * mm, 8 * mm, f"Page {d_.page}")
        canvas.restoreState()

    doc.build(el, onFirstPage=pied, onLaterPages=pied)
    buf.seek(0)
    return buf.getvalue()


def generer_pdf_comptes(sit_df, detail_df, infos: dict, kpis: dict, titre_compte="") -> bytes:
    """Rapport PDF situation comptes / un compte."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
    )
    from xml.sax.saxutils import escape

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=14 * mm, bottomMargin=14 * mm,
        title="Comptes Caisse", author=ENTREPRISE,
    )
    W = A4[0] - 28 * mm

    st_t = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=14, alignment=TA_CENTER,
                          textColor=colors.HexColor("#" + NAVY), leading=17)
    st_s = ParagraphStyle("s", fontName="Helvetica", fontSize=9, alignment=TA_CENTER,
                          textColor=colors.HexColor("#444444"), leading=11)
    st_h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=11,
                          textColor=colors.HexColor("#" + NAVY), spaceBefore=8, spaceAfter=3)
    st_c = ParagraphStyle("c", fontName="Helvetica", fontSize=8.5, leading=11)
    st_r = ParagraphStyle("r", fontName="Helvetica", fontSize=8.5, leading=11, alignment=TA_RIGHT)

    titre = "SITUATION DES COMPTES CAISSE"
    if titre_compte:
        titre = f"COMPTE CAISSE — {titre_compte}"

    el = [
        Paragraph(ENTREPRISE, st_t),
        Paragraph(titre, st_s),
        Paragraph(
            " · ".join(f"{escape(str(k))} : {escape(str(v))}" for k, v in infos.items()),
            st_s,
        ),
        Spacer(1, 4 * mm),
    ]

    items = [
        ("Dépenses", fmt(kpis.get("depenses", 0)) + " F"),
        ("Entrées", fmt(kpis.get("entrees", 0)) + " F"),
        ("Solde", fmt(kpis.get("solde", 0)) + " F"),
        ("Lignes", kpis.get("nb", 0)),
    ]
    tk = Table(
        [[Paragraph(f"<b>{escape(str(v))}</b><br/><font size='7' color='#666'>{escape(k)}</font>", st_s)
          for k, v in items]],
        colWidths=[W / 4] * 4,
    )
    tk.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#" + LIGHT)),
        ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#9AA7B5")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    el += [tk, Spacer(1, 3 * mm)]

    if sit_df is not None and not sit_df.empty and not titre_compte:
        el.append(Paragraph("Synthèse par tiers", st_h))
        data = [["Tiers", "Type", "Nb", "Dépenses", "Entrées", "Solde"]]
        for _, r in sit_df.iterrows():
            data.append([
                Paragraph(escape(_s(r.get("Tiers"))), st_c),
                _s(r.get("Type_Tiers")),
                str(int(r.get("Nb") or 0)),
                Paragraph(fmt(r.get("Total_Depenses")), st_r),
                Paragraph(fmt(r.get("Total_Entrees")), st_r),
                Paragraph(fmt(r.get("Solde")), st_r),
            ])
        t = Table(data, colWidths=[W*0.28, W*0.14, W*0.08, W*0.17, W*0.17, W*0.16], repeatRows=1)
        t.setStyle(TableStyle([
            ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + NAVY)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONT", (0, 1), (-1, -1), "Helvetica", 8),
            ("ALIGN", (2, 0), (-1, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#9AA7B5")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        el.append(t)

    if detail_df is not None and not detail_df.empty:
        el.append(Paragraph("Détail des mouvements", st_h))
        data = [["Date", "Type", "Catégorie", "Libellé", "Montant", "Mode"]]
        aff = detail_df.copy()
        if "Date" in aff.columns:
            aff["Date"] = pd.to_datetime(aff["Date"], errors="coerce")
        for _, r in aff.iterrows():
            d = r.get("Date")
            dtxt = d.strftime("%d/%m/%Y") if hasattr(d, "strftime") and pd.notna(d) else ""
            data.append([
                dtxt,
                _s(r.get("Type")),
                Paragraph(escape(_s(r.get("Categorie"))), st_c),
                Paragraph(escape(_s(r.get("Libelle"))[:40]), st_c),
                Paragraph(fmt(r.get("Montant")), st_r),
                _s(r.get("Mode")),
            ])
        t2 = Table(data, colWidths=[22*mm, 18*mm, 28*mm, W - 110*mm, 26*mm, 22*mm], repeatRows=1)
        t2.setStyle(TableStyle([
            ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + NAVY)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONT", (0, 1), (-1, -1), "Helvetica", 8),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#9AA7B5")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        el.append(t2)

    def pied(canvas, d_):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#555555"))
        canvas.drawString(14 * mm, 8 * mm, f"{ENTREPRISE} — Caisse comptes — {datetime.now():%d/%m/%Y %H:%M}")
        canvas.drawRightString(A4[0] - 14 * mm, 8 * mm, f"Page {d_.page}")
        canvas.restoreState()

    doc.build(el, onFirstPage=pied, onLaterPages=pied)
    buf.seek(0)
    return buf.getvalue()



def generer_pdf_fermeture(cloture: dict, infos: dict) -> bytes:
    """Rapport de fermeture de caisse journalière."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
    )
    from xml.sax.saxutils import escape

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=14 * mm, bottomMargin=16 * mm,
        title="Fermeture de caisse", author=ENTREPRISE,
    )
    W = A4[0] - 28 * mm
    jour = cloture.get("date")
    dtxt = jour.strftime("%d/%m/%Y") if hasattr(jour, "strftime") else str(jour)

    st_t = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=14, alignment=TA_CENTER,
                          textColor=colors.HexColor("#" + NAVY), leading=17)
    st_s = ParagraphStyle("s", fontName="Helvetica", fontSize=10, alignment=TA_CENTER,
                          textColor=colors.HexColor("#444444"), leading=13)
    st_h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=11,
                          textColor=colors.HexColor("#" + NAVY), spaceBefore=8, spaceAfter=4)
    st_c = ParagraphStyle("c", fontName="Helvetica", fontSize=9, leading=11)
    st_r = ParagraphStyle("r", fontName="Helvetica", fontSize=9, leading=11, alignment=TA_RIGHT)
    st_b = ParagraphStyle("b", fontName="Helvetica-Bold", fontSize=10, leading=12)

    el = [
        Paragraph(ENTREPRISE, st_t),
        Paragraph("RAPPORT DE FERMETURE DE CAISSE", st_s),
        Paragraph(f"Journée du <b>{dtxt}</b>", st_s),
        Paragraph(
            " · ".join(f"{escape(str(k))} : {escape(str(v))}" for k, v in infos.items()),
            ParagraphStyle("i", parent=st_s, fontSize=8),
        ),
        Spacer(1, 5 * mm),
    ]

    # Récap soldes
    rows = [
        ["Libellé", "Montant (FCFA)"],
        ["Solde de départ", fmt(cloture.get("solde_depart", 0))],
        [f"+ Entrées ({int(cloture.get('nb_entrees', 0))})", fmt(cloture.get("entrees", 0))],
        [f"− Dépenses ({int(cloture.get('nb_depenses', 0))})", fmt(cloture.get("depenses", 0))],
        ["Solde de clôture", fmt(cloture.get("solde_cloture", 0))],
    ]
    t = Table(rows, colWidths=[W * 0.55, W * 0.45])
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 9),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + NAVY)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONT", (0, 1), (-1, -2), "Helvetica", 10),
        ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 11),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E8F5E9")),
        ("TEXTCOLOR", (1, 2), (1, 2), colors.HexColor("#1B7A3D")),
        ("TEXTCOLOR", (1, 3), (1, 3), colors.HexColor("#B00020")),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#9AA7B5")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    el += [t, Spacer(1, 5 * mm)]

    mv = cloture.get("mouvements")
    el.append(Paragraph("Mouvements de la journée", st_h))
    if mv is None or (hasattr(mv, "empty") and mv.empty):
        el.append(Paragraph("Aucun mouvement enregistré ce jour.", st_c))
    else:
        data = [["Heure / Réf.", "Type", "Catégorie", "Libellé", "Tiers", "Montant", "Mode"]]
        aff = mv.copy()
        if "Date" in aff.columns:
            aff = aff.sort_values("ID") if "ID" in aff.columns else aff
        for _, r in aff.iterrows():
            data.append([
                Paragraph(escape(_s(r.get("Reference")) or _s(r.get("Heure_Creation"))[:16]), st_c),
                _s(r.get("Type")),
                Paragraph(escape(_s(r.get("Categorie"))), st_c),
                Paragraph(escape(_s(r.get("Libelle"))[:35]), st_c),
                Paragraph(escape(_s(r.get("Tiers"))[:18]), st_c),
                Paragraph(fmt(r.get("Montant")), st_r),
                _s(r.get("Mode")),
            ])
        t2 = Table(
            data,
            colWidths=[28*mm, 18*mm, 24*mm, W-130*mm, 28*mm, 24*mm, 20*mm],
            repeatRows=1,
        )
        sty = [
            ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 7.5),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + NAVY)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONT", (0, 1), (-1, -1), "Helvetica", 7.5),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#9AA7B5")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]
        for i in range(1, len(data)):
            if str(data[i][1]).upper() == "DEPENSE":
                sty.append(("TEXTCOLOR", (5, i), (5, i), colors.HexColor("#B00020")))
            elif str(data[i][1]).upper() == "ENTREE":
                sty.append(("TEXTCOLOR", (5, i), (5, i), colors.HexColor("#1B7A3D")))
        t2.setStyle(TableStyle(sty))
        el.append(t2)

    el.append(Spacer(1, 10 * mm))
    sig = Table(
        [[
            Paragraph("Caissier / Responsable<br/><br/><br/>_______________", st_c),
            Paragraph("Contrôle / Direction<br/><br/><br/>_______________", st_c),
        ]],
        colWidths=[W / 2, W / 2],
    )
    sig.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOX", (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
    ]))
    el.append(sig)

    def pied(canvas, d_):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#555555"))
        canvas.drawString(14 * mm, 8 * mm, f"{ENTREPRISE} — Fermeture caisse {dtxt} — {datetime.now():%d/%m/%Y %H:%M}")
        canvas.drawRightString(A4[0] - 14 * mm, 8 * mm, f"Page {d_.page}")
        canvas.restoreState()

    doc.build(el, onFirstPage=pied, onLaterPages=pied)
    buf.seek(0)
    return buf.getvalue()
