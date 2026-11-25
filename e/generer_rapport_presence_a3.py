# -*- coding: utf-8 -*-
"""Rapport de présence A3 — EXPOSÉ N° 3 · URHC PAROISSE ZOGBADJÉ."""
from __future__ import annotations

from datetime import datetime
from io import BytesIO

import pandas as pd


MOIS_FR = [
    "", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
]

# ZapfDingbats : "4" = ✔ , "8" = ✘ (Helvetica n'a pas ces symboles)
COCHE = '<font name="ZapfDingbats">4</font>'
CROIX = '<font name="ZapfDingbats">8</font>'


def generer_pdf_presence_a3(annee: int, mois: int, dims, lignes) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, HRFlowable,
    )
    from xml.sax.saxutils import escape

    buf = BytesIO()
    page = landscape(A3)
    doc = SimpleDocTemplate(
        buf, pagesize=page,
        leftMargin=12 * mm, rightMargin=12 * mm,
        topMargin=14 * mm, bottomMargin=14 * mm,
        title=f"Présences Exposé N° 3 — {MOIS_FR[mois]} {annee}",
        author="URHC Paroisse Zogbadjé",
    )
    W = page[0] - 24 * mm
    navy = colors.HexColor("#1A5276")
    green = colors.HexColor("#1E8449")
    light = colors.HexColor("#D5F5E3")
    grey = colors.HexColor("#555555")

    # ── Styles de l'entête (grands interlignes = pas collé) ──
    st_expose = ParagraphStyle("expose", fontName="Helvetica-Bold", fontSize=30,
                               leading=38, alignment=TA_CENTER, textColor=navy,
                               spaceAfter=6)
    st_paroisse = ParagraphStyle("paroisse", fontName="Helvetica-Bold", fontSize=20,
                                 leading=28, alignment=TA_CENTER,
                                 textColor=colors.HexColor("#2E86C1"), spaceAfter=10)
    st_registre = ParagraphStyle("registre", fontName="Helvetica-Bold", fontSize=15,
                                 leading=22, alignment=TA_CENTER,
                                 textColor=colors.HexColor("#222222"), spaceAfter=6)
    st_info = ParagraphStyle("info", fontName="Helvetica", fontSize=10, leading=15,
                             alignment=TA_CENTER, textColor=grey)

    st_s = ParagraphStyle("s", fontName="Helvetica", fontSize=10, leading=14,
                          alignment=TA_CENTER, textColor=colors.HexColor("#333"))
    st_c = ParagraphStyle("c", fontName="Helvetica", fontSize=9, leading=11, alignment=TA_CENTER)
    st_l = ParagraphStyle("l", fontName="Helvetica", fontSize=9, leading=11)
    st_h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=9, leading=12,
                          textColor=colors.white, alignment=TA_CENTER)
    st_v = ParagraphStyle("v", parent=st_c, textColor=green, fontSize=12, leading=12)
    st_tot = ParagraphStyle("tot", parent=st_c, fontName="Helvetica-Bold")

    el = [
        Paragraph("EXPOSÉ N° 3", st_expose),
        Paragraph("URHC — PAROISSE ZOGBADJÉ", st_paroisse),
        Spacer(1, 2 * mm),
        Paragraph(f"REGISTRE DE PRÉSENCE — {MOIS_FR[mois].upper()} {annee}", st_registre),
        Paragraph(
            "Présences validées uniquement &nbsp;(" + COCHE + ") &nbsp;·&nbsp; "
            f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
            st_info,
        ),
        Spacer(1, 5 * mm),
        HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#5DADE2"),
                   spaceBefore=2, spaceAfter=8),
    ]

    # En-têtes : N°, Nom, Prénom, puis chaque dimanche
    headers = ["N°", "Nom", "Prénom"]
    for i, d in enumerate(dims, 1):
        headers.append(f"S{i}<br/>{d.strftime('%d/%m')}")
    headers.append("Total")

    data = [[Paragraph(h, st_h) for h in headers]]

    for idx, L in enumerate(lignes, 1):
        row = [
            Paragraph(str(idx), st_c),
            Paragraph(escape(str(L.get("Nom") or "")), st_l),
            Paragraph(escape(str(L.get("Prenom") or "")), st_l),
        ]
        for d in dims:
            val = L.get(d.isoformat(), "")
            if val == "V":
                row.append(Paragraph(COCHE, st_v))
            elif val == "A":
                row.append(Paragraph("…", st_c))
            elif val == "R":
                row.append(Paragraph(CROIX, st_c))
            else:
                row.append(Paragraph("", st_c))
        row.append(Paragraph(f"<b>{L.get('Total_Valide', 0)}</b>", st_c))
        data.append(row)

    # Ligne de totaux par dimanche
    tot_row = [Paragraph("", st_c), Paragraph("TOTAL PRÉSENTS", st_tot), Paragraph("", st_c)]
    for d in dims:
        n = sum(1 for L in lignes if L.get(d.isoformat()) == "V")
        tot_row.append(Paragraph(str(n), st_tot))
    tot_row.append(Paragraph(str(sum(L.get("Total_Valide", 0) for L in lignes)), st_tot))
    data.append(tot_row)

    n_dim = max(len(dims), 1)
    w_nom = 50 * mm
    w_pre = 45 * mm
    w_num = 12 * mm
    w_tot = 18 * mm
    reste = W - w_num - w_nom - w_pre - w_tot
    w_dim = reste / n_dim
    col_w = [w_num, w_nom, w_pre] + [w_dim] * n_dim + [w_tot]

    t = Table(data, colWidths=col_w, repeatRows=1)
    sty = [
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (3, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#AAB7B8")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        # ligne de total
        ("BACKGROUND", (0, len(data) - 1), (-1, len(data) - 1), colors.HexColor("#D6EAF8")),
        ("LINEABOVE", (0, len(data) - 1), (-1, len(data) - 1), 1.2, navy),
    ]
    for i in range(1, len(data) - 1):
        if i % 2 == 0:
            sty.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F4F9FC")))
        for j, d in enumerate(dims):
            if lignes[i - 1].get(d.isoformat()) == "V":
                sty.append(("BACKGROUND", (3 + j, i), (3 + j, i), light))
    t.setStyle(TableStyle(sty))
    el.append(t)
    el.append(Spacer(1, 8 * mm))
    el.append(Paragraph(
        f"Légende : {COCHE} = présence validée par le responsable &nbsp;&nbsp; "
        "<b>…</b> = signalée en attente &nbsp;&nbsp; "
        f"{CROIX} = refusée &nbsp;&nbsp; "
        "case vide = absence / non signalée",
        st_s,
    ))
    el.append(Spacer(1, 2 * mm))
    el.append(Paragraph(
        "Exposé N° 3 — URHC Paroisse Zogbadjé — Document interne — "
        "Seules les présences validées sont officielles.",
        ParagraphStyle("f", parent=st_s, fontSize=8, textColor=colors.HexColor("#777")),
    ))

    def pied(canvas, d_):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#666"))
        canvas.drawString(12 * mm, 7 * mm,
                          f"Exposé N° 3 — URHC Paroisse Zogbadjé — Présences {MOIS_FR[mois]} {annee}")
        canvas.drawRightString(page[0] - 12 * mm, 7 * mm, f"Page {d_.page}")
        canvas.restoreState()

    doc.build(el, onFirstPage=pied, onLaterPages=pied)
    buf.seek(0)
    return buf.getvalue()
