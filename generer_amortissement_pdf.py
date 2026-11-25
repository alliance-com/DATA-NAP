# -*- coding: utf-8 -*-
"""Tableau d'amortissement PDF — NAP SARL."""
from __future__ import annotations

from datetime import datetime
from io import BytesIO

import pandas as pd


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


def generer_pdf_amortissement(pret, echeances_df, resume: dict) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, HRFlowable,
    )
    from xml.sax.saxutils import escape

    buf = BytesIO()
    page = landscape(A4)
    doc = SimpleDocTemplate(
        buf, pagesize=page,
        leftMargin=10 * mm, rightMargin=10 * mm,
        topMargin=10 * mm, bottomMargin=12 * mm,
        title="Tableau d'amortissement", author="STE NAP SARL",
    )
    W = page[0] - 20 * mm

    navy = colors.HexColor("#1A5276")
    light = colors.HexColor("#D6EAF8")
    st_t = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=13, alignment=TA_CENTER, textColor=navy)
    st_s = ParagraphStyle("s", fontName="Helvetica", fontSize=9, alignment=TA_CENTER, textColor=colors.HexColor("#333333"))
    st_c = ParagraphStyle("c", fontName="Helvetica", fontSize=7.5, leading=9)
    st_r = ParagraphStyle("r", fontName="Helvetica", fontSize=7.5, leading=9, alignment=TA_RIGHT)
    st_h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=7.5, leading=9, textColor=colors.white, alignment=TA_CENTER)

    pid = int(pret.get("ID_Pret") or 0)
    el = [
        Paragraph("STE NAP SARL", st_t),
        Paragraph("TABLEAU D'AMORTISSEMENT", st_t),
        Paragraph(
            f"Prêt #{pid} · {escape(_s(pret.get('Banque')))} · Réf. {escape(_s(pret.get('Reference')) or '—')}",
            st_s,
        ),
        Spacer(1, 2 * mm),
        HRFlowable(width="100%", thickness=1, color=colors.HexColor("#5DADE2"), spaceAfter=3),
    ]

    meta = [
        ["Objet", _s(pret.get("Objet")) or "—", "Montant autorisé", fmt(pret.get("Montant_Autorise")) + " FCFA"],
        ["Taux annuel", f"{float(pret.get('Taux_Annuel_Pct') or 0):.2f} %", "Durée", f"{int(pret.get('Duree_Mois') or 0)} mois"],
        ["Périodicité", _s(pret.get("Periodicite")), "Type", _s(pret.get("Type_Amort"))],
        ["Date début", (
            pret["Date_Debut"].strftime("%d/%m/%Y")
            if hasattr(pret.get("Date_Debut"), "strftime") and pd.notna(pret.get("Date_Debut"))
            else "—"
        ), "Capital amorti", fmt(resume.get("tire") or pret.get("Montant_Autorise")) + " FCFA"],
    ]
    tm = Table(meta, colWidths=[W * 0.15, W * 0.35, W * 0.15, W * 0.35])
    tm.setStyle(TableStyle([
        ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 8),
        ("FONT", (2, 0), (2, -1), "Helvetica-Bold", 8),
        ("FONT", (1, 0), (1, -1), "Helvetica", 8),
        ("FONT", (3, 0), (3, -1), "Helvetica", 8),
        ("BACKGROUND", (0, 0), (-1, -1), light),
        ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#AAB7B8")),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ]))
    el += [tm, Spacer(1, 3 * mm)]

    headers = ["N°", "Date", "Capital début", "Amortissement", "Intérêts", "Assurance", "Échéance", "Capital fin", "Statut", "Payé"]
    data = [[Paragraph(h, st_h) for h in headers]]
    if echeances_df is not None and not echeances_df.empty:
        for _, r in echeances_df.iterrows():
            d = r.get("Date_Echeance")
            dtxt = d.strftime("%d/%m/%Y") if hasattr(d, "strftime") and pd.notna(d) else ""
            data.append([
                str(int(r.get("N_Echeance") or 0)),
                dtxt,
                Paragraph(fmt(r.get("Capital_Debut")), st_r),
                Paragraph(fmt(r.get("Amortissement")), st_r),
                Paragraph(fmt(r.get("Interets")), st_r),
                Paragraph(fmt(r.get("Assurance")), st_r),
                Paragraph(fmt(r.get("Echeance_Totale")), st_r),
                Paragraph(fmt(r.get("Capital_Fin")), st_r),
                _s(r.get("Statut")),
                Paragraph(fmt(r.get("Montant_Paye")), st_r),
            ])
        # totals
        data.append([
            "", "", "",
            Paragraph(f"<b>{fmt(echeances_df['Amortissement'].sum())}</b>", st_r),
            Paragraph(f"<b>{fmt(echeances_df['Interets'].sum())}</b>", st_r),
            Paragraph(f"<b>{fmt(echeances_df['Assurance'].sum())}</b>", st_r),
            Paragraph(f"<b>{fmt(echeances_df['Echeance_Totale'].sum())}</b>", st_r),
            "",
            Paragraph("<b>TOTAL</b>", st_c),
            Paragraph(f"<b>{fmt(echeances_df['Montant_Paye'].sum())}</b>", st_r),
        ])

    cw = [12*mm, 22*mm, 28*mm, 28*mm, 24*mm, 22*mm, 28*mm, 28*mm, 22*mm, 24*mm]
    t = Table(data, colWidths=cw, repeatRows=1)
    sty = [
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONT", (0, 1), (-1, -1), "Helvetica", 7.5),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("ALIGN", (8, 1), (8, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#AAB7B8")),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("BACKGROUND", (0, -1), (-1, -1), light),
    ]
    for i in range(1, len(data) - 1):
        if i % 2 == 0:
            sty.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F8FBFD")))
    t.setStyle(TableStyle(sty))
    el.append(t)
    el.append(Spacer(1, 4 * mm))
    el.append(Paragraph(
        f"Capital restant estimé : <b>{fmt(resume.get('capital_restant'))} FCFA</b> · "
        f"Reste à payer (échéances) : <b>{fmt(resume.get('reste_a_payer'))} FCFA</b> · "
        f"Généré le {datetime.now():%d/%m/%Y %H:%M}",
        st_s,
    ))

    def pied(canvas, d_):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#666666"))
        canvas.drawString(10 * mm, 6 * mm, "STE NAP SARL — Emprunts / Tableau d'amortissement")
        canvas.drawRightString(page[0] - 10 * mm, 6 * mm, f"Page {d_.page}")
        canvas.restoreState()

    doc.build(el, onFirstPage=pied, onLaterPages=pied)
    buf.seek(0)
    return buf.getvalue()
