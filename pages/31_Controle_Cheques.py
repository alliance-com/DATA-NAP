# -*- coding: utf-8 -*-
"""
Contrôle des chèques — par lots de 100
--------------------------------------
Objectif : vérifier que les n° déchirés figurent dans le suivi,
et lister les TROUS (numéros manquants) dans chaque lot réellement utilisé.

• Source : Suivi de paiement et/ou fichier importé
• Classement par lots de 100 (ex. 2000300-2000399)
• Uniquement les lots où au moins 1 chèque a été trouvé (pas de millions de lignes)
• Filtre banque + période
• Export Excel + PDF
"""
from __future__ import annotations

import io
import re
from datetime import datetime, date
from pathlib import Path

import pandas as pd

ENTREPRISE = "STE NAP SARL"
FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"
FEUILLE_DEPENSES = "DEPENSES"
RACINE = Path(__file__).resolve().parent.parent
NB_CHIFFRES = 7
TAILLE_LOT = 100
NAVY, LIGHT = "1F3A5F", "EEF3F8"

MODES = {
    "entre": "Entre le plus petit et le plus grand n° trouvés du lot (recommandé)",
    "apres_min": "Du plus petit n° trouvé jusqu’à la fin du lot (…99)",
    "complet": "Lot entier de 100 (…00 à …99) — à éviter si le carnet ne commence pas à …00",
}


def fmt(n) -> str:
    try:
        return f"{int(round(float(n))):,}".replace(",", " ")
    except Exception:
        return ""


def _norm(s) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def extraire_numeros(valeur, nb=NB_CHIFFRES) -> list:
    if valeur is None or (isinstance(valeur, float) and pd.isna(valeur)):
        return []
    t = str(int(valeur)) if isinstance(valeur, float) and valeur.is_integer() else str(valeur)
    return [int(x) for x in re.findall(rf"(?<!\d)\d{{{nb}}}(?!\d)", t)]


def plages(nums) -> str:
    nums = sorted(set(int(n) for n in nums))
    if not nums:
        return "—"
    out, deb, prec = [], nums[0], nums[0]
    for n in nums[1:]:
        if n == prec + 1:
            prec = n
            continue
        out.append(str(deb) if deb == prec else f"{deb}-{prec}")
        deb = prec = n
    out.append(str(deb) if deb == prec else f"{deb}-{prec}")
    return ", ".join(out)


def chemin_base(chemin=None) -> Path:
    c = Path(chemin or FICHIER_PAIEMENTS)
    return c if c.exists() else RACINE / c.name


def charger_base(chemin=None, nb=NB_CHIFFRES):
    p = chemin_base(chemin)
    if not p.exists():
        raise FileNotFoundError(f"« {p.name} » introuvable dans le dossier du projet.")
    xl = pd.ExcelFile(p)
    feuille = FEUILLE_DEPENSES if FEUILLE_DEPENSES in xl.sheet_names else xl.sheet_names[0]
    df = None
    for hdr in range(0, 8):
        tmp = pd.read_excel(p, sheet_name=feuille, header=hdr)
        if any("cheque" in _norm(c) for c in tmp.columns):
            df = tmp
            break
    if df is None:
        raise ValueError("Colonne chèque introuvable.")

    cols = {_norm(c): c for c in df.columns}

    def col(*cands):
        for k, v in cols.items():
            if any(c in k for c in cands):
                return v
        return None

    c_num = next(
        (v for k, v in cols.items() if "cheque" in k and ("num" in k or "no" in k or k == "cheque")),
        None,
    ) or col("cheque")
    c_mode = col("modepaiement", "mode")
    c_date, c_trx = col("date"), col("transaction")
    c_tiers = col("prestataire", "fournisseur", "beneficiaire")
    c_mt, c_bq = col("montant"), col("banque")

    lignes, sans = [], []
    for _, r in df.iterrows():
        mode = str(r[c_mode]).upper() if c_mode else ""
        est = "CHEQUE" in mode or "CHQ" in mode or "CHÈQUE" in mode
        nums = extraire_numeros(r[c_num], nb)
        base = {
            "Date": r[c_date] if c_date else None,
            "Transaction": r[c_trx] if c_trx else "",
            "Tiers": r[c_tiers] if c_tiers else "",
            "Montant": r[c_mt] if c_mt else None,
            "Banque": str(r[c_bq]).strip() if c_bq and pd.notna(r[c_bq]) else "",
            "Source": "Suivi de paiement",
        }
        if nums:
            for n in nums:
                lignes.append({"Numero": n, **base})
        elif est:
            sans.append({**base, "Saisi": "" if pd.isna(r[c_num]) else str(r[c_num])})

    ch = pd.DataFrame(lignes)
    if not ch.empty:
        ch["Date"] = pd.to_datetime(ch["Date"], errors="coerce")
        ch["Montant"] = pd.to_numeric(ch["Montant"], errors="coerce")
        ch["Numero"] = ch["Numero"].astype(int)
    sans_df = pd.DataFrame(sans)
    if not sans_df.empty and "Date" in sans_df.columns:
        sans_df["Date"] = pd.to_datetime(sans_df["Date"], errors="coerce")
    return ch, sans_df


def lire_import(data: bytes, nom: str):
    if nom.lower().endswith((".csv", ".txt")):
        for enc in ("utf-8-sig", "latin-1"):
            try:
                df = pd.read_csv(
                    io.BytesIO(data), sep=None, engine="python",
                    header=None, dtype=object, encoding=enc,
                )
                return {"CSV": df}, ["CSV"]
            except Exception:
                continue
        raise ValueError("CSV illisible.")
    xl = pd.ExcelFile(io.BytesIO(data))
    return (
        {s: pd.read_excel(io.BytesIO(data), sheet_name=s, header=None, dtype=object)
         for s in xl.sheet_names},
        list(xl.sheet_names),
    )


def detecter_colonne(raw: pd.DataFrame, nb=NB_CHIFFRES) -> pd.DataFrame:
    res = []
    for j in raw.columns:
        compte, premier = 0, None
        for i, v in raw[j].items():
            if extraire_numeros(v, nb):
                compte += 1
                if premier is None:
                    premier = i
        if not compte:
            continue
        ent = ""
        for k in range((premier or 0) - 1, max((premier or 0) - 4, -1), -1):
            v = raw.iat[k, j] if k >= 0 else None
            if isinstance(v, str) and v.strip():
                ent = v.strip()
                break
        bonus = 1000 if re.search(r"ch[eèé]que|chq|cheq", ent, re.I) else 0
        res.append({"col": j, "en_tete": ent or f"col {j+1}", "n": compte, "score": compte + bonus})
    return (
        pd.DataFrame(res).sort_values("score", ascending=False).reset_index(drop=True)
        if res else pd.DataFrame(columns=["col", "en_tete", "n", "score"])
    )


def charger_import(raw, col, nb=NB_CHIFFRES, nom="Import") -> pd.DataFrame:
    lignes = []
    for i, r in raw.iterrows():
        for n in extraire_numeros(r[col], nb):
            autres = [str(v).strip() for j, v in r.items() if j != col and pd.notna(v) and str(v).strip()][:3]
            lignes.append({
                "Numero": n, "Date": None, "Transaction": "",
                "Tiers": " · ".join(autres), "Montant": None,
                "Banque": "", "Source": nom,
            })
    df = pd.DataFrame(lignes)
    if not df.empty:
        df["Numero"] = df["Numero"].astype(int)
    return df


def analyser(cheques: pd.DataFrame, mode="entre", taille=TAILLE_LOT):
    """
    Analyse uniquement les lots où au moins 1 chèque est présent.
    Ne commence JAMAIS avant le plus petit n° réellement trouvé dans le lot
    (ni avant le plus petit n° de la sélection / banque).
    """
    if cheques is None or cheques.empty:
        empty = pd.DataFrame()
        return empty, empty, empty, empty

    det = cheques.copy()
    det["Numero"] = det["Numero"].astype(int)
    plancher_global = int(det["Numero"].min())  # aucun manquant sous ce n°
    det["Lot_debut"] = (det["Numero"] // taille) * taille
    det["Lot"] = det["Lot_debut"].map(lambda d: f"{d} - {d + taille - 1}")
    det = det.sort_values(["Numero", "Date"]).reset_index(drop=True)

    counts = det["Numero"].value_counts()
    doublons = det[det["Numero"].map(counts) > 1].copy()

    syn_rows, manq_rows = [], []
    for debut, g in det.groupby("Lot_debut"):
        trouves = set(int(x) for x in g["Numero"])
        mn, mx = min(trouves), max(trouves)
        # jamais avant le plus petit n° de la sélection
        mn = max(mn, plancher_global)
        lot_fin = int(debut) + taille - 1

        if mode == "complet":
            # lot entier mais borné au plancher global
            start = max(int(debut), plancher_global)
            attendu = range(start, lot_fin + 1)
        elif mode == "apres_min":
            # du premier trouvé dans le lot → fin du lot
            attendu = range(mn, lot_fin + 1)
        else:
            # "entre" (recommandé) : uniquement entre premier et dernier TROUVÉS
            attendu = range(mn, mx + 1)

        absents = [n for n in attendu if n not in trouves]
        lot = f"{int(debut)} - {lot_fin}"
        syn_rows.append({
            "Lot": lot,
            "Premier trouvé": mn,
            "Dernier trouvé": mx,
            "Trouvés": len(trouves),
            "Manquants": len(absents),
            "Doublons": sum(1 for n in trouves if counts[n] > 1),
            "Plages manquantes": plages(absents),
        })
        for n in absents:
            manq_rows.append({"Lot": lot, "Numéro manquant": n})

    syn = pd.DataFrame(syn_rows)
    manq = pd.DataFrame(manq_rows, columns=["Lot", "Numéro manquant"])
    return syn, manq, det, doublons


def excel_bytes(syn, manq, det, dbl, sans, infos) -> bytes:
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()
    # Limite sécurité manquants
    manq_export = manq.head(50000) if len(manq) > 50000 else manq
    detail = det.drop(columns=["Lot_debut"], errors="ignore") if not det.empty else pd.DataFrame()
    if not detail.empty and "Date" in detail.columns:
        detail = detail.copy()
        detail["Date"] = pd.to_datetime(detail["Date"], errors="coerce").dt.strftime("%d/%m/%Y")

    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        syn.to_excel(w, sheet_name="Synthèse par lot", index=False)
        manq_export.to_excel(w, sheet_name="Manquants", index=False)
        if not detail.empty:
            cols = [c for c in ["Lot", "Numero", "Date", "Transaction", "Tiers", "Montant", "Banque", "Source"] if c in detail.columns]
            detail[cols].to_excel(w, sheet_name="Détail", index=False)
        if dbl is not None and not dbl.empty:
            d2 = dbl.drop(columns=["Lot_debut"], errors="ignore")
            d2.to_excel(w, sheet_name="Doublons", index=False)
        if sans is not None and not sans.empty:
            sans.to_excel(w, sheet_name="Sans numéro", index=False)
        pd.DataFrame(list(infos.items()), columns=["Paramètre", "Valeur"]).to_excel(
            w, sheet_name="Paramètres", index=False
        )
        thin = Side(style="thin", color="9AA7B5")
        for ws in w.book.worksheets:
            ws.freeze_panes = "A2"
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor=NAVY)
            for col in ws.columns:
                letter = get_column_letter(col[0].column)
                largeur = max([len(str(c.value or "")) for c in col[:200]] + [10]) + 2
                ws.column_dimensions[letter].width = min(largeur, 55)
                for c in col[1:]:
                    c.border = Border(bottom=thin)
        # total synthèse
        if "Synthèse par lot" in w.book.sheetnames and not syn.empty:
            ws = w.book["Synthèse par lot"]
            n = ws.max_row + 1
            ws.cell(n, 1, "TOTAL").font = Font(bold=True)
            for col_idx, name in ((4, "Trouvés"), (5, "Manquants"), (6, "Doublons")):
                if name in syn.columns:
                    ws.cell(n, col_idx, int(syn[name].sum())).font = Font(bold=True)
    return buf.getvalue()


def pdf_bytes(syn, manq, dbl, sans, infos, kpis) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, KeepTogether
    from xml.sax.saxutils import escape

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=12 * mm, rightMargin=12 * mm,
        topMargin=12 * mm, bottomMargin=14 * mm,
        title="Contrôle chèques", author=ENTREPRISE,
    )
    W = landscape(A4)[0] - 24 * mm
    st_t = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=14, alignment=TA_CENTER,
                          textColor=colors.HexColor("#" + NAVY), leading=17)
    st_s = ParagraphStyle("s", fontName="Helvetica", fontSize=9, alignment=TA_CENTER,
                          textColor=colors.HexColor("#444"), leading=11)
    st_h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=11,
                          textColor=colors.HexColor("#" + NAVY), spaceBefore=8, spaceAfter=3)
    st_c = ParagraphStyle("c", fontName="Helvetica", fontSize=8, leading=10)

    el = [
        Paragraph(ENTREPRISE, st_t),
        Paragraph("CONTRÔLE DES CHÈQUES — LOTS DE 100 — NUMÉROS MANQUANTS", st_s),
        Paragraph(" · ".join(f"{k} : {escape(str(v))}" for k, v in infos.items()), st_s),
        Spacer(1, 4 * mm),
    ]
    kp = [[Paragraph(f"<b>{fmt(v)}</b><br/><font size='7' color='#666'>{escape(k)}</font>", st_s)
           for k, v in kpis.items()]]
    tk = Table(kp, colWidths=[W / max(len(kpis), 1)] * len(kpis))
    tk.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#" + LIGHT)),
        ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#9AA7B5")),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    el += [tk, Spacer(1, 3 * mm), Paragraph("Synthèse par lot", st_h)]

    data = [["Lot", "Premier", "Dernier", "Trouvés", "Manquants", "Doublons", "Plages manquantes"]]
    for _, r in syn.iterrows():
        data.append([
            r["Lot"], str(r["Premier trouvé"]), str(r["Dernier trouvé"]),
            str(r["Trouvés"]), str(r["Manquants"]), str(r["Doublons"]),
            Paragraph(escape(str(r["Plages manquantes"])), st_c),
        ])
    cw = [32 * mm, 22 * mm, 22 * mm, 18 * mm, 20 * mm, 18 * mm, W - 132 * mm]
    t = Table(data, colWidths=cw, repeatRows=1)
    sty = [
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + NAVY)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONT", (0, 1), (-1, -1), "Helvetica", 8),
        ("ALIGN", (1, 0), (5, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#9AA7B5")),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for i, (_, r) in enumerate(syn.iterrows(), start=1):
        if r["Manquants"]:
            sty.append(("TEXTCOLOR", (4, i), (4, i), colors.HexColor("#B00020")))
            sty.append(("FONT", (4, i), (4, i), "Helvetica-Bold", 8))
        if i % 2 == 0:
            sty.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F6F8FB")))
    t.setStyle(TableStyle(sty))
    el.append(t)

    if not manq.empty:
        el.append(Paragraph("Détail des manquants par lot", st_h))
        for lot, g in manq.groupby("Lot", sort=False):
            nums = list(g["Numéro manquant"])
            el.append(KeepTogether([
                Paragraph(f"<b>Lot {lot}</b> — {len(nums)} manquant(s) : {escape(plages(nums))}", st_c),
                Spacer(1, 1.5 * mm),
            ]))

    def pied(canvas, d_):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#555555"))
        canvas.drawString(12 * mm, 8 * mm, f"{ENTREPRISE} — Contrôle chèques — {datetime.now():%d/%m/%Y %H:%M}")
        canvas.drawRightString(landscape(A4)[0] - 12 * mm, 8 * mm, f"Page {d_.page}")
        canvas.restoreState()

    doc.build(el, onFirstPage=pied, onLaterPages=pied)
    return buf.getvalue()


# ====================== UI ======================
import streamlit as st

st.set_page_config(page_title="Contrôle chèques", page_icon="🧾", layout="wide")
st.title("🧾 Contrôle des chèques")
st.info(
    "**Comment utiliser**\n\n"
    "1. Choisissez la source (Suivi de paiement recommandé)\n"
    "2. Filtrez **une banque** (ex. ECOBANK) — obligatoire pour un contrôle clair\n"
    "3. Optionnel : période\n"
    "4. Lisez la synthèse par lot de 100 : colonne **Manquants** = chèques non retrouvés\n"
    "5. Téléchargez Excel ou PDF\n\n"
    "On ne contrôle que les lots où au moins un chèque apparaît (pas de millions de lignes)."
)

SRC_BASE, SRC_IMP, SRC_DEUX = "Suivi de paiement", "Fichier importé", "Les deux"
c1, c2, c3 = st.columns([2, 2, 1])
source = c1.radio("Source", [SRC_BASE, SRC_IMP, SRC_DEUX], horizontal=True)
mode = c2.selectbox("Mode des manquants", list(MODES), format_func=MODES.get, index=0)  # entre = recommandé
nb = c3.number_input("Chiffres n°", min_value=5, max_value=12, value=NB_CHIFFRES)

with st.expander("Filtrer par période", expanded=False):
    act_per = st.checkbox("Activer la période", value=False)
    p1, p2 = st.columns(2)
    debut_p = p1.date_input("Du", value=date(date.today().year, 1, 1), disabled=not act_per)
    fin_p = p2.date_input("Au", value=date.today(), disabled=not act_per)

cheques = pd.DataFrame()
sans = pd.DataFrame()
infos = {"Source": source, "Mode": MODES[mode].split(" (")[0]}
parties = []

if source in (SRC_BASE, SRC_DEUX):
    with st.expander("Fichier base", expanded=False):
        chemin = st.text_input("Chemin", value=FICHIER_PAIEMENTS)
    try:
        ch_b, sans = charger_base(chemin, int(nb))
        banques = sorted({b for b in ch_b["Banque"] if b}) if not ch_b.empty else []
        bq = st.multiselect(
            "Banque(s) — choisissez ECOBANK pour la série Ecobank",
            banques, default=[],
            placeholder="Sélectionnez une banque…",
        )
        if bq:
            ch_b = ch_b[ch_b["Banque"].isin(bq)]
            if not sans.empty and "Banque" in sans.columns:
                sans = sans[sans["Banque"].isin(bq)]
            infos["Banque"] = ", ".join(bq)
        elif not ch_b.empty:
            st.warning("Conseil : sélectionnez **une banque** pour éviter de mélanger plusieurs séries de chéquiers.")
        parties.append(ch_b)
        st.caption(
            f"Base : **{ch_b['Numero'].nunique() if not ch_b.empty else 0}** n° distincts"
            + (f" · plus petit n° : **{int(ch_b['Numero'].min())}** · plus grand : **{int(ch_b['Numero'].max())}**"
               if not ch_b.empty else "")
        )
        infos["Fichier"] = chemin_base(chemin).name
    except Exception as e:
        st.error(str(e))

if source in (SRC_IMP, SRC_DEUX):
    up = st.file_uploader("Fichier chèque (Excel / CSV)", type=["xlsx", "xlsm", "csv", "txt"])
    if up is not None:
        try:
            feuilles, noms = lire_import(up.getvalue(), up.name)
            dets = {s: detecter_colonne(r, int(nb)) for s, r in feuilles.items()}
            meilleure = max(noms, key=lambda s: int(dets[s]["n"].max()) if not dets[s].empty else 0)
            u1, u2 = st.columns(2)
            feuille = u1.selectbox("Feuille", noms, index=noms.index(meilleure)) if len(noms) > 1 else noms[0]
            det = dets[feuille]
            if det.empty:
                st.error(f"Aucun n° de {int(nb)} chiffres.")
            else:
                labs = [f"{r.en_tete} — {r.n} n°" for r in det.itertuples()]
                choix = u2.selectbox("Colonne n°", range(len(labs)), format_func=lambda i: labs[i])
                ch_f = charger_import(feuilles[feuille], int(det.iloc[choix]["col"]), int(nb), up.name)
                parties.append(ch_f)
                infos["Import"] = up.name
                st.caption(f"Import : **{ch_f['Numero'].nunique()}** n°")
        except Exception as e:
            st.error(str(e))

if parties:
    cheques = pd.concat(parties, ignore_index=True)

if act_per and not cheques.empty and "Date" in cheques.columns:
    cheques = cheques.copy()
    cheques["Date"] = pd.to_datetime(cheques["Date"], errors="coerce")
    m = cheques["Date"].notna()
    cheques = cheques[(~m) | ((cheques["Date"].dt.date >= debut_p) & (cheques["Date"].dt.date <= fin_p))]
    infos["Période"] = f"{debut_p:%d/%m/%Y} → {fin_p:%d/%m/%Y}"
    if not sans.empty and "Date" in sans.columns:
        sans = sans.copy()
        sans["Date"] = pd.to_datetime(sans["Date"], errors="coerce")
        sans = sans[(sans["Date"].isna()) | ((sans["Date"].dt.date >= debut_p) & (sans["Date"].dt.date <= fin_p))]

if cheques.empty:
    st.info("Aucun chèque chargé. Vérifiez le fichier ou la banque sélectionnée.")
    st.stop()

plancher = int(cheques["Numero"].min())
plafond = int(cheques["Numero"].max())
st.success(
    f"Contrôle calé sur les n° **réellement présents** dans la sélection : "
    f"de **{plancher}** à **{plafond}**. "
    f"Aucun trou n’est signalé avant {plancher}."
)
infos["Plus petit n°"] = plancher
infos["Plus grand n°"] = plafond

syn, manq, det, dbl = analyser(cheques, mode=mode)
kpis = {
    "N° distincts": int(cheques["Numero"].nunique()),
    "Lots touchés": len(syn),
    "Manquants": len(manq),
    "Doublons": int(dbl["Numero"].nunique()) if not dbl.empty else 0,
}
infos["Édité le"] = f"{datetime.now():%d/%m/%Y %H:%M}"

k = st.columns(4)
for col, (lab, val) in zip(k, kpis.items()):
    col.metric(lab, fmt(val))

if len(manq) == 0:
    st.success("Aucun trou dans les lots analysés : tous les n° attendus (selon le mode) sont présents.")
else:
    st.warning(f"**{len(manq)} numéro(s) manquant(s)** répartis sur **{len(syn[syn['Manquants']>0])}** lot(s).")

t1, t2, t3, t4 = st.tabs([
    "Synthèse par lot",
    f"Manquants ({len(manq)})",
    "Détail des chèques",
    f"Doublons / sans n°",
])
with t1:
    st.dataframe(syn, hide_index=True, width="stretch",
                 column_config={"Plages manquantes": st.column_config.TextColumn(width="large")})
with t2:
    if manq.empty:
        st.write("Aucun manquant.")
    else:
        lot = st.selectbox("Filtrer un lot", ["Tous"] + list(syn["Lot"]))
        vue = manq if lot == "Tous" else manq[manq["Lot"] == lot]
        st.dataframe(vue, hide_index=True, width="stretch", height=380)
with t3:
    aff = det.drop(columns=["Lot_debut"], errors="ignore").copy()
    if "Date" in aff.columns:
        aff["Date"] = pd.to_datetime(aff["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
    st.dataframe(aff, hide_index=True, width="stretch", height=400)
with t4:
    if dbl.empty:
        st.success("Pas de doublon.")
    else:
        st.dataframe(dbl.drop(columns=["Lot_debut"], errors="ignore"), hide_index=True, width="stretch")
    if not sans.empty:
        st.write("Paiements CHEQUE sans numéro :")
        s = sans.copy()
        if "Date" in s.columns:
            s["Date"] = pd.to_datetime(s["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(s, hide_index=True, width="stretch")

st.markdown("---")
st.subheader("Télécharger le rapport")
horo = f"{datetime.now():%Y%m%d_%H%M}"
d1, d2 = st.columns(2)

try:
    xbytes = excel_bytes(syn, manq, det, dbl, sans, infos)
    d1.download_button(
        "📊 Télécharger Excel",
        data=xbytes,
        file_name=f"Controle_cheques_{horo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
except Exception as e:
    d1.error(f"Excel : {e}")

try:
    pbytes = pdf_bytes(syn, manq, dbl, sans, infos, kpis)
    d2.download_button(
        "📄 Télécharger PDF",
        data=pbytes,
        file_name=f"Controle_cheques_{horo}.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=True,
    )
except Exception as e:
    d2.error(f"PDF : {e}")
