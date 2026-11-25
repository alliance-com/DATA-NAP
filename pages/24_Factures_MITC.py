# -*- coding: utf-8 -*-
"""
Page Streamlit — Gestion des factures MITC
MITHA INTERNATIONAL TRADING COMPANY (entité qui figure sur les factures)

- Import d'un ou plusieurs PDF de factures MECeF (une facture par page)
  -> extraction automatique : Date, N° Fact, Nom Client, Montant HT, TVA, AIB, Total
  -> remplissage de « Gestion_des_factures_MITC.xlsx »
- Formulaire de saisie manuelle à côté
- Journal + totaux + téléchargement de l'Excel

Dépendances : streamlit, pandas, openpyxl, pdfplumber, reportlab  (pip install pdfplumber reportlab)
À placer dans le dossier  DATA/pages/
"""
import io
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# ====================== PARAMÈTRES ======================
# Le fichier Excel est cherché à la racine du projet (dossier parent de /pages).
# Tu peux aussi changer le chemin dans l'écran (volet « Fichier Excel »).
FICHIER = Path(__file__).resolve().parent.parent / "Gestion_des_factures_MITC.xlsx"

# Valeur inscrite dans la colonne « N° Fact » :
#   "ref"   -> Réf Fact imprimée sur la facture (ex. 0002469)
#   "mecef" -> N° MECeF (ex. IT01005553-930)
NUM_FACT_SOURCE = "ref"

TAUX_TVA = 0.18
TAUX_AIB = 0.01
TOLERANCE = 2          # écart d'arrondi accepté (FCFA) entre HT+TVA+AIB et Total

COLS = ["Date", "N° Fact", "Nom Client", "Montant HT", "TVA", "AIB", "Total"]
FORMAT_MONTANT = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-'
FORMAT_DATE = "dd/mm/yyyy"

MOIS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
    "janvier": 1, "février": 2, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
    "juillet": 7, "août": 8, "aout": 8, "septembre": 9, "octobre": 10, "novembre": 11,
    "décembre": 12, "decembre": 12,
}


# ====================== EXTRACTION PDF ======================
def lire_pages_pdf(data: bytes) -> list:
    """Retourne le texte de chaque page du PDF (pdfplumber, sinon pypdf)."""
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            return [(p.extract_text() or "") for p in pdf.pages]
    except ImportError:
        pass
    try:
        from pypdf import PdfReader
        return [(p.extract_text() or "") for p in PdfReader(io.BytesIO(data)).pages]
    except ImportError:
        raise RuntimeError("Installe pdfplumber :  pip install pdfplumber")


def _norm(texte: str) -> str:
    for ch in ("\u00a0", "\u202f", "\u2009"):
        texte = texte.replace(ch, " ")
    return texte


def _entier(s) -> int:
    chiffres = re.sub(r"\D", "", str(s or ""))
    return int(chiffres) if chiffres else 0


_NUM = r"(\d[\d ]*\d|\d)"
_TAXE = r"(?:\[[A-Z0-9\-]+\])?\s*(?:\d+(?:[.,]\d+)?\s*%)?\s*"


def _somme(motif: str, texte: str):
    """Somme de tous les montants qui suivent un libellé (None si libellé absent)."""
    trouves = re.findall(motif, texte, flags=re.I)
    if not trouves:
        return None
    return sum(_entier(t) for t in trouves)


def _nettoyer_client(nom: str) -> str:
    nom = re.sub(r"\s+", " ", nom).strip()
    nom = re.sub(r"\(\s*(.*?)\s*\)", r"(\1)", nom)
    return nom


def _date_facture(t: str):
    m = re.search(r"MECeF\s+Date\s+Heure\s*:\s*(\d{2})/(\d{2})/(\d{4})", t, re.I)
    if m:
        j, mo, a = map(int, m.groups())
        return date(a, mo, j)
    m = re.search(r"Cotonou,?\s+le\s+(\d{1,2})\s+([A-Za-zéûÉ]+)\s+(\d{4})", t, re.I)
    if m and m.group(2).lower() in MOIS:
        return date(int(m.group(3)), MOIS[m.group(2).lower()], int(m.group(1)))
    return None


def analyser_facture(t: str, source: str = "") -> dict:
    """Extrait les champs d'UNE facture MECeF à partir de son texte."""
    t = _norm(t)
    ref = (re.search(r"R[ée]f\.?\s*Fact\s*:\s*(\d+)", t, re.I) or [None, ""])[1]
    mecef = (re.search(r"N[°º]\s*:\s*([A-Z]{1,3}\d+-\d+)", t) or [None, ""])[1]
    m = re.search(r"Client\s*\n\s*Nom\s*:\s*(.+)", t) or re.search(r"\bNom\s*:\s*(.+)", t)
    client = _nettoyer_client(m.group(1)) if m else ""

    ht_b = _somme(r"TOT\s*AL\s+H\.?\s*T\.?\s*" + _TAXE + _NUM, t)
    exo = _somme(r"TOT\s*AL\s+EXON[EÉ]R[EÉ]E?S?\s*" + _TAXE + _NUM, t)
    tva = _somme(r"TOT\s*AL\s+T\s*V\s*A\s*" + _TAXE + _NUM, t) or 0
    aib = _somme(r"TOT\s*AL\s+AIB\s*" + _TAXE + _NUM, t) or 0

    m_tot = re.search(r"\(\s*" + _NUM + r"\s*\)\s*Francs", t, re.I)
    total = _entier(m_tot.group(1)) if m_tot else None

    ht = (ht_b or 0) + (exo or 0)
    remarques = []
    if ht_b is None and exo is None:
        if total is not None:
            ht = total - tva - aib
            remarques.append("HT déduit du total")
        else:
            remarques.append("HT introuvable")
    if total is None:
        total = ht + tva + aib
        remarques.append("Total recalculé")
    elif abs(total - (ht + tva + aib)) > TOLERANCE:
        remarques.append(f"Écart {total - (ht + tva + aib):+,} FCFA".replace(",", " "))
    if not ref:
        remarques.append("Réf Fact introuvable")
    if not client:
        remarques.append("Client introuvable")
    d = _date_facture(t)
    if d is None:
        remarques.append("Date introuvable")

    numero = mecef if (NUM_FACT_SOURCE == "mecef" and mecef) else ref
    return {
        "Date": d, "N° Fact": numero, "Nom Client": client,
        "Montant HT": ht, "TVA": tva, "AIB": aib, "Total": total,
        "N° MECeF": mecef, "Contrôle": "OK" if not remarques else " · ".join(remarques),
        "Source": source,
    }


def analyser_pdf(data: bytes, nom_fichier: str = "") -> list:
    """Découpe le PDF en factures (une nouvelle facture commence à « Réf Fact »)."""
    blocs = []
    for i, texte in enumerate(lire_pages_pdf(data), start=1):
        if re.search(r"R[ée]f\.?\s*Fact", texte, re.I) or not blocs:
            blocs.append([i, texte])
        else:
            blocs[-1][1] += "\n" + texte          # suite d'une facture sur plusieurs pages
    factures = []
    for page, texte in blocs:
        if not texte.strip():
            continue
        f = analyser_facture(texte, f"{nom_fichier} p.{page}")
        if f["N° Fact"] or f["Montant HT"]:
            factures.append(f)
    return factures


# ====================== EXCEL ======================
def _cle(ref) -> str:
    if isinstance(ref, float) and ref.is_integer():
        ref = int(ref)
    return str(ref or "").strip().lstrip("0").upper()


def _creer_classeur(chemin: Path):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Feuil1"
    ws.append(COLS)
    ws.freeze_panes = "A2"
    for col, w in zip("ABCDEFG", (12, 14, 30, 16, 14, 12, 16)):
        ws.column_dimensions[col].width = w
    wb.save(chemin)


def _feuille(wb):
    return wb["Feuil1"] if "Feuil1" in wb.sheetnames else wb.worksheets[0]


def _ligne_vide(ws, r: int) -> bool:
    return all(ws.cell(r, c).value in (None, "") for c in range(1, 5))


def refs_existantes(chemin) -> set:
    chemin = Path(chemin)
    if not chemin.exists():
        return set()
    from openpyxl import load_workbook
    ws = _feuille(load_workbook(chemin, data_only=True))
    return {_cle(ws.cell(r, 2).value) for r in range(2, ws.max_row + 1) if ws.cell(r, 2).value not in (None, "")}


def ecrire_lignes(chemin, lignes: list) -> int:
    """Écrit les lignes (valeurs, pas de formules) dans les premières lignes libres."""
    chemin = Path(chemin)
    if not chemin.exists():
        _creer_classeur(chemin)
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    wb = load_workbook(chemin)
    ws = _feuille(wb)
    for col, w in zip("ABCDEFG", (12, 14, 30, 16, 14, 12, 16)):   # largeurs seulement si non définies
        if col not in ws.column_dimensions:
            ws.column_dimensions[col].width = w
    r = 2
    for l in lignes:
        while r <= ws.max_row and not _ligne_vide(ws, r):
            r += 1
        modele = r - 1 if r > ws.max_row and r > 2 else r
        valeurs = [l["Date"], str(l["N° Fact"]), l["Nom Client"],
                   int(l["Montant HT"]), int(l["TVA"]), int(l["AIB"]), int(l["Total"])]
        for c, v in enumerate(valeurs, start=1):
            cell = ws.cell(r, c, v)
            cell.font = Font(name="Calibri", size=11)
            if c == 1:
                cell.number_format = FORMAT_DATE
            elif c == 2:
                cell.number_format = "@"
            elif c >= 4:
                ref_fmt = ws.cell(modele, c).number_format
                cell.number_format = ref_fmt if ref_fmt != "General" else FORMAT_MONTANT
        r += 1
    wb.save(chemin)
    return len(lignes)


def lire_journal(chemin) -> pd.DataFrame:
    chemin = Path(chemin)
    if not chemin.exists():
        return pd.DataFrame(columns=COLS)
    df = pd.read_excel(chemin, sheet_name=0, converters={1: lambda v: "" if pd.isna(v) else str(v).strip()})
    df = df.iloc[:, :7]
    df.columns = COLS[:df.shape[1]]
    for c in COLS:
        if c not in df.columns:
            df[c] = None
    df = df.dropna(how="all", subset=COLS[:4]).copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    for c in COLS[3:]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype("int64")
    df["N° Fact"] = df["N° Fact"].astype(str).str.replace(r"\.0$", "", regex=True)
    return df.reset_index(drop=True)


def calcul_taxes(ht: int, avec_tva: bool, avec_aib: bool):
    tva = round(ht * TAUX_TVA) if avec_tva else 0
    aib = round(ht * TAUX_AIB) if avec_aib else 0
    return tva, aib, ht + tva + aib


def fmt(n) -> str:
    return f"{int(n):,}".replace(",", " ")



def modifier_ligne(chemin, ref_origine, l: dict) -> bool:
    """Remplace les valeurs de la ligne dont le N° Fact est ref_origine."""
    from openpyxl import load_workbook
    wb = load_workbook(chemin)
    ws = _feuille(wb)
    cle = _cle(ref_origine)
    for r in range(2, ws.max_row + 1):
        if _cle(ws.cell(r, 2).value) == cle:
            valeurs = [l["Date"], str(l["N° Fact"]), l["Nom Client"],
                       int(l["Montant HT"]), int(l["TVA"]), int(l["AIB"]), int(l["Total"])]
            for c, v in enumerate(valeurs, start=1):
                cell = ws.cell(r, c)
                cell.value = v
                if c == 1:
                    cell.number_format = FORMAT_DATE
                elif c == 2:
                    cell.number_format = "@"
                elif c >= 4 and cell.number_format == "General":
                    cell.number_format = FORMAT_MONTANT
            wb.save(chemin)
            return True
    return False


def construire_releve_html(df: pd.DataFrame, du, au, client: str) -> str:
    """Relevé des factures en HTML (format A4) avec bouton Imprimer."""
    import html as H
    d = df.sort_values(["Date", "N° Fact"])
    lignes = "".join(
        f"<tr><td nowrap>{r.Date:%d/%m/%Y}</td><td nowrap>{H.escape(str(r.ref))}</td>"
        f"<td class='l'>{H.escape(str(r.client))}</td>"
        f"<td class='n' nowrap>{fmt(r.ht)}</td><td class='n' nowrap>{fmt(r.TVA)}</td>"
        f"<td class='n' nowrap>{fmt(r.AIB)}</td><td class='n' nowrap><b>{fmt(r.Total)}</b></td></tr>"
        for r in d.rename(columns={"N° Fact": "ref", "Nom Client": "client", "Montant HT": "ht"}).itertuples()
    )
    tot = d[["Montant HT", "TVA", "AIB", "Total"]].sum()
    return f"""<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8">
<style>
 @page {{ size: A4; margin: 12mm; }}
 body {{ font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: #111; margin: 12px; }}
 .entete {{ text-align: center; margin-bottom: 10px; }}
 h1 {{ font-size: 16px; margin: 0; }}
 .entete .sous {{ font-size: 10px; color: #444; margin-top: 2px; }}
 h2 {{ text-align: center; font-size: 15px; margin: 14px 0 4px; letter-spacing: 1px; }}
 .infos {{ text-align: center; font-size: 11px; margin-bottom: 10px; }}
 table {{ width: 100%; border-collapse: collapse; }}
 th, td {{ border: 1px solid #555; padding: 4px 6px; white-space: nowrap; }}
 th {{ background: #d9d9d9; font-size: 11px; }}
 thead {{ display: table-header-group; }}
 tr {{ page-break-inside: avoid; }}
 td.n {{ text-align: right; white-space: nowrap; }}
 td.l {{ text-align: left; white-space: normal; }}
 tfoot td {{ background: #efefef; font-weight: bold; }}
 .pied {{ margin-top: 10px; font-size: 11px; }}
 .bouton {{ text-align: right; margin-bottom: 8px; }}
 .bouton button {{ padding: 7px 16px; font-size: 13px; cursor: pointer; }}
 @media print {{ .bouton {{ display: none; }} body {{ margin: 0; }} }}
</style></head><body>
<div class="bouton"><button onclick="window.print()">🖨️ Imprimer</button></div>
<div class="entete"><h1>MITHA INTERNATIONAL TRADING COMPANY</h1>
<div class="sous">Tokpa · IFU : 3201200447119 · RCCM : RB/COT/12 B 8543 · Tél : 97974597</div></div>
<h2>RELEVÉ DES FACTURES</h2>
<div class="infos">Période : du {du:%d/%m/%Y} au {au:%d/%m/%Y} · Client : {H.escape(client)} · Édité le {date.today():%d/%m/%Y}</div>
<table border="1" cellspacing="0" cellpadding="4"><thead><tr><th>Date</th><th>N° Fact</th><th>Nom Client</th><th>Montant HT</th><th>TVA</th><th>AIB</th><th>Total TTC</th></tr></thead>
<tbody>{lignes}</tbody>
<tfoot><tr><td colspan="3">TOTAL ({len(d)} facture{'s' if len(d) > 1 else ''})</td>
<td class="n">{fmt(tot['Montant HT'])}</td><td class="n">{fmt(tot['TVA'])}</td>
<td class="n">{fmt(tot['AIB'])}</td><td class="n">{fmt(tot['Total'])}</td></tr></tfoot></table>
<div class="pied">Montants en francs CFA (FCFA).</div>
</body></html>"""



def construire_releve_pdf(df: pd.DataFrame, du, au, client: str) -> bytes:
    """Relevé des factures en PDF A4 (en-tête MITC, tableau, totaux, pagination)."""
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
        from reportlab.lib.enums import TA_CENTER
    except ImportError:
        raise RuntimeError("Installe reportlab :  pip install reportlab")
    from xml.sax.saxutils import escape

    d = df.sort_values(["Date", "N° Fact"])
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=12 * mm, bottomMargin=16 * mm,
                            title="Relevé des factures MITC", author="MITHA INTERNATIONAL TRADING COMPANY")
    nom = ParagraphStyle("nom", fontName="Helvetica-Bold", fontSize=15, alignment=TA_CENTER, leading=18)
    coord = ParagraphStyle("coord", fontName="Helvetica", fontSize=8.5, alignment=TA_CENTER,
                           leading=11, textColor=colors.HexColor("#444444"))
    titre = ParagraphStyle("titre", fontName="Helvetica-Bold", fontSize=13, alignment=TA_CENTER, leading=16)
    infos = ParagraphStyle("infos", fontName="Helvetica", fontSize=9, alignment=TA_CENTER, leading=12)
    cell = ParagraphStyle("cell", fontName="Helvetica", fontSize=8.5, leading=10)

    elems = [
        Paragraph("MITHA INTERNATIONAL TRADING COMPANY", nom),
        Paragraph("Tokpa, Cotonou (Bénin)", coord),
        Paragraph("IFU : 3201200447119 &nbsp;·&nbsp; RCCM : RB/COT/12 B 8543 &nbsp;·&nbsp; Tél : 97974597", coord),
        Spacer(1, 7 * mm),
        Paragraph("RELEVÉ DES FACTURES", titre),
        Spacer(1, 2 * mm),
        Paragraph(f"Période : du {du:%d/%m/%Y} au {au:%d/%m/%Y} &nbsp;·&nbsp; Client : {escape(client)} "
                  f"&nbsp;·&nbsp; Édité le {date.today():%d/%m/%Y}", infos),
        Spacer(1, 5 * mm),
    ]

    data = [["Date", "N° Fact", "Nom Client", "Montant HT", "TVA", "AIB", "Total TTC"]]
    for r in d.itertuples(index=False):
        data.append([f"{r[0]:%d/%m/%Y}", str(r[1]), Paragraph(escape(str(r[2])), cell),
                     fmt(r[3]), fmt(r[4]), fmt(r[5]), fmt(r[6])])
    tot = d[["Montant HT", "TVA", "AIB", "Total"]].sum()
    data.append([f"TOTAL ({len(d)} facture{'s' if len(d) > 1 else ''})", "", "",
                 fmt(tot["Montant HT"]), fmt(tot["TVA"]), fmt(tot["AIB"]), fmt(tot["Total"])])

    largeurs = [21 * mm, 20 * mm, 46 * mm, 29 * mm, 22 * mm, 17 * mm, 29 * mm]
    n = len(data)
    t = Table(data, colWidths=largeurs, repeatRows=1)
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Helvetica", 8.5),
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8.5),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9d9d9")),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (3, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONT", (6, 1), (6, n - 2), "Helvetica-Bold", 8.5),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#555555")),
        ("BACKGROUND", (0, n - 1), (-1, n - 1), colors.HexColor("#efefef")),
        ("FONT", (0, n - 1), (-1, n - 1), "Helvetica-Bold", 8.5),
        ("SPAN", (0, n - 1), (2, n - 1)),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]))
    elems += [t, Spacer(1, 4 * mm), Paragraph("Montants en francs CFA (FCFA).", coord)]

    def pied(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#555555"))
        canvas.drawString(12 * mm, 9 * mm, "MITHA INTERNATIONAL TRADING COMPANY — Relevé des factures")
        canvas.drawRightString(A4[0] - 12 * mm, 9 * mm, f"Page {doc_.page}")
        canvas.restoreState()

    doc.build(elems, onFirstPage=pied, onLaterPages=pied)
    return buf.getvalue()


# ====================== INTERFACE ======================
if __name__ != "__main__" or True:   # (toujours exécuté par Streamlit)
    import streamlit as st

    st.set_page_config(page_title="Factures MITC", page_icon="🧾", layout="wide")
    ss = st.session_state

    st.title("🧾 Gestion des factures — MITC")
    st.caption("MITHA INTERNATIONAL TRADING COMPANY · factures MECeF/DGI → fichier Excel")

    with st.expander("⚙️ Fichier Excel utilisé"):
        chemin = st.text_input("Chemin du fichier", value=str(FICHIER))
    chemin = Path(chemin)
    st.caption(f"Fichier : `{chemin.name}`" + ("" if chemin.exists() else " — sera créé au premier enregistrement"))

    flash = ss.pop("fac_flash", None)
    if flash:
        getattr(st, flash[0])(flash[1])

    # ---------- callback saisie manuelle ----------
    def _montants_manuels():
        ht = int(ss.get("m_ht", 0) or 0)
        tva_a, aib_a, _ = calcul_taxes(ht, ss.get("m_tva_on", True), ss.get("m_aib_on", True))
        if ss.get("m_corriger"):
            tva = int(ss.get(f"m_tva_{tva_a}", tva_a) or 0)
            aib = int(ss.get(f"m_aib_{aib_a}", aib_a) or 0)
        else:
            tva, aib = tva_a, aib_a
        return ht, tva, aib, ht + tva + aib, tva_a, aib_a

    def _cb_ajouter_manuel():
        ht, tva, aib, total, _, _ = _montants_manuels()
        ref = str(ss.get("m_ref", "")).strip()
        client = re.sub(r"\s+", " ", str(ss.get("m_client", ""))).strip()
        if not ref or not client or ht <= 0:
            ss["fac_flash"] = ("error", "Renseigne le N° Fact, le nom du client et un Montant HT supérieur à 0.")
            return
        if _cle(ref) in refs_existantes(chemin):
            ss["fac_flash"] = ("error", f"La facture {ref} existe déjà dans l'Excel.")
            return
        try:
            ecrire_lignes(chemin, [{"Date": ss["m_date"], "N° Fact": ref, "Nom Client": client,
                                    "Montant HT": ht, "TVA": tva, "AIB": aib, "Total": total}])
        except PermissionError:
            ss["fac_flash"] = ("error", "Impossible d'écrire : ferme le fichier Excel puis réessaie.")
            return
        ss["fac_flash"] = ("success", f"Facture {ref} ({client}) ajoutée — total {fmt(total)} FCFA.")
        ss["m_ref"], ss["m_client"], ss["m_ht"] = "", "", 0

    # ---------- import + saisie manuelle côte à côte ----------
    col_imp, col_man = st.columns([3, 2], gap="large")

    with col_imp:
        st.subheader("📥 Importer des factures (PDF)")
        fichiers = st.file_uploader(
            "Dépose le « Jeu de facture » (une facture par page) — ou plusieurs PDF",
            type=["pdf"], accept_multiple_files=True, key="fac_upload")
        if fichiers:
            sig = tuple((f.name, f.size) for f in fichiers)
            if ss.get("fac_sig") != sig:
                lignes, erreurs = [], []
                for f in fichiers:
                    try:
                        lignes += analyser_pdf(f.getvalue(), f.name)
                    except Exception as e:                      # noqa
                        erreurs.append(f"{f.name} : {e}")
                if not lignes and not erreurs:
                    erreurs.append("Aucun texte lisible : PDF scanné ? Utilise la saisie manuelle.")
                dfp = pd.DataFrame(lignes)
                if not dfp.empty:
                    deja = refs_existantes(chemin)
                    dfp.insert(0, "Importer", [(_cle(r) not in deja and c == "OK")
                                               for r, c in zip(dfp["N° Fact"], dfp["Contrôle"])])
                    dfp.loc[[_cle(r) in deja for r in dfp["N° Fact"]], "Contrôle"] = "Déjà dans l'Excel"
                    dfp["Date"] = pd.to_datetime(dfp["Date"], errors="coerce")
                    dfp = dfp.sort_values("N° Fact").reset_index(drop=True)
                ss["fac_sig"], ss["fac_df"], ss["fac_err"] = sig, dfp, erreurs
                ss["fac_v"] = ss.get("fac_v", 0) + 1
            for e in ss.get("fac_err", []):
                st.error(e)
        else:
            ss.pop("fac_sig", None)
            ss.pop("fac_df", None)
            st.info("Dépose ton PDF : les factures sont lues automatiquement, tu vérifies, puis tu enregistres.")

    with col_man:
        st.subheader("✍️ Saisie manuelle")
        st.date_input("Date", value=date.today(), format="DD/MM/YYYY", key="m_date")
        st.text_input("N° Fact", placeholder="ex. 0002479", key="m_ref")
        st.text_input("Nom client", placeholder="ex. OKCHOUKOU", key="m_client")
        st.number_input("Montant HT (FCFA)", min_value=0, step=1000, key="m_ht")
        c1, c2 = st.columns(2)
        c1.checkbox("TVA 18 %", value=True, key="m_tva_on")
        c2.checkbox("AIB 1 %", value=True, key="m_aib_on")
        st.checkbox("Corriger TVA / AIB à la main", key="m_corriger")
        _, _, _, _, tva_a, aib_a = _montants_manuels()
        if ss.get("m_corriger"):
            d1, d2 = st.columns(2)
            d1.number_input("TVA", min_value=0, value=tva_a, step=1, key=f"m_tva_{tva_a}")
            d2.number_input("AIB", min_value=0, value=aib_a, step=1, key=f"m_aib_{aib_a}")
        ht_m, tva_m, aib_m, total_m, _, _ = _montants_manuels()
        st.markdown(f"**TVA** {fmt(tva_m)} · **AIB** {fmt(aib_m)} · **Total TTC {fmt(total_m)} FCFA**")
        st.button("➕ Ajouter à l'Excel", on_click=_cb_ajouter_manuel, type="primary")

    # ---------- vérification avant enregistrement ----------
    dfp = ss.get("fac_df")
    if dfp is not None and not dfp.empty:
        st.divider()
        st.subheader(f"🔎 Vérifier les {len(dfp)} facture(s) lues")
        st.caption("Corrige une cellule si besoin. Décoche « Importer » pour ignorer une facture. "
                   "Pas de TVA ou pas d'AIB sur la facture → mets 0.")
        edite = st.data_editor(
            dfp, key=f"fac_editor_{ss.get('fac_v', 0)}", hide_index=True, width="stretch",
            num_rows="fixed",
            disabled=["Contrôle", "N° MECeF", "Source"],
            column_config={
                "Importer": st.column_config.CheckboxColumn("Importer", width="small"),
                "Date": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
                "Montant HT": st.column_config.NumberColumn("Montant HT", step=1, format="%d"),
                "TVA": st.column_config.NumberColumn("TVA", step=1, format="%d"),
                "AIB": st.column_config.NumberColumn("AIB", step=1, format="%d"),
                "Total": st.column_config.NumberColumn("Total", step=1, format="%d"),
            })
        sel = edite[edite["Importer"] == True]                      # noqa: E712
        ecart = (sel["Total"] - (sel["Montant HT"] + sel["TVA"] + sel["AIB"])).abs()
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Factures à importer", len(sel))
        k2.metric("Total HT", fmt(sel["Montant HT"].sum()))
        k3.metric("Total TTC", fmt(sel["Total"].sum()))
        k4.metric("À vérifier (écart > 2 FCFA)", int((ecart > TOLERANCE).sum()))

        if st.button("💾 Enregistrer dans l'Excel", type="primary", disabled=sel.empty):
            deja = refs_existantes(chemin)
            a_ecrire, doublons, incomplets = [], [], []
            for _, r in sel.iterrows():
                if pd.isna(r["Date"]) or not str(r["N° Fact"]).strip() or not str(r["Nom Client"]).strip():
                    incomplets.append(str(r["N° Fact"]) or "?")
                elif _cle(r["N° Fact"]) in deja:
                    doublons.append(str(r["N° Fact"]))
                else:
                    a_ecrire.append({**r.to_dict(), "Date": r["Date"].date()})
            if incomplets:
                st.error("Date, N° Fact ou client manquant : " + ", ".join(incomplets))
            else:
                try:
                    n = ecrire_lignes(chemin, a_ecrire) if a_ecrire else 0
                except PermissionError:
                    st.error("Impossible d'écrire : ferme le fichier Excel puis réessaie.")
                else:
                    msg = f"{n} facture(s) enregistrée(s)."
                    if doublons:
                        msg += " Ignorées (déjà présentes) : " + ", ".join(doublons)
                    ss["fac_flash"] = ("success", msg)
                    ss["fac_df"] = None
                    st.rerun()

    # ---------- journal ----------
    st.divider()
    st.subheader("📋 Journal des factures")
    dfj_all = lire_journal(chemin)
    if dfj_all.empty:
        st.info("Aucune facture enregistrée pour l'instant.")
    else:
        dfj = dfj_all
        q = st.text_input("Recherche (client, N° Fact…)", key="fac_q")
        if q.strip():
            texte = dfj.fillna("").astype(str).agg(" ".join, axis=1).str.lower()
            dfj = dfj[texte.str.contains(q.strip().lower(), regex=False, na=False)]
        k = st.columns(5)
        k[0].metric("Factures", len(dfj))
        k[1].metric("Total HT", fmt(dfj["Montant HT"].sum()))
        k[2].metric("TVA", fmt(dfj["TVA"].sum()))
        k[3].metric("AIB", fmt(dfj["AIB"].sum()))
        k[4].metric("Total TTC", fmt(dfj["Total"].sum()))
        aff = dfj.sort_values(["Date", "N° Fact"], ascending=False).copy()
        aff["Date"] = aff["Date"].dt.strftime("%d/%m/%Y")
        st.dataframe(aff, width="stretch", hide_index=True, height=420)
        if chemin.exists():
            st.download_button("⬇️ Télécharger l'Excel", data=chemin.read_bytes(), file_name=chemin.name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # ---------- modification d'une facture ----------
        st.divider()
        st.subheader("✏️ Modifier une facture")
        tri = dfj_all.sort_values(["Date", "N° Fact"], ascending=False)
        lignes_par_cle = {}
        for _, r in tri.iterrows():
            lignes_par_cle.setdefault(_cle(r["N° Fact"]), r)
        sel = st.selectbox(
            "Choisis la facture à modifier", list(lignes_par_cle.keys()),
            format_func=lambda c: (f"{lignes_par_cle[c]['N° Fact']} — {lignes_par_cle[c]['Nom Client']} — "
                                   f"{lignes_par_cle[c]['Date']:%d/%m/%Y} — {fmt(lignes_par_cle[c]['Total'])} FCFA"),
            key="mod_sel")
        lg = lignes_par_cle[sel]

        def _cb_modifier(cle_sel, ref_origine):
            g = lambda n: ss.get(f"e_{n}_{cle_sel}")
            ref = str(g("ref") or "").strip()
            client = re.sub(r"\s+", " ", str(g("client") or "")).strip()
            if not ref or not client:
                ss["fac_flash"] = ("error", "Le N° Fact et le nom du client ne peuvent pas être vides.")
                return
            if _cle(ref) != cle_sel and _cle(ref) in refs_existantes(chemin):
                ss["fac_flash"] = ("error", f"Le N° Fact {ref} existe déjà dans l'Excel.")
                return
            try:
                ok = modifier_ligne(chemin, ref_origine, {
                    "Date": g("date"), "N° Fact": ref, "Nom Client": client,
                    "Montant HT": g("ht"), "TVA": g("tva"), "AIB": g("aib"), "Total": g("total")})
            except PermissionError:
                ss["fac_flash"] = ("error", "Impossible d'écrire : ferme le fichier Excel puis réessaie.")
                return
            ss["fac_flash"] = ("success", f"Facture {ref} modifiée.") if ok else \
                              ("error", "Facture introuvable dans l'Excel.")

        e1, e2 = st.columns(2)
        e1.date_input("Date", value=lg["Date"].date() if pd.notna(lg["Date"]) else date.today(),
                      format="DD/MM/YYYY", key=f"e_date_{sel}")
        e2.text_input("N° Fact", value=str(lg["N° Fact"]), key=f"e_ref_{sel}")
        st.text_input("Nom client", value=str(lg["Nom Client"]), key=f"e_client_{sel}")
        f1, f2, f3, f4 = st.columns(4)
        ht_e = f1.number_input("Montant HT", min_value=0, step=1, value=int(lg["Montant HT"]), key=f"e_ht_{sel}")
        tva_e = f2.number_input("TVA", min_value=0, step=1, value=int(lg["TVA"]), key=f"e_tva_{sel}")
        aib_e = f3.number_input("AIB", min_value=0, step=1, value=int(lg["AIB"]), key=f"e_aib_{sel}")
        tot_e = f4.number_input("Total", min_value=0, step=1, value=int(lg["Total"]), key=f"e_total_{sel}")
        ecart_e = tot_e - (ht_e + tva_e + aib_e)
        if abs(ecart_e) > TOLERANCE:
            st.warning(f"Le total diffère de HT + TVA + AIB de {fmt(abs(ecart_e))} FCFA "
                       f"(HT + TVA + AIB = {fmt(ht_e + tva_e + aib_e)}).")
        st.button("💾 Enregistrer la modification", on_click=_cb_modifier,
                  args=(sel, str(lg["N° Fact"])), key="btn_modifier")

        # ---------- relevé à imprimer ----------
        st.divider()
        st.subheader("🖨️ Relevé des factures à imprimer")
        d_min, d_max = dfj_all["Date"].min().date(), dfj_all["Date"].max().date()
        r1, r2, r3 = st.columns(3)
        du = r1.date_input("Du", value=d_min, format="DD/MM/YYYY")
        au = r2.date_input("Au", value=d_max, format="DD/MM/YYYY")
        clients = ["Tous"] + sorted(dfj_all["Nom Client"].dropna().astype(str).unique())
        cl = r3.selectbox("Client", clients)
        rel = dfj_all[(dfj_all["Date"].dt.date >= du) & (dfj_all["Date"].dt.date <= au)]
        if cl != "Tous":
            rel = rel[rel["Nom Client"].astype(str) == cl]
        if rel.empty:
            st.info("Aucune facture sur cette période / pour ce client.")
        else:
            html_releve = construire_releve_html(rel, du, au, cl)
            st.caption("Clique sur « Imprimer » dans l'aperçu ci-dessous. "
                       "Pour un PDF : choisis « Enregistrer au format PDF » comme imprimante.")
            try:
                pdf_releve = construire_releve_pdf(rel, du, au, cl)
                st.download_button("📄 Télécharger le relevé en PDF", data=pdf_releve, type="primary",
                                   file_name=f"Releve_factures_MITC_{du:%Y%m%d}_{au:%Y%m%d}.pdf",
                                   mime="application/pdf")
            except RuntimeError as e:
                st.warning(str(e))
            import streamlit.components.v1 as components
            components.html(html_releve, height=min(1100, 360 + 30 * len(rel)), scrolling=True)
            st.download_button("⬇️ Télécharger le relevé (HTML)", data=html_releve.encode("utf-8"),
                               file_name=f"Releve_factures_MITC_{du:%Y%m%d}_{au:%Y%m%d}.html",
                               mime="text/html")
