# -*- coding: utf-8 -*-
"""
STE NAP SARL - Génération des bons de commande (format GDIZ : en-tête blanc encadré bleu ciel, corps noir/gris).

- Lit donnees_bons.xlsx (même fichier que l'application Streamlit)
- Produit un PDF d'une page A4 par bon, dans le dossier bons_generes/
- Garde la même fonction que l'ancienne version : generer_tous_les_bons(forcer_tout=False)
  -> app_bons.py continue de fonctionner sans changement d'appel.

Fichiers à placer à la racine du projet (à côté de app_bons.py) :
    LOGO_NAP_SARL.jpeg     (logo NAP)
    logo_gdiz.png          (logo GDIZ)
"""
import hashlib
import hmac
import os
import random
import re
import string
import sys
from datetime import datetime, timedelta

from openpyxl import load_workbook
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Image, KeepInFrame, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

# ==================== CONFIGURATION ====================
FICHIER_EXCEL = "donnees_bons.xlsx"
DOSSIER_SORTIE = "bons_generes"   # relatif au dossier de lancement de l'application
SUFFIXE_REFERENCE = "-BC-SOYABEAN/NAP-2026"
CAMPAGNE = "2026-2027"
VALIDITE_JOURS = 5
# Clé secrète pour le code de sécurité / l'empreinte : changez-la (et gardez-la privée)
SEL_SECURITE = "nap_2026_changez_moi"

LOGO_NAP = "LOGO_NAP_SARL.jpeg"
LOGO_GDIZ = "logo_gdiz.png"
# Signature numérique Direction (cachet + paraphe) — optionnelle à la génération
SIGNATURE_DIRECTION = "signature_direction_nap.png"
SIGNATURE_DIRECTION_JPG = "signature_direction_nap.jpg"

# Libellé affiché + critères de qualité (exigence max, en %) par produit
PRODUITS = {
    "SHEA": {
        "titre": "KARITÉ",
        "designation": "KARITÉ (SHEA NUT)",
        "qualite": [("Humidité (Moisture)", "8 %"), ("Acidité (FFA)", "5 %")],
    },
    "SOYABEAN": {
        "titre": "SOJA",
        "designation": "SOJA GRAIN",
        "qualite": [("Humidité (Moisture)", "10 %"),
                    ("Grains défectueux (Damaged)", "2 %"),
                    ("Corps étrangers (OFM)", "2 %")],
    },
    # À AJUSTER : critères coton non fournis, repris du soja par défaut
    "COTTON": {
        "titre": "COTON",
        "designation": "GRAINES DE COTON (COTTON SEED)",
        "qualite": [("Humidité (Moisture)", "10 %"),
                    ("Grains défectueux (Damaged)", "2 %"),
                    ("Corps étrangers (OFM)", "2 %")],
    },
}
LIBELLES_PRODUIT = {k: v["titre"] for k, v in PRODUITS.items()}

PHRASE_QR = ("Scannez le QR code pour contrôler ce bon. Document généré électroniquement, "
             "valable avec signature et cachet de STE NAP SARL.")
NOTES_BON = [
    "Ce bon de commande n'est pas négociable.",
    "L'usage frauduleux de ce bon par un individu autre que le fournisseur qui a apposé sa signature est proscrit et punissable par les lois en vigueur.",
    "Ce bon de commande ne fait en aucun cas l'engagement de fixation de prix. Il accompagne juste la marchandise jusqu'à l'usine.",
    "Bon validé non livré = pénalité de 10 à 20 CFA/Kg imputée sur la commande précédente.",
    "Toute autorité administrative, policière ou douanière peut contacter le +229 0196269036 / 0190233797 pour des contrôles.",
]
# Côté acheteur (cadre « POUR L'ACHETEUR — STE NAP SARL ») :
#  - Nom      = utilisateur indiqué dans le formulaire du bon (colonne « Utilisateur » de donnees_bons.xlsx)
#  - Fonction = valeur ci-dessous (« Service Logistique » sur tous les bons)
ACHETEUR_FONCTION_DEFAUT = "Service Logistique"
ACHETEUR_NOM_DEFAUT = ""   # utilisé seulement si la colonne « Utilisateur » est vide
MODE_PAIEMENT = "Chèque"
DELAI_PAIEMENT = "Sous 72 h après le déchargement à l'usine"
LIEU_LIVRAISON = "Glo-Djigbé Industrial Zone (GDIZ), Tangbo-Djèvié, Bénin"
# =======================================================

for _s in (sys.stdout, sys.stderr):
    if _s and hasattr(_s, "reconfigure"):
        try:
            _s.reconfigure(encoding="utf-8")
        except Exception:
            pass

os.makedirs(DOSSIER_SORTIE, exist_ok=True)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ==================== POLICES (Arial Narrow si dispo, sinon Helvetica) ====================
_CANDIDATS_POLICES = [
    ("C:/Windows/Fonts/ARIALN.TTF", "C:/Windows/Fonts/ARIALNB.TTF", "C:/Windows/Fonts/ARIALNI.TTF"),
    ("/usr/share/fonts/truetype/liberation/LiberationSansNarrow-Regular.ttf",
     "/usr/share/fonts/truetype/liberation/LiberationSansNarrow-Bold.ttf",
     "/usr/share/fonts/truetype/liberation/LiberationSansNarrow-Italic.ttf"),
    (os.path.join(BASE_DIR, "polices", "ARIALN.TTF"),
     os.path.join(BASE_DIR, "polices", "ARIALNB.TTF"),
     os.path.join(BASE_DIR, "polices", "ARIALNI.TTF")),
]
F_N, F_B, F_I = "Helvetica", "Helvetica-Bold", "Helvetica-Oblique"
for _r, _b, _i in _CANDIDATS_POLICES:
    if all(os.path.exists(p) for p in (_r, _b, _i)):
        try:
            pdfmetrics.registerFont(TTFont("NAPN", _r))
            pdfmetrics.registerFont(TTFont("NAPN-B", _b))
            pdfmetrics.registerFont(TTFont("NAPN-I", _i))
            pdfmetrics.registerFontFamily("NAPN", normal="NAPN", bold="NAPN-B", italic="NAPN-I", boldItalic="NAPN-B")
            F_N, F_B, F_I = "NAPN", "NAPN-B", "NAPN-I"
            break
        except Exception:
            pass

# ==================== STYLE ====================
BLEU = colors.HexColor("#0B3C8C")
BLEU_FONCE = colors.HexColor("#0A2A63")
VERT = colors.HexColor("#2E8B2E")
OR = colors.HexColor("#C9A227")
ROUGE = colors.HexColor("#C00000")
GRIS = colors.HexColor("#F2F5FA")
LIGNE = colors.HexColor("#9AA9C4")

W, H = A4
MARGE = 12 * mm
LARGEUR = W - 2 * MARGE


def st(name, **kw):
    base = dict(fontName=F_N, fontSize=9, leading=11)
    base.update(kw)
    return ParagraphStyle(name, **base)


S_TXT = st("txt")
S_B = st("b", fontName=F_B)
S_ROUGE = st("rouge", fontName=F_B, textColor=ROUGE)
S_TH = st("th", fontName=F_B, textColor=colors.white)
S_SEC = st("sec", fontName=F_B, fontSize=9.5, textColor=BLEU, spaceBefore=3, spaceAfter=1)
S_CAMP = st("camp", fontName=F_I, fontSize=10.5, leading=13, textColor=BLEU, alignment=1)
S_TITRE = st("titre", fontName=F_B, fontSize=13.5, leading=16, textColor=BLEU, alignment=1)
S_ADR = st("adr", fontName=F_B, fontSize=8.5, leading=10.5, alignment=1)
S_NOTE = st("note", fontName=F_I, fontSize=8.3, leading=10)
S_BAND = st("band", fontName=F_B, fontSize=8.5, textColor=BLEU_FONCE, alignment=1)
S_FOOT = st("foot", fontSize=7.6, leading=9, textColor=colors.HexColor("#444444"))


def P(t, s=S_TXT):
    return Paragraph(t, s)


def _esc(t):
    return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ==================== OUTILS ====================
def trouver_fichier(nom):
    for base in (os.getcwd(), BASE_DIR, os.path.join(os.getcwd(), "logos"), os.path.join(BASE_DIR, "logos")):
        p = os.path.join(base, nom)
        if os.path.exists(p):
            return p
    return None


def nettoyer_valeur(val):
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    s = str(val).strip()
    return "" if s.lower() in ("none", "nan", "nat") else s


def extraire_numero(reference):
    if not reference:
        return 0
    base = os.path.basename(str(reference)).split("-")[0]
    m = re.search(r"(\d+)", base)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)", str(reference))
    return int(m.group(1)) if m else 0


def detecter_code_produit(reference_brute, produit=None):
    ref = str(reference_brute or "").upper().replace("\\", "/")
    for code in PRODUITS:
        if f"-BC-{code}" in ref:
            return code
    mapping = {"shea": "SHEA", "karité": "SHEA", "karite": "SHEA",
               "soja": "SOYABEAN", "soy": "SOYABEAN", "cotton": "COTTON", "coton": "COTTON"}
    if produit:
        p = str(produit).lower()
        for k, v in mapping.items():
            if k in p:
                return v
    return None


def formater_reference(reference_brute, produit=None):
    ref = str(reference_brute or "").strip()
    num = extraire_numero(ref)
    for code in PRODUITS:
        if f"-BC-{code}/NAP-2026" in ref.replace("\\", "/"):
            return f"{num:04d}-BC-{code}/NAP-2026" if num > 0 else ref
    code = detecter_code_produit("", produit)
    if num > 0:
        return f"{num:04d}-BC-{code}/NAP-2026" if code else f"{num:04d}{SUFFIXE_REFERENCE}"
    return ref


def assainir_nom_fichier(nom):
    if not nom:
        return ""
    n = re.sub(r'[\\/*?:"<>|]', "", str(nom)).strip()
    return re.sub(r"\s+", " ", n)


def separer_quantite(q):
    """'40 MT' -> ('40', 'TONNE') ; '12500 kg' -> ('12500', 'KG') ; '50' -> ('50', 'TONNE')"""
    q = nettoyer_valeur(q)
    if not q:
        return "", "TONNE"
    m = re.match(r"^\s*([\d][\d\s.,]*)\s*([A-Za-z]*)", q)
    if not m:
        return q, "TONNE"
    nombre, unite = m.group(1).strip(), m.group(2).upper()
    if unite in ("", "MT", "T", "TONNE", "TONNES", "TON"):
        unite = "TONNE"
    return nombre, unite


def code_securite(reference):
    rng = random.Random(f"{SEL_SECURITE}|{reference}")
    return "".join(rng.choices(string.ascii_uppercase + string.digits, k=10))


def empreinte(reference, fournisseur, quantite):
    h = hmac.new(SEL_SECURITE.encode(), f"{reference}|{fournisseur}|{quantite}".encode(), hashlib.sha256).hexdigest()
    return f"{int(h[:6], 16) % 1000000:06d}-{int(h[6:11], 16) % 100000:05d}"


def qr_drawing(texte, taille):
    w = qr.QrCodeWidget(texte)
    x0, y0, x1, y1 = w.getBounds()
    d = Drawing(taille, taille, transform=[taille / (x1 - x0), 0, 0, taille / (y1 - y0), 0, 0])
    d.add(w)
    return d


# ==================== BLOCS DU DOCUMENT (format V3 : en-tête encadré bleu ciel, corps noir/gris) ====================
NOIR = colors.HexColor("#111111")
GRIS_T = colors.HexColor("#555555")
GRIS_L = colors.HexColor("#BDBDBD")
CENDRE = colors.HexColor("#D9D9D9")      # en-têtes de tableaux
ROUGE_CLAIR = colors.HexColor("#E8605C")  # titre « Qualité à la réception »
ROUGE_CLAIR_FOND = colors.HexColor("#F7BDBA")  # en-tête du tableau qualité
BLEU_CIEL = colors.HexColor("#7EC8E3")   # cadre de l'en-tête
HDR_H = 30 * mm

S_LAB = st("lab3", fontSize=8, leading=10, fontName=F_N, textColor=GRIS_T)
S_VAL = st("val3", fontSize=10, leading=12, fontName=F_B, textColor=NOIR)
S_SEC = st("sec3", fontSize=9.5, leading=11, fontName=F_B, textColor=NOIR)
S_TH = st("th3", fontSize=8.8, leading=10.5, fontName=F_B, textColor=NOIR)
S_TD = st("td3", fontSize=9.3, leading=11.5, fontName=F_N, textColor=NOIR)
S_NOTE = st("note3", fontSize=8.2, leading=10, fontName=F_I, textColor=colors.HexColor("#333333"))
S_ALERTE = st("alerte3", fontSize=9.5, leading=12, fontName=F_B, textColor=colors.HexColor("#E10000"))
S_REF = st("ref3", fontSize=15, leading=17, fontName=F_B, textColor=NOIR)
S_QTE = st("qte3", fontSize=19, leading=21, fontName=F_B, textColor=NOIR)


def section(titre, couleur=None):
    style = S_SEC if couleur is None else ParagraphStyle("sec_c", parent=S_SEC, textColor=couleur)
    t = Table([[P(titre.upper(), style)]], colWidths=[LARGEUR])
    t.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 1, NOIR), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                           ("TOPPADDING", (0, 0), (-1, -1), 4.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]))
    return t


def kv(lignes):
    rows = []
    for ligne in lignes:
        row = []
        for lab, val in ligne:
            row += [P(lab, S_LAB), P(_esc(val) or "-", S_VAL)]
        while len(row) < 4:
            row += ["", ""]
        rows.append(row)
    t = Table(rows, colWidths=[LARGEUR * 0.17, LARGEUR * 0.33, LARGEUR * 0.17, LARGEUR * 0.33])
    t.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.4, GRIS_L), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("LEFTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 2.6),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 2.6)]))
    return t


def grille(rows, widths, fond_entete=None):
    t = Table(rows, colWidths=widths)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.6, NOIR), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("TOPPADDING", (0, 0), (-1, -1), 3.2), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.2),
                           ("LEFTPADDING", (0, 0), (-1, -1), 6), ("BACKGROUND", (0, 0), (-1, 0), fond_entete or CENDRE)]))
    return t


def generer_un_bon(d, chemin):
    """d : dictionnaire (voir generer_tous_les_bons). Écrit le PDF dans `chemin`."""
    prod = PRODUITS.get(d["code_produit"], PRODUITS["SOYABEAN"])
    e = lambda k, defaut="": d.get(k) or defaut
    s = [Spacer(1, HDR_H + 2 * mm)]

    top = Table([[[P("N° BON DE COMMANDE", S_LAB), P(_esc(e("reference")), S_REF)],
                  [P("DATE D'ÉMISSION", S_LAB), P(_esc(e("date_emission")), S_VAL)],
                  [P(f"VALIDITÉ ({VALIDITE_JOURS} JOURS)", S_LAB),
                   P(f"{_esc(e('date_emission'))} au {_esc(e('date_fin'))}", S_VAL)]]],
                colWidths=[LARGEUR * 0.46, LARGEUR * 0.22, LARGEUR * 0.32])
    top.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 1, NOIR), ("LINEAFTER", (0, 0), (-2, 0), 0.4, GRIS_L),
                             ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 3.5),
                             ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5), ("LEFTPADDING", (0, 0), (-1, -1), 8)]))
    s += [top]

    s += [section("Fournisseur / agrégateur"), kv([
        [("Nom", e("fournisseur")), ("Téléphone", e("tel_fournisseur"))],
        [("N° pièce d'identité", e("piece_identite")), ("Commune d'origine", e("commune"))],
        [("Lieu d'achat", e("lieu_achat") or e("commune")), ("Lieu de livraison", LIEU_LIVRAISON)],
    ])]

    s += [section("Commande"), Spacer(1, 3), grille([
        [P("Désignation", S_TH), P("Origine", S_TH), P("Quantité", S_TH), P("Unité", S_TH), P("Prix unitaire (FCFA)", S_TH)],
        [P(f"<b>{_esc(prod['designation'])}</b>", S_TD), P(_esc(e("commune")), S_TD), P(_esc(e("quantite_nb")), S_QTE),
         P(_esc(e("unite")), S_TD), P(_esc(e("prix", "Suivant conditions du marché")), S_TD)],
    ], [LARGEUR * 0.27, LARGEUR * 0.19, LARGEUR * 0.14, LARGEUR * 0.12, LARGEUR * 0.28])]

    s += [section("Transport et convoyage"), kv([
        [("Chauffeur", e("chauffeur")), ("Téléphone", e("tel_chauffeur"))],
        [("N° tracteur / remorque", e("tracteur")), ("N° permis", e("permis"))],
        [("Type de véhicule", e("type_vehicule")), ("Date limite livraison", e("date_fin"))],
    ])]

    q = [[P("Critère", S_TH), P("Exigence indiquée", S_TH), P("Paramètre de rejet", S_TH)]]
    for crit, mx in prod["qualite"]:
        q.append([P(f"<b>{crit}</b>", S_TD), P(f"Maximum {mx}", S_TD), P(f"supérieur à {mx}", S_TD)])
    s += [section("Qualité à la réception", ROUGE_CLAIR), Spacer(1, 3),
          grille(q, [LARGEUR * 0.40, LARGEUR * 0.30, LARGEUR * 0.30], ROUGE_CLAIR_FOND),
          Spacer(1, 2), P("NB : Les critères de qualité ci-dessus sont à confirmer avant validation du bon de commande.", S_NOTE)]

    s += [section("Modalités de paiement et de contrôle"), kv([
        [("Mode de paiement", MODE_PAIEMENT), ("Délai de paiement", DELAI_PAIEMENT)],
        [("Poids & qualité", "Déterminés à l'usine de transformation"), ("Lot non conforme", "Rejet ou rabais selon critères")],
    ])]

    s += [Spacer(1, 4)]
    for x in NOTES_BON:
        s.append(P("– " + x, S_NOTE))

    sec = Table([[qr_drawing(d["texte_qr"], 18 * mm),
                  [P("CODE DE SÉCURITÉ", S_LAB), P(_esc(e("code_securite")), S_VAL), Spacer(1, 2), P("EMPREINTE", S_LAB),
                   P(_esc(e("empreinte")), S_VAL)],
                  P(PHRASE_QR, S_ALERTE)]], colWidths=[23 * mm, 55 * mm, LARGEUR - 78 * mm])
    sec.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.6, GRIS_L), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                             ("LEFTPADDING", (0, 0), (-1, -1), 6), ("TOPPADDING", (0, 0), (-1, -1), 3),
                             ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
    s += [Spacer(1, 6), sec, section("Validation")]

    def boite(titre, nom, lab2, val2, avec_signature=False):
        sig_cell = [P("Signature et cachet :", S_LAB)]
        if avec_signature:
            chemin_sig = trouver_fichier(SIGNATURE_DIRECTION) or trouver_fichier(SIGNATURE_DIRECTION_JPG)
            if chemin_sig:
                try:
                    sig_cell.append(Image(chemin_sig, width=42 * mm, height=28 * mm, kind="proportional"))
                except Exception:
                    sig_cell.append(Spacer(1, 14 * mm))
            else:
                sig_cell.append(Spacer(1, 14 * mm))
        else:
            sig_cell.append(Spacer(1, 14 * mm))
        t = Table([[P(titre, S_TH)], [P(f"Nom : <b>{_esc(nom)}</b>", S_TD)], [P(f"{lab2} : <b>{_esc(val2)}</b>", S_TD)],
                   [P(f"Date : <b>{_esc(e('date_emission'))}</b>", S_TD)],
                   [sig_cell]], colWidths=[(LARGEUR - 6 * mm) / 2])
        t.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, NOIR), ("BACKGROUND", (0, 0), (-1, 0), CENDRE),
                               ("LINEBELOW", (0, 0), (-1, 0), 0.6, NOIR), ("LEFTPADDING", (0, 0), (-1, -1), 6),
                               ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                               ("VALIGN", (0, 4), (-1, 4), "TOP")]))
        return t

    sig_digitale = bool(d.get("signature_digitale"))
    sg = Table([[boite("POUR L'ACHETEUR — STE NAP SARL", e("acheteur_nom"), "Fonction", e("acheteur_fonction"),
                       avec_signature=sig_digitale), "",
                 boite("POUR LE FOURNISSEUR / AGRÉGATEUR", e("fournisseur"), "Téléphone", e("tel_fournisseur"),
                       avec_signature=False)]],
               colWidths=[(LARGEUR - 6 * mm) / 2, 6 * mm, (LARGEUR - 6 * mm) / 2])
    sg.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 4),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    s += [sg]

    ref, titre = d["reference"], prod["titre"]
    TOP = 11 * mm

    def cadre(canvas, _doc):
        canvas.saveState()
        # Bordure de la page : double filet (bleu ciel à l'extérieur, noir fin à l'intérieur)
        canvas.setStrokeColor(BLEU_CIEL)
        canvas.setLineWidth(2)
        canvas.roundRect(5 * mm, 5 * mm, W - 10 * mm, H - 10 * mm, 2.5 * mm, stroke=1, fill=0)
        canvas.setStrokeColor(NOIR)
        canvas.setLineWidth(0.5)
        canvas.roundRect(6.8 * mm, 6.8 * mm, W - 13.6 * mm, H - 13.6 * mm, 1.8 * mm, stroke=1, fill=0)
        # En-tête blanc encadré bleu ciel
        y0 = H - TOP - HDR_H
        canvas.setFillColor(colors.white)
        canvas.setStrokeColor(BLEU_CIEL)
        canvas.setLineWidth(2)
        canvas.roundRect(MARGE, y0, LARGEUR, HDR_H, 3 * mm, stroke=1, fill=1)
        ln = trouver_fichier(LOGO_NAP)
        if ln:
            lw = 34 * mm
            canvas.drawImage(ln, MARGE + 4 * mm, y0 + (HDR_H - lw * 745 / 1280) / 2, width=lw, height=lw * 745 / 1280)
        lg = trouver_fichier(LOGO_GDIZ)
        if lg:
            gw = 36 * mm
            canvas.drawImage(lg, MARGE + LARGEUR - gw - 4 * mm, y0 + (HDR_H - gw * 112 / 292) / 2,
                             width=gw, height=gw * 112 / 292)
        cx = W / 2
        canvas.setFillColor(BLEU_FONCE)
        canvas.setFont(F_B, 18)
        canvas.drawCentredString(cx, y0 + HDR_H - 9 * mm, "BON DE COMMANDE")
        canvas.setFillColor(NOIR)
        canvas.setFont(F_B, 10)
        canvas.drawCentredString(cx, y0 + HDR_H - 14.5 * mm, f"ACHAT DE {titre}  •  CAMPAGNE {CAMPAGNE}")
        canvas.setFillColor(GRIS_T)
        canvas.setFont(F_N, 7.4)
        canvas.drawCentredString(cx, y0 + HDR_H - 19.5 * mm, "STE NAP SARL – Plot No 14, Glo-Djigbé Economic Zone, Bénin")
        canvas.drawCentredString(cx, y0 + HDR_H - 23 * mm, "RCCM RB / COT / 21 B 30710 • IFU 3202113334014")
        canvas.drawCentredString(cx, y0 + HDR_H - 26.4 * mm, "Tél : +229 0196269036 / 0190233797")
        # Pied de page sobre
        canvas.setStrokeColor(NOIR)
        canvas.setLineWidth(0.6)
        canvas.line(MARGE, 13 * mm, W - MARGE, 13 * mm)
        canvas.setFillColor(GRIS_T)
        canvas.setFont(F_N, 7.6)
        canvas.drawString(MARGE, 9.5 * mm, "STE NAP SARL – Agro-industrie & Commerce")
        canvas.drawRightString(W - MARGE, 9.5 * mm, f"{ref}  •  Page 1/1")
        canvas.restoreState()

    doc = SimpleDocTemplate(chemin, pagesize=A4, leftMargin=MARGE, rightMargin=MARGE, topMargin=TOP,
                            bottomMargin=16 * mm, title=f"Bon de commande {ref}", author="STE NAP SARL")
    doc.build([KeepInFrame(LARGEUR, H - TOP - 16 * mm, s, mode="shrink")], onFirstPage=cadre, onLaterPages=cadre)
    return chemin


# ==================== VERSION WORD (même présentation que le PDF) ====================
import io

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor, Twips
from reportlab.graphics.barcode import qrencoder

W_FONT = "Arial Narrow"
W_NOIR, W_GRIS, W_GRIS_L = "111111", "555555", "BDBDBD"
W_CENDRE, W_CIEL, W_MARINE = "D9D9D9", "7EC8E3", "0A2A63"
W_ROUGE_CLAIR, W_ROUGE_CLAIR_FOND = "E8605C", "F7BDBA"
W_ALERTE = "E10000"

_ORD = {
    "tcPr": ["tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection",
             "tcFitText", "vAlign", "hideMark"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
              "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"],
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
            "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
            "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
    "sectPr": ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc",
               "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg",
               "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"],
}


def _ins(parent, el, kind):
    """Insère `el` dans `parent` à la bonne place (Word exige l'ordre des balises)."""
    order = _ORD[kind]
    tag = el.tag.split("}")[1]
    for old in parent.findall(qn("w:" + tag)):
        parent.remove(old)
    idx = order.index(tag)
    for child in list(parent):
        ctag = child.tag.split("}")[1]
        if ctag in order and order.index(ctag) > idx:
            child.addprevious(el)
            return
    parent.append(el)


def _mm_dxa(mm_val):
    return int(round(mm_val * 56.6929))


def _run(p, text, size=9.3, bold=False, italic=False, color=W_NOIR):
    r = p.add_run(text)
    f = r.font
    f.name = W_FONT
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = RGBColor.from_string(color)
    rf = r._element.get_or_add_rPr().find(qn("w:rFonts"))
    for a in ("cs", "eastAsia"):
        rf.set(qn("w:" + a), W_FONT)
    return r


def R(text, size=9.3, bold=False, italic=False, color=W_NOIR):
    return (text, size, bold, italic, color)


_ALIGN = {"l": WD_ALIGN_PARAGRAPH.LEFT, "c": WD_ALIGN_PARAGRAPH.CENTER, "r": WD_ALIGN_PARAGRAPH.RIGHT}


def _par(c, runs, align="l", before=0, after=0):
    """Ajoute un paragraphe à une cellule (réutilise le 1er paragraphe vide) ou au document."""
    if hasattr(c, "_tc") and len(c.paragraphs) == 1 and not c.paragraphs[0].runs:
        p = c.paragraphs[0]
    else:
        p = c.add_paragraph()
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(before), Pt(after), 1.0
    p.alignment = _ALIGN[align]
    for r in runs:
        _run(p, *r)
    if not runs:
        _run(p, "", 2)
    return p


def _borders(cell, **b):
    """b : top/left/bottom/right = (épaisseur en 1/8 pt, 'RRGGBB') ou None (aucun trait)."""
    el = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        v = b.get(side)
        e = OxmlElement("w:" + side)
        if v is None:
            e.set(qn("w:val"), "nil")
        else:
            e.set(qn("w:val"), "single")
            e.set(qn("w:sz"), str(v[0]))
            e.set(qn("w:space"), "0")
            e.set(qn("w:color"), v[1])
        el.append(e)
    _ins(cell._tc.get_or_add_tcPr(), el, "tcPr")


def _shade(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    _ins(cell._tc.get_or_add_tcPr(), shd, "tcPr")


def _new_table(doc, rows, widths_mm, marg=(40, 80, 40, 80)):
    t = doc.add_table(rows=rows, cols=len(widths_mm))
    t.autofit = False
    tblPr = t._tbl.tblPr
    w = OxmlElement("w:tblW")
    w.set(qn("w:w"), str(sum(_mm_dxa(x) for x in widths_mm)))
    w.set(qn("w:type"), "dxa")
    _ins(tblPr, w, "tblPr")
    ind = OxmlElement("w:tblInd")
    ind.set(qn("w:w"), "0")
    ind.set(qn("w:type"), "dxa")
    _ins(tblPr, ind, "tblPr")
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    _ins(tblPr, lay, "tblPr")
    mar = OxmlElement("w:tblCellMar")
    for side, val in zip(("top", "left", "bottom", "right"), marg):
        e = OxmlElement("w:" + side)
        e.set(qn("w:w"), str(val))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    _ins(tblPr, mar, "tblPr")
    for i, wm in enumerate(widths_mm):
        t.columns[i].width = Mm(wm)
        for row in t.rows:
            row.cells[i].width = Mm(wm)
            row.cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return t


def _section(doc, titre, couleur=W_NOIR):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing, pf.keep_with_next = Pt(5), Pt(2), 1.0, True
    _run(p, titre.upper(), 9.5, True, False, couleur)
    bd = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", "8"), ("space", "1"), ("color", W_NOIR)):
        b.set(qn("w:" + k), v)
    bd.append(b)
    _ins(p._p.get_or_add_pPr(), bd, "pPr")
    return p


def _kv(doc, lignes):
    t = _new_table(doc, len(lignes), [31.6, 61.4, 31.6, 61.4], marg=(45, 0, 45, 60))
    for i, ligne in enumerate(lignes):
        cells = t.rows[i].cells
        for j, (lab, val) in enumerate(ligne):
            _par(cells[2 * j], [R(lab, 8, color=W_GRIS)])
            _par(cells[2 * j + 1], [R(val or "-", 10, True)])
        for c in cells:
            _borders(c, bottom=(4, W_GRIS_L))
    return t


def _grille(doc, rows, widths, fond=W_CENDRE):
    t = _new_table(doc, len(rows), widths, marg=(60, 100, 60, 100))
    for i, row in enumerate(rows):
        for j, runs in enumerate(row):
            c = t.rows[i].cells[j]
            _par(c, runs)
            _borders(c, top=(6, W_NOIR), left=(6, W_NOIR), bottom=(6, W_NOIR), right=(6, W_NOIR))
            if i == 0:
                _shade(c, fond)
    return t


def _qr_png(texte, box=6, border=1):
    from PIL import Image
    q = qrencoder.QRCode(None, qrencoder.QRErrorCorrectLevel.M)
    q.addData(texte)
    q.make()
    n = q.getModuleCount()
    size = (n + 2 * border) * box
    img = Image.new("RGB", (size, size), "white")
    px = img.load()
    for r in range(n):
        for c in range(n):
            if q.isDark(r, c):
                for dy in range(box):
                    for dx in range(box):
                        px[(c + border) * box + dx, (r + border) * box + dy] = (0, 0, 0)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    return buf


def _spacer(doc, pts):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(pts)
    _run(p, "", 1)


def generer_un_bon_word(d, chemin):
    prod = PRODUITS.get(d["code_produit"], PRODUITS["SOYABEAN"])
    e = lambda k, defaut="": d.get(k) or defaut
    doc = Document()
    doc.core_properties.title = f"Bon de commande {e('reference')}"
    doc.core_properties.author = "STE NAP SARL"
    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = W_FONT, Pt(9.3)
    normal.paragraph_format.space_after = Pt(0)

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(12)
    sec.top_margin, sec.bottom_margin = Mm(11), Mm(19)
    sec.footer_distance = Mm(11)
    # Bordure de page (double filet bleu ciel)
    bd = OxmlElement("w:pgBorders")
    bd.set(qn("w:offsetFrom"), "page")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement("w:" + side)
        for k, v in (("val", "double"), ("sz", "12"), ("space", "18"), ("color", W_CIEL)):
            b.set(qn("w:" + k), v)
        bd.append(b)
    _ins(sec._sectPr, bd, "sectPr")

    # ---- Pied de page
    fp = sec.footer.paragraphs[0]
    for pos in (4680, 9360):   # supprime les tabulations par défaut du style « Pied de page »
        fp.paragraph_format.tab_stops.add_tab_stop(Twips(pos), WD_TAB_ALIGNMENT.CLEAR)
    fp.paragraph_format.tab_stops.add_tab_stop(Mm(186), WD_TAB_ALIGNMENT.RIGHT)
    _run(fp, "STE NAP SARL – Agro-industrie & Commerce", 7.6, color=W_GRIS)
    _run(fp, "\t" + f"{e('reference')}  •  Page 1/1", 7.6, color=W_GRIS)
    fbd = OxmlElement("w:pBdr")
    fb = OxmlElement("w:top")
    for k, v in (("val", "single"), ("sz", "4"), ("space", "3"), ("color", W_NOIR)):
        fb.set(qn("w:" + k), v)
    fbd.append(fb)
    _ins(fp._p.get_or_add_pPr(), fbd, "pPr")

    # ---- En-tête : logo NAP | titre | logo GDIZ, encadré bleu ciel
    t = _new_table(doc, 1, [42, 102, 42], marg=(60, 100, 60, 100))
    t.rows[0].height, t.rows[0].height_rule = Mm(28), WD_ROW_HEIGHT_RULE.AT_LEAST
    cl, cc, cr = t.rows[0].cells
    ciel = (18, W_CIEL)
    _borders(cl, top=ciel, left=ciel, bottom=ciel)
    _borders(cc, top=ciel, bottom=ciel)
    _borders(cr, top=ciel, right=ciel, bottom=ciel)
    for cell, chemin_logo, larg in ((cl, trouver_fichier(LOGO_NAP), 34), (cr, trouver_fichier(LOGO_GDIZ), 36)):
        p = _par(cell, [], "c")
        if chemin_logo:
            p.add_run().add_picture(chemin_logo, width=Mm(larg))
    _par(cc, [R("BON DE COMMANDE", 18, True, color=W_MARINE)], "c")
    _par(cc, [R(f"ACHAT DE {prod['titre']}  •  CAMPAGNE {CAMPAGNE}", 10, True)], "c", 1, 2)
    _par(cc, [R("STE NAP SARL – Plot No 14, Glo-Djigbé Economic Zone, Bénin", 7.4, color=W_GRIS)], "c")
    _par(cc, [R("RCCM RB / COT / 21 B 30710 • IFU 3202113334014", 7.4, color=W_GRIS)], "c")
    _par(cc, [R("Tél : +229 0196269036 / 0190233797", 7.4, color=W_GRIS)], "c")
    _spacer(doc, 6)

    # ---- Référence / dates
    t = _new_table(doc, 1, [85.6, 40.9, 59.5], marg=(70, 140, 70, 100))
    c0, c1, c2 = t.rows[0].cells
    _par(c0, [R("N° BON DE COMMANDE", 8, color=W_GRIS)])
    _par(c0, [R(e("reference"), 15, True)])
    _par(c1, [R("DATE D'ÉMISSION", 8, color=W_GRIS)])
    _par(c1, [R(e("date_emission"), 10, True)])
    _par(c2, [R(f"VALIDITÉ ({VALIDITE_JOURS} JOURS)", 8, color=W_GRIS)])
    _par(c2, [R(f"{e('date_emission')} au {e('date_fin')}", 10, True)])
    noir, gris = (8, W_NOIR), (4, W_GRIS_L)
    _borders(c0, top=noir, left=noir, bottom=noir, right=gris)
    _borders(c1, top=noir, bottom=noir, right=gris)
    _borders(c2, top=noir, bottom=noir, right=noir)

    _section(doc, "Fournisseur / agrégateur")
    _kv(doc, [
        [("Nom", e("fournisseur")), ("Téléphone", e("tel_fournisseur"))],
        [("N° pièce d'identité", e("piece_identite")), ("Commune d'origine", e("commune"))],
        [("Lieu d'achat", e("lieu_achat") or e("commune")), ("Lieu de livraison", LIEU_LIVRAISON)],
    ])

    _section(doc, "Commande")
    _spacer(doc, 3)
    _grille(doc, [
        [[R("Désignation", 8.8, True)], [R("Origine", 8.8, True)], [R("Quantité", 8.8, True)],
         [R("Unité", 8.8, True)], [R("Prix unitaire (FCFA)", 8.8, True)]],
        [[R(prod["designation"], 9.3, True)], [R(e("commune"))], [R(e("quantite_nb"), 19, True)],
         [R(e("unite"))], [R(e("prix", "Suivant conditions du marché"))]],
    ], [50.2, 35.3, 26, 22.3, 52.2])

    _section(doc, "Transport et convoyage")
    _kv(doc, [
        [("Chauffeur", e("chauffeur")), ("Téléphone", e("tel_chauffeur"))],
        [("N° tracteur / remorque", e("tracteur")), ("N° permis", e("permis"))],
        [("Type de véhicule", e("type_vehicule")), ("Date limite livraison", e("date_fin"))],
    ])

    _section(doc, "Qualité à la réception", W_ROUGE_CLAIR)
    _spacer(doc, 3)
    lignes_q = [[[R("Critère", 8.8, True)], [R("Exigence indiquée", 8.8, True)], [R("Paramètre de rejet", 8.8, True)]]]
    for crit, mx in prod["qualite"]:
        lignes_q.append([[R(crit, 9.3, True)], [R(f"Maximum {mx}")], [R(f"supérieur à {mx}")]])
    _grille(doc, lignes_q, [74.4, 55.8, 55.8], fond=W_ROUGE_CLAIR_FOND)
    _par(doc, [R("NB : Les critères de qualité ci-dessus sont à confirmer avant validation du bon de commande.",
                 8.2, italic=True)], before=2)

    _section(doc, "Modalités de paiement et de contrôle")
    _kv(doc, [
        [("Mode de paiement", MODE_PAIEMENT), ("Délai de paiement", DELAI_PAIEMENT)],
        [("Poids & qualité", "Déterminés à l'usine de transformation"), ("Lot non conforme", "Rejet ou rabais selon critères")],
    ])
    _spacer(doc, 4)
    for x in NOTES_BON:
        _par(doc, [R("– " + x, 8.2, italic=True)])

    # ---- Sécurité : QR + codes + phrase d'alerte
    _spacer(doc, 6)
    t = _new_table(doc, 1, [23, 55, 108], marg=(50, 100, 50, 100))
    cq, cm, cd = t.rows[0].cells
    _par(cq, []).runs[0].add_picture(_qr_png(d["texte_qr"]), width=Mm(18))
    _par(cm, [R("CODE DE SÉCURITÉ", 8, color=W_GRIS)])
    _par(cm, [R(e("code_securite"), 10, True)], after=2)
    _par(cm, [R("EMPREINTE", 8, color=W_GRIS)])
    _par(cm, [R(e("empreinte"), 10, True)])
    _par(cd, [R(PHRASE_QR, 9.5, True, False, W_ALERTE)])
    for c in (cq, cm, cd):
        _borders(c, top=(6, W_GRIS_L), bottom=(6, W_GRIS_L), left=(6, W_GRIS_L) if c is cq else None,
                 right=(6, W_GRIS_L) if c is cd else None)

    # ---- Validation
    _section(doc, "Validation")
    _spacer(doc, 4)
    t = _new_table(doc, 5, [90, 6, 90], marg=(45, 100, 45, 100))
    blocs = {0: ("POUR L'ACHETEUR — STE NAP SARL", e("acheteur_nom"), "Fonction", e("acheteur_fonction")),
             2: ("POUR LE FOURNISSEUR / AGRÉGATEUR", e("fournisseur"), "Téléphone", e("tel_fournisseur"))}
    t.rows[4].height, t.rows[4].height_rule = Mm(28), WD_ROW_HEIGHT_RULE.AT_LEAST
    sig_digitale = bool(d.get("signature_digitale"))
    for col, (titre, nom, lab2, val2) in blocs.items():
        contenus = [[R(titre, 8.8, True)], [R("Nom : "), R(nom, 9.3, True)], [R(f"{lab2} : "), R(val2, 9.3, True)],
                    [R("Date : "), R(e("date_emission"), 9.3, True)], [R("Signature et cachet :", 8, color=W_GRIS)]]
        for r in range(5):
            c = t.rows[r].cells[col]
            _par(c, contenus[r])
            _borders(c, top=(6, W_NOIR) if r in (0,) else None, bottom=(6, W_NOIR) if r in (0, 4) else None,
                     left=(6, W_NOIR), right=(6, W_NOIR))
            if r == 0:
                _shade(c, W_CENDRE)
            if r == 4:
                c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
                # Signature numérique uniquement côté acheteur (col 0)
                if col == 0 and sig_digitale:
                    chemin_sig = trouver_fichier(SIGNATURE_DIRECTION_JPG) or trouver_fichier(SIGNATURE_DIRECTION)
                    if chemin_sig:
                        try:
                            p_sig = c.paragraphs[-1] if c.paragraphs else c.add_paragraph()
                            p_sig.alignment = 1  # center
                            run = p_sig.add_run()
                            run.add_picture(chemin_sig, width=Mm(40))
                        except Exception:
                            pass
    for r in range(5):
        _borders(t.rows[r].cells[1])
    _par(doc, [], before=0)

    z = doc.settings.element.find(qn("w:zoom"))
    if z is not None and z.get(qn("w:percent")) is None:
        z.set(qn("w:percent"), "100")
    doc.save(chemin)
    return chemin


# ==================== GÉNÉRATION EN SÉRIE (appelée par app_bons.py) ====================
def _date_texte(val):
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    v = nettoyer_valeur(val)
    return v or datetime.now().strftime("%d/%m/%Y")


def generer_tous_les_bons(forcer_tout=False, signature_digitale=False):
    """Génère Word+PDF. signature_digitale=True → cachet Direction dans la zone acheteur."""
    os.makedirs(DOSSIER_SORTIE, exist_ok=True)
    resultat = {"generes": 0, "ignores": 0, "erreurs": [], "dossier": os.path.abspath(DOSSIER_SORTIE)}
    if not os.path.exists(FICHIER_EXCEL):
        print(f"Erreur : Le fichier {FICHIER_EXCEL} est introuvable.")
        resultat["erreurs"].append(f"Fichier {FICHIER_EXCEL} introuvable (dossier de lancement : {os.getcwd()})")
        return resultat

    ws = load_workbook(FICHIER_EXCEL).active
    headers = [c.value for c in ws[1]]
    compteur = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        data = dict(zip(headers, row))
        ref_brute = data.get("Reference")
        if not ref_brute:
            continue

        nom = nettoyer_valeur(data.get("Nom_Fournisseur"))
        reference = formater_reference(ref_brute, produit=data.get("Produit"))
        code_produit = detecter_code_produit(reference, data.get("Produit")) or "SOYABEAN"
        ref_propre = reference.replace("/", "-").replace("\\", "-").strip()
        fourn_propre = assainir_nom_fichier(nom)
        base_fichier = f"{ref_propre} - {fourn_propre}" if fourn_propre else ref_propre
        chemin_pdf = os.path.join(DOSSIER_SORTIE, base_fichier + ".pdf")
        chemin_word = os.path.join(DOSSIER_SORTIE, base_fichier + ".docx")

        # Un bon est complet seulement s'il existe en PDF ET en Word
        if os.path.exists(chemin_pdf) and os.path.exists(chemin_word) and not forcer_tout:
            resultat["ignores"] += 1
            continue

        date_emission = _date_texte(data.get("Date_Emission"))
        try:
            dt = datetime.strptime(date_emission, "%d/%m/%Y")
            date_fin = (dt + timedelta(days=VALIDITE_JOURS)).strftime("%d/%m/%Y")
        except ValueError:
            date_fin = ""

        nb, unite = separer_quantite(data.get("Quantite"))
        heure_gen = datetime.now().strftime("%H:%M:%S")
        utilisateur = nettoyer_valeur(data.get("Utilisateur")) or "Systeme"
        code = code_securite(reference)
        emp = empreinte(reference, nom, nb)

        d = {
            "reference": reference,
            "code_produit": code_produit,
            "date_emission": date_emission,
            "date_fin": date_fin,
            "fournisseur": nom,
            "tel_fournisseur": nettoyer_valeur(data.get("Telephone_Fournisseur")),
            "piece_identite": nettoyer_valeur(data.get("Piece_Identite")),
            "commune": nettoyer_valeur(data.get("Commune_Origine")),
            "lieu_achat": nettoyer_valeur(data.get("Lieu_Achat")),
            "quantite_nb": nb,
            "unite": unite,
            "prix": nettoyer_valeur(data.get("Prix_Unitaire")),
            "chauffeur": nettoyer_valeur(data.get("Chauffeur")),
            "tel_chauffeur": nettoyer_valeur(data.get("Telephone_Chauffeur")),
            "tracteur": nettoyer_valeur(data.get("Numero_Tracteur")),
            "permis": nettoyer_valeur(data.get("Numero_Permis")),
            "type_vehicule": nettoyer_valeur(data.get("Type_Vehicule")),
            "acheteur_nom": (nettoyer_valeur(data.get("Acheteur_Nom")) or nettoyer_valeur(data.get("Utilisateur"))
                             or ACHETEUR_NOM_DEFAUT),
            "acheteur_fonction": nettoyer_valeur(data.get("Acheteur_Fonction")) or ACHETEUR_FONCTION_DEFAUT,
            "code_securite": code,
            "empreinte": emp,
            "texte_qr": f"STE NAP SARL|BC|{reference}|{nom}|{date_emission}|{heure_gen}|{utilisateur}|{code}|{emp}",
            "signature_digitale": signature_digitale,
        }
        ok_pdf = ok_word = False
        try:
            generer_un_bon(d, chemin_pdf)
            ok_pdf = True
        except Exception as ex:
            print(f"  ⚠ Erreur PDF sur {reference} : {ex}")
            resultat["erreurs"].append(f"{reference} (PDF) : {ex}")
        try:
            generer_un_bon_word(d, chemin_word)
            ok_word = True
        except Exception as ex:
            print(f"  ⚠ Erreur Word sur {reference} : {ex}")
            resultat["erreurs"].append(f"{reference} (Word) : {ex}")
        if ok_pdf and ok_word:
            compteur += 1
            print(f"[OK] Bon genere (Word + PDF) : {base_fichier}")

    print("-" * 60)
    if compteur == 0:
        print("Aucun nouveau bon à générer (Word et PDF déjà présents).")
    else:
        print(f"Terminé ! {compteur} bon(s) généré(s) en Word et PDF.")
    resultat["generes"] = compteur
    return resultat


if __name__ == "__main__":
    forcer = any(a in sys.argv for a in ("--tout", "--force", "--regenerer"))
    avec_sig = any(a in sys.argv for a in ("--signature", "--sig", "--cachet"))
    res = generer_tous_les_bons(forcer_tout=forcer, signature_digitale=avec_sig)
    if avec_sig:
        print("Mode : signature numérique Direction activée.")
    if res and res.get("erreurs"):
        print("\nERREURS :", file=sys.stderr)
        for err in res["erreurs"]:
            print(" - " + err, file=sys.stderr)
        sys.exit(1)
