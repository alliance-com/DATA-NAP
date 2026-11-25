"""
NAP SARL — Charte graphique BLEU CIEL
Utilisée par : bons de commande (Word), bons de paiement (Word), rapports (PDF / Excel).
À placer à la RACINE du projet (à côté de generer_bons.py).

Pour changer la teinte partout, modifiez simplement les 5 couleurs ci-dessous (code hexadécimal sans #).
"""
import colorsys

from docx.oxml.ns import qn
from lxml import etree

CIEL = "87CEEB"          # bleu ciel : fonds des titres et des en-têtes de tableaux
CIEL_BORD = "5DADE2"     # bleu ciel soutenu : bordures et traits
CIEL_TEXTE = "2E86C1"    # bleu pour les titres / étiquettes en texte
CIEL_FONCE = "1B4F72"    # bleu foncé : texte posé sur un fond bleu ciel (lisibilité)
CIEL_CLAIR = "EAF6FD"    # bleu très clair : lignes alternées des tableaux
CIEL_MOYEN = "BFE3F7"    # bleu clair intermédiaire

_THEME_FILL = ("themeFill", "themeFillTint", "themeFillShade")
_THEME_COLOR = ("themeColor", "themeTint", "themeShade")
_THEME_TEXTE_COULEUR = {"accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hyperlink",
                        "followedHyperlink", "text2", "dark2", "tx2"}
_THEME_BLANC = {"bg1", "background1", "light1", "window"}


def _rgb(h):
    h = (h or "").strip().lstrip("#")
    if len(h) != 6:
        return None
    try:
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except ValueError:
        return None


def _luminance(h):
    c = _rgb(h)
    return None if c is None else (0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]) / 255


def est_orange(h):
    """True pour les oranges (y compris orange clair / ambre) : ces couleurs sont TOUJOURS conservées."""
    c = _rgb(h)
    if c is None:
        return False
    teinte, sat, val = colorsys.rgb_to_hsv(c[0] / 255, c[1] / 255, c[2] / 255)
    return 8 <= teinte * 360 <= 50 and sat >= 0.10 and val >= 0.35


def _couleurs_theme(doc):
    """Couleurs du thème Word du document (accent1..6, etc.) -> {'accent2': 'ED7D31', ...}"""
    noms = {"dk1": "text1", "lt1": "background1", "dk2": "text2", "lt2": "background2", "hlink": "hyperlink",
            "folHlink": "followedHyperlink"}
    out = {}
    try:
        for part in doc.part.package.iter_parts():
            if "/theme/" in str(part.partname):
                racine = etree.fromstring(part.blob)
                ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
                schema = racine.find(".//a:clrScheme", ns)
                for enfant in schema:
                    nom = etree.QName(enfant).localname
                    if len(enfant):
                        val = enfant[0].get("val") if enfant[0].get("val") not in (None, "windowText", "window") \
                            else enfant[0].get("lastClr")
                        out[noms.get(nom, nom)] = val
                break
    except Exception:
        pass
    return out


def _hex_effectif(valeur, theme, couleurs_theme):
    """Couleur réelle d'un attribut : code hexadécimal, ou couleur du thème résolue."""
    if theme:
        return couleurs_theme.get(theme)
    return valeur if _rgb(valeur) else None


def _sans_theme(el, attrs):
    for a in attrs:
        key = qn(f"w:{a}")
        if key in el.attrib:
            del el.attrib[key]


def _racines(doc):
    """Corps du document + en-têtes / pieds de page qui existent réellement."""
    racines = [doc.element.body]
    for s in doc.sections:
        for p in (s.header, s.footer, s.first_page_header, s.first_page_footer,
                  s.even_page_header, s.even_page_footer):
            try:
                if not p.is_linked_to_previous:
                    racines.append(p._element)
            except Exception:
                pass
    return racines


def appliquer_theme_docx(doc):
    """Recolorie un document Word en bleu ciel : bordures, fonds colorés, titres colorés.
    - le noir, le gris et le blanc du texte ne changent pas ;
    - l'ORANGE (fonds, texte, bordures) est conservé tel quel ;
    - le rouge (ex. numéro de référence) est conservé ;
    - un fond foncé devient bleu ciel et son texte blanc devient bleu foncé (lisible).
    Retourne le nombre d'éléments modifiés."""
    modifies = 0
    couleurs_theme = _couleurs_theme(doc)
    for racine in _racines(doc):
        # 1) bordures de tableaux, de cellules, de paragraphes et de page
        for el in racine.xpath(".//w:tcBorders/* | .//w:tblBorders/* | .//w:pBdr/* | .//w:pgBorders/*"):
            if el.get(qn("w:val")) in (None, "nil", "none"):
                continue
            if est_orange(_hex_effectif(el.get(qn("w:color")), el.get(qn("w:themeColor")), couleurs_theme)):
                continue                       # bordure orange : conservée
            _sans_theme(el, _THEME_COLOR)
            el.set(qn("w:color"), CIEL_BORD)
            modifies += 1

        # 2) couleurs de texte (avant les fonds, pour ne pas retoucher le texte corrigé plus bas)
        for el in racine.xpath(".//w:rPr/w:color"):
            val, theme = el.get(qn("w:val")), el.get(qn("w:themeColor"))
            if est_orange(_hex_effectif(val, theme, couleurs_theme)):
                continue                       # texte orange : conservé
            if theme:
                if theme in _THEME_TEXTE_COULEUR:
                    _sans_theme(el, _THEME_COLOR)
                    el.set(qn("w:val"), CIEL_TEXTE)
                    modifies += 1
                continue
            c = _rgb(val)
            if c is None:
                continue
            r, g, b = c
            if max(c) - min(c) < 25:          # noir / gris / blanc : inchangé
                continue
            if r > 150 and g < 110 and b < 110:  # rouge : conservé
                continue
            el.set(qn("w:val"), CIEL_TEXTE)
            modifies += 1

        # 3) fonds colorés
        for shd in racine.xpath(".//w:shd"):
            fill, theme_fill = shd.get(qn("w:fill")), shd.get(qn("w:themeFill"))
            if est_orange(_hex_effectif(fill, theme_fill, couleurs_theme)):
                continue                       # fond orange : conservé (et son texte aussi)
            if theme_fill:
                if theme_fill in _THEME_BLANC:
                    continue
                lum = 0.3 if theme_fill in ("text1", "dark1", "tx1", "text2", "dark2", "tx2") else 0.6
            elif fill is None or fill.lower() in ("auto", "ffffff"):
                continue
            else:
                lum = _luminance(fill)
                if lum is None:
                    continue
            if lum >= 0.96:                     # quasi blanc : inchangé
                continue
            nouveau = CIEL if lum < 0.45 else (CIEL_CLAIR if lum > 0.88 else CIEL_MOYEN)
            _sans_theme(shd, _THEME_FILL)
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), nouveau)
            modifies += 1
            if lum < 0.45:                      # texte blanc sur fond foncé -> bleu foncé sur bleu ciel
                parent = shd.getparent()
                conteneur = parent.getparent() if parent is not None else None
                if conteneur is not None:
                    for col in conteneur.xpath(".//w:rPr/w:color"):
                        v, t = col.get(qn("w:val")), col.get(qn("w:themeColor"))
                        l = _luminance(v)
                        if t in _THEME_BLANC or (not t and l is not None and l > 0.8):
                            _sans_theme(col, _THEME_COLOR)
                            col.set(qn("w:val"), CIEL_FONCE)
    return modifies
