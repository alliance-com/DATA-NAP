import streamlit as st
import pandas as pd
from datetime import datetime, date
import os
import re
from io import BytesIO
import importlib

import telechargement_bons as tb
importlib.reload(tb)   # relit le fichier à chaque affichage

st.set_page_config(page_title="Paiements", page_icon="💳", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"
username = st.session_state.get("username") or ""

# ==================== CONFIG ====================
FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"
FEUILLE_DEPENSES = "DEPENSES"
DOSSIER_BONS = "bons_paiement"
os.makedirs(DOSSIER_BONS, exist_ok=True)

TYPES_OP = ["PURCHASE", "EXPORT", "SERVICE", "SALARY", "AUTRE"]
PRODUITS = ["SHEA", "SOYABEAN", "COTTON", "FRET EXPORT", "TRANSFERT DHL", "SERVICE", "ENTRETIEN", "SALARY", "BFU", "AUTRE"]
MODES = ["CHEQUE", "ESPECES", "VIREMENT", "MOBILE MONEY"]
BANQUES = ["ECOBANK BENIN", "CAISSE", "BIIC", "ATLANTIC BANK BENIN", "BGFI", "AUTRE"]
OBSERVATIONS = ["PAYE", "AVANCE", "PARTIEL", "EN ATTENTE"]

COLONNES = [
    "Date", "No de transaction", "Type d'operation", "Produit",
    "Prestataires/Fournisseurs", "Nom du bénéficiaire", "Num. Pieces d'identite",
    "Date d'expiration", "Addresse /Telephone", "Ref doc",
    "Description/Objet du paiement", "Montant", "Observation",
    "Montant en lettre", "Mode paiement", "Numero de cheque", "Banque",
    "N_BC",  # N° bon(s) de commande couverts (PURCHASE / fournisseurs)
    "N_Dechargement",  # N° transaction(s) de déchargement / pesée liés
]
# =================================================



def valeurs_uniques(df, colonne, extras=None):
    """Liste déroulante = valeurs déjà saisies dans Excel + extras de base."""
    vals = []
    if extras:
        vals.extend(extras)
    if df is not None and not df.empty and colonne in df.columns:
        for v in df[colonne].dropna().astype(str):
            v = v.strip()
            if v and v.lower() not in ("nan", "none", "-"):
                vals.append(v)
    # unique preserve order
    seen = set()
    out = []
    for v in vals:
        key = v.casefold()
        if key not in seen:
            seen.add(key)
            out.append(v)
    return out


def choix_ou_nouveau(label, options, key_prefix):
    """Liste déroulante (valeurs Excel) + champ libre pour une NOUVELLE valeur.
    Si le champ libre est rempli, il a priorité sur la liste.
    """
    options = list(options) if options else []
    # Ligne vide en tête pour forcer un choix explicite
    options_affichees = [""] + options
    choix = st.selectbox(
        label,
        options=options_affichees,
        key=f"{key_prefix}_sel",
        format_func=lambda x: "— Choisir —" if x == "" else x,
    )
    nouveau = st.text_input(
        f"✏️ Nouvelle valeur (si absente de la liste)",
        key=f"{key_prefix}_new",
        placeholder="Écrire ici pour ajouter une valeur qui n'est pas dans la liste",
        help="Laissez vide pour utiliser la liste. Si vous tapez ici, cette valeur sera utilisée.",
    )
    nouveau = (nouveau or "").strip()
    if nouveau:
        return nouveau
    return (choix or "").strip()

# ==================== ANNUAIRE (listes déroulantes liées à l'Excel) ====================
COL_PRESTA = "Prestataires/Fournisseurs"
COL_BENEF = "Nom du bénéficiaire"
COL_PIECE = "Num. Pieces d'identite"
COL_EXP = "Date d'expiration"
COL_TEL = "Addresse /Telephone"


def _propre(v):
    """Valeur Excel -> texte propre ('' si vide ou placeholder '-')."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, (datetime, date, pd.Timestamp)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    t = str(v).strip()
    if t.lower() in ("", "nan", "none", "nat", "-"):
        return ""
    # date texte "2031-06-02 00:00:00" -> "02/06/2031"
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", t)
    if m:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    return t


def _uniques_recents(serie):
    """Valeurs uniques non vides, la plus récente (dernière ligne Excel) en premier."""
    out, vus = [], set()
    for v in reversed(serie.tolist()):
        t = _propre(v)
        if t and t.casefold() not in vus:
            vus.add(t.casefold())
            out.append(t)
    return out


def charger_annuaire(df):
    """Construit l'annuaire depuis la feuille DEPENSES (base de données).
    Retourne {prestataire: DataFrame de ses lignes}, trié par nom."""
    if df is None or df.empty or COL_PRESTA not in df.columns:
        return {}
    annuaire = {}
    tmp = df.copy()
    tmp["_presta"] = tmp[COL_PRESTA].map(_propre)
    tmp = tmp[tmp["_presta"] != ""]
    for nom, sous in tmp.groupby("_presta", sort=True):
        annuaire[nom] = sous
    return annuaire


def extraire_num_trans(val):
    if not val:
        return 0
    m = re.search(r"(\d+)", str(val).replace(" ", ""))
    return int(m.group(1)) if m else 0


def prochain_numero(df, prefix="CHQ"):
    """CHQ-0390-2026 ou CC-0138-2026 selon le mode."""
    annee = datetime.now().year
    nums = []
    if not df.empty and "No de transaction" in df.columns:
        for v in df["No de transaction"].dropna().astype(str):
            if prefix in v.upper() or (prefix == "CHQ" and "CHQ" in v.upper()) or (prefix == "CC" and v.upper().startswith("CC")):
                n = extraire_num_trans(v)
                if n:
                    nums.append(n)
    n = max(nums) + 1 if nums else 1
    if prefix == "CHQ":
        return f"CHQ-{n:04d}-{annee}"
    return f"CC-{n:04d}-{annee}"


def montant_en_lettres(n):
    """Conversion basique FCFA en lettres (français)."""
    try:
        n = int(round(float(n)))
    except Exception:
        return ""
    if n == 0:
        return "Zéro franc CFA"

    uns = ["", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf",
           "dix", "onze", "douze", "treize", "quatorze", "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf"]
    diz = ["", "dix", "vingt", "trente", "quarante", "cinquante", "soixante", "soixante", "quatre-vingt", "quatre-vingt"]

    def sous_mille(x):
        if x < 20:
            return uns[x]
        if x < 100:
            d, u = divmod(x, 10)
            if d == 7 or d == 9:  # 70-79, 90-99
                base = 60 if d == 7 else 80
                reste = x - base
                liaison = "-" if reste else ""
                prefix = "soixante" if d == 7 else "quatre-vingt"
                if d == 9:
                    return f"quatre-vingt-{uns[10 + (x - 90)]}" if x >= 90 else prefix
                if reste < 20:
                    return f"{prefix}-{uns[reste]}" if reste else prefix
            if d == 8:
                return f"quatre-vingt-{uns[u]}" if u else "quatre-vingts"
            sep = " et " if u == 1 and d in (2, 3, 4, 5, 6) else ("-" if u else "")
            return f"{diz[d]}{sep}{uns[u]}" if u else diz[d]
        c = x // 100
        r = x % 100
        cent = "cent" if c == 1 else f"{uns[c]} cent"
        if r == 0 and c > 1:
            cent = f"{uns[c]} cents"
        return f"{cent} {sous_mille(r)}".strip() if r else cent

    parts = []
    millions = n // 1_000_000
    reste = n % 1_000_000
    milliers = reste // 1000
    unites = reste % 1000

    if millions:
        if millions == 1:
            parts.append("un million")
        else:
            parts.append(f"{sous_mille(millions)} millions")
    if milliers:
        if milliers == 1:
            parts.append("mille")
        else:
            parts.append(f"{sous_mille(milliers)} mille")
    if unites:
        parts.append(sous_mille(unites))

    texte = " ".join(parts).strip()
    if texte:
        texte = texte[0].upper() + texte[1:] + " francs CFA"
    return texte



def charger_liste_bc():
    """Liste des N° BC depuis donnees_bons.xlsx (si présent)."""
    fichiers = ["donnees_bons.xlsx", "donnees_bons_soyabean.xlsx", "donnees_bons_shea.xlsx"]
    refs = []
    for f in fichiers:
        if not os.path.exists(f):
            continue
        try:
            d = pd.read_excel(f)
            col = None
            for c in d.columns:
                if str(c).strip().lower() in ("reference", "ref_bc", "n° bc", "ref"):
                    col = c
                    break
            if col is None and len(d.columns):
                col = d.columns[0]
            for v in d[col].dropna().astype(str):
                v = v.strip()
                if v and v.lower() not in ("nan", "none"):
                    refs.append(v)
        except Exception:
            pass
    return refs


def charger_liste_dechargements():
    """
    Liste des déchargements (Transaction) depuis dechargements_*.xlsx
    et dechargement_data si disponible.
    Retourne liste de labels « Trans · Frs · Camion · Poids · BC » et map label -> Transaction.
    """
    labels = []
    map_tr = {}
    frames = []

    # 1) via dechargement_data (tous produits)
    try:
        import dechargement_data as dd
        ancien = dd.get_produit()
        try:
            for p in getattr(dd, "PRODUITS", []):
                try:
                    dd.set_produit(p)
                    d = dd.charger_dechargements()
                    if d is not None and not d.empty:
                        d = d.copy()
                        d["_Produit"] = p
                        frames.append(d)
                except Exception:
                    continue
        finally:
            try:
                dd.set_produit(ancien)
            except Exception:
                pass
    except Exception:
        pass

    # 2) fichiers Excel directs si module absent / vide
    if not frames:
        import glob as _glob
        for f in _glob.glob("dechargements_*.xlsx") + (
            ["dechargements.xlsx"] if os.path.exists("dechargements.xlsx") else []
        ):
            try:
                d = pd.read_excel(f, sheet_name=0)
                if "Transaction" in d.columns or any("transaction" in str(c).lower() for c in d.columns):
                    frames.append(d)
            except Exception:
                continue

    if not frames:
        return labels, map_tr

    df = pd.concat(frames, ignore_index=True)
    # normaliser colonnes
    rename = {}
    for c in df.columns:
        k = str(c).strip().lower()
        if k == "transaction":
            rename[c] = "Transaction"
        elif k in ("fournisseur",):
            rename[c] = "Fournisseur"
        elif k in ("n_camion", "n° camion", "camion"):
            rename[c] = "N_Camion"
        elif k in ("poids_net_kg", "poids net (kg)", "poids_net"):
            rename[c] = "Poids_Net_KG"
        elif k in ("ref_bc", "n° bc", "reference"):
            rename[c] = "Ref_BC"
        elif k in ("weighment_ticket", "ticket"):
            rename[c] = "Weighment_Ticket"
        elif k in ("date_dechargement", "date"):
            rename[c] = "Date_Dechargement"
    df = df.rename(columns=rename)
    if "Transaction" not in df.columns:
        return labels, map_tr

    def _s(v):
        if v is None:
            return ""
        try:
            if pd.isna(v):
                return ""
        except Exception:
            pass
        t = str(v).strip()
        return "" if t.lower() in ("nan", "none", "nat") else t

    # plus récents en premier
    if "Date_Dechargement" in df.columns:
        df["_d"] = pd.to_datetime(df["Date_Dechargement"], errors="coerce", dayfirst=True)
        df = df.sort_values("_d", ascending=False)

    seen = set()
    for _, r in df.iterrows():
        tr = _s(r.get("Transaction"))
        if not tr or tr in seen:
            continue
        seen.add(tr)
        frs = _s(r.get("Fournisseur"))
        cam = _s(r.get("N_Camion"))
        try:
            poids = float(r.get("Poids_Net_KG") or 0)
            ptxt = f"{poids:,.0f} kg".replace(",", " ") if poids else ""
        except Exception:
            ptxt = ""
        bc = _s(r.get("Ref_BC"))
        ticket = _s(r.get("Weighment_Ticket"))
        parts = [tr]
        if frs:
            parts.append(frs[:22])
        if cam:
            parts.append(cam)
        if ptxt:
            parts.append(ptxt)
        if bc:
            parts.append(f"BC {bc}")
        elif ticket:
            parts.append(f"Tkt {ticket}")
        lab = " · ".join(parts)
        labels.append(lab)
        map_tr[lab] = tr
    return labels, map_tr


def charger_depenses():
    """Charge la feuille DEPENSES.
    Dans le fichier NAP, les en-têtes sont en ligne 3 (header=2).
    """
    if not os.path.exists(FICHIER_PAIEMENTS):
        return pd.DataFrame(columns=COLONNES)
    try:
        xl = pd.ExcelFile(FICHIER_PAIEMENTS)
        sheet = FEUILLE_DEPENSES if FEUILLE_DEPENSES in xl.sheet_names else xl.sheet_names[0]

        # Essayer plusieurs lignes d'en-tête (fichier NAP = ligne 3)
        df = None
        for hdr in (2, 1, 0, 3):
            try:
                tmp = pd.read_excel(FICHIER_PAIEMENTS, sheet_name=sheet, header=hdr)
                cols_str = " ".join(str(c).lower() for c in tmp.columns)
                if "prestataire" in cols_str or "transaction" in cols_str or "bénéficiaire" in cols_str or "beneficiaire" in cols_str:
                    df = tmp
                    break
            except Exception:
                continue
        if df is None:
            df = pd.read_excel(FICHIER_PAIEMENTS, sheet_name=sheet, header=2)

        # Supprimer colonnes vides / Unnamed inutiles
        df = df.dropna(axis=1, how="all")
        df = df.loc[:, [c for c in df.columns if not str(c).startswith("Unnamed")]]

        # Mapping flexible des noms de colonnes
        alias = {
            "date": "Date",
            "no de transaction": "No de transaction",
            "n° de transaction": "No de transaction",
            "numero de transaction": "No de transaction",
            "type d'operation": "Type d'operation",
            "type d'opération": "Type d'operation",
            "type operation": "Type d'operation",
            "produit": "Produit",
            "prestataires/fournisseurs": "Prestataires/Fournisseurs",
            "prestataire": "Prestataires/Fournisseurs",
            "fournisseur": "Prestataires/Fournisseurs",
            "nom du bénéficiaire": "Nom du bénéficiaire",
            "nom du beneficiaire": "Nom du bénéficiaire",
            "bénéficiaire": "Nom du bénéficiaire",
            "beneficiaire": "Nom du bénéficiaire",
            "nom du percepteur": "Nom du bénéficiaire",
            "percepteur": "Nom du bénéficiaire",
            "num. pieces d'identite": "Num. Pieces d'identite",
            "num pieces d'identite": "Num. Pieces d'identite",
            "piece d'identite": "Num. Pieces d'identite",
            "date d'expiration": "Date d'expiration",
            "addresse /telephone": "Addresse /Telephone",
            "adresse /telephone": "Addresse /Telephone",
            "telephone": "Addresse /Telephone",
            "ref doc": "Ref doc",
            "description/objet du paiement": "Description/Objet du paiement",
            "description": "Description/Objet du paiement",
            "objet du paiement": "Description/Objet du paiement",
            "montant": "Montant",
            "observation": "Observation",
            "montant en lettre": "Montant en lettre",
            "mode paiement": "Mode paiement",
            "numero de cheque": "Numero de cheque",
            "banque": "Banque",
            "n_bc": "N_BC",
            "n° bc": "N_BC",
            "n° bon de commande": "N_BC",
            "bons de commande": "N_BC",
            "ref bc": "N_BC",
            "n_dechargement": "N_Dechargement",
            "n° déchargement": "N_Dechargement",
            "n° dechargement": "N_Dechargement",
            "transactions déchargement": "N_Dechargement",
            "transactions dechargement": "N_Dechargement",
            "déchargements": "N_Dechargement",
            "dechargements": "N_Dechargement",
        }
        rename = {}
        for c in df.columns:
            key = str(c).strip().lower().replace("  ", " ")
            if key in alias:
                rename[c] = alias[key]
            else:
                for cible in COLONNES:
                    if key == cible.lower():
                        rename[c] = cible
                        break
        df = df.rename(columns=rename)

        for c in COLONNES:
            if c not in df.columns:
                df[c] = ""

        # Enlever lignes totalement vides (sans transaction ni montant)
        if "No de transaction" in df.columns:
            df = df[df["No de transaction"].notna()]
            df = df[df["No de transaction"].astype(str).str.strip().str.lower().isin(["", "nan", "none"]) == False]

        return df[COLONNES].reset_index(drop=True)
    except Exception as e:
        st.error(f"Erreur lecture Excel : {e}")
        return pd.DataFrame(columns=COLONNES)


def sauvegarder_ligne(nouvelle):
    """Ajoute une ligne à la fin du tableau de la feuille DEPENSES.
    - repère automatiquement la ligne d'en-têtes (ligne 3, colonne B dans le fichier NAP)
    - écrit sous la dernière opération, colonne par colonne, en recopiant le format
    - agrandit le tableau Excel (Table1) pour inclure la nouvelle ligne
    Les autres feuilles ne sont pas modifiées."""
    from copy import copy
    from openpyxl import load_workbook, Workbook

    if not os.path.exists(FICHIER_PAIEMENTS):
        wb = Workbook()
        ws = wb.active
        ws.title = FEUILLE_DEPENSES
        ws.append(COLONNES)
        wb.save(FICHIER_PAIEMENTS)

    wb = load_workbook(FICHIER_PAIEMENTS)
    ws = wb[FEUILLE_DEPENSES] if FEUILLE_DEPENSES in wb.sheetnames else wb.create_sheet(FEUILLE_DEPENSES)

    def norm(x):
        return str(x or "").strip().lower()

    # 1) ligne d'en-têtes : celle qui contient « No de transaction »
    hdr_row = None
    for r in range(1, 15):
        if any(norm(c.value) == "no de transaction" for c in ws[r]):
            hdr_row = r
            break
    if hdr_row is None:  # feuille vide : on crée les en-têtes en ligne 1
        for i, h in enumerate(COLONNES, 1):
            ws.cell(1, i, h)
        hdr_row = 1

    col_of = {norm(c.value): c.column for c in ws[hdr_row] if c.value}
    col_trans = col_of.get("no de transaction", 1)

    # 2) dernière ligne réellement remplie (colonne N° transaction)
    last = hdr_row
    for r in range(ws.max_row, hdr_row, -1):
        if ws.cell(r, col_trans).value not in (None, ""):
            last = r
            break
    new_row = last + 1

    # 3) écriture colonne par colonne + format de la ligne précédente
    valeurs = {norm(k): v for k, v in nouvelle.items()}
    for h, col in col_of.items():
        val = valeurs.get(h, None)
        if h == "date" and isinstance(val, str):
            try:
                val = datetime.strptime(val, "%d/%m/%Y")
            except ValueError:
                pass
        cell = ws.cell(new_row, col, val if val != "" else None)
        if last > hdr_row:
            src = ws.cell(last, col)
            if src.has_style:
                cell._style = copy(src._style)

    # 4) agrandir le tableau Excel pour inclure la nouvelle ligne
    for t in ws.tables.values():
        t.ref = f"{t.ref.split(':')[0]}:{t.ref.split(':')[1].rstrip('0123456789')}{new_row}"

    wb.save(FICHIER_PAIEMENTS)


# ---------- Sidebar formulaire ----------
st.sidebar.header("➕ Nouveau paiement")

df = charger_depenses()

# Listes issues de l'Excel (historique) + valeurs de base
opt_types = valeurs_uniques(df, "Type d'operation", TYPES_OP)
opt_produits = valeurs_uniques(df, "Produit", PRODUITS)
opt_prestataires = valeurs_uniques(df, "Prestataires/Fournisseurs")
opt_beneficiaires = valeurs_uniques(df, "Nom du bénéficiaire")
opt_descriptions = valeurs_uniques(df, "Description/Objet du paiement")
opt_banques = valeurs_uniques(df, "Banque", BANQUES)
opt_modes = valeurs_uniques(df, "Mode paiement", MODES)

# Champs en listes déroulantes HORS formulaire → "➕ Nouveau…" affiche tout de suite le champ texte
st.sidebar.markdown("**Listes Excel** — pour une valeur absente, écrivez-la dans le champ *Nouvelle valeur*")
type_op = choix_ou_nouveau("Type d'opération *", opt_types, "type_op")
produit = choix_ou_nouveau("Produit *", opt_produits, "produit")
prestataire = choix_ou_nouveau("Prestataire / Fournisseur *", opt_prestataires, "prestataire")
beneficiaire = choix_ou_nouveau("Nom du percepteur / bénéficiaire *", opt_beneficiaires, "benef")
description = choix_ou_nouveau("Description / Objet *", opt_descriptions, "desc")

mode_defaut = st.sidebar.selectbox("Mode (pour N° auto)", opt_modes if opt_modes else MODES, index=0)
prefix = "CHQ" if "CHEQUE" in str(mode_defaut).upper() else "CC"
num_suggere = prochain_numero(df, prefix=prefix)

with st.sidebar.form("form_paiement", clear_on_submit=True):
    date_p = st.date_input("Date *", value=date.today())
    no_trans = st.text_input("N° transaction *", value=num_suggere)
    piece_id = st.text_input("N° pièce d'identité")
    date_exp = st.text_input("Date expiration pièce", placeholder="JJ/MM/AAAA")
    telephone = st.text_input("Adresse / Téléphone")
    ref_doc = st.text_input("Réf. Doc", placeholder="Camion_MT / facture / waybill…")
    montant = st.number_input("Montant (FCFA) *", min_value=0.0, step=100.0, format="%.0f")
    observation = st.selectbox("Observation", OBSERVATIONS)
    mode_pay = st.selectbox("Mode paiement *", opt_modes if opt_modes else MODES)
    num_cheque = st.text_input("N° chèque", placeholder="si chèque")
    banque = st.selectbox("Banque", opt_banques if opt_banques else BANQUES)

    # N° BC — pour PURCHASE / fournisseurs (un ou plusieurs)
    liste_bc = charger_liste_bc()
    # + BC déjà saisis dans N_BC des paiements
    if "N_BC" in df.columns:
        for v in df["N_BC"].dropna().astype(str):
            for part in re.split(r"[,;&|/]+", v):
                part = part.strip()
                if part and part not in liste_bc:
                    liste_bc.append(part)
    # dédoublonnage en gardant l'ordre
    _seen = set()
    liste_bc_u = []
    for x in liste_bc:
        k = x.casefold()
        if k not in _seen:
            _seen.add(k)
            liste_bc_u.append(x)
    n_bc_sel = st.multiselect(
        "N° Bon(s) de commande (fournisseurs / PURCHASE)",
        options=liste_bc_u,
        help="Sélectionnez tous les bons de commande couverts par ce paiement. Laisser vide si non applicable.",
    )
    n_bc_libre = st.text_input(
        "Autre N° BC non listé (séparés par , ou &)",
        placeholder="ex. 0012-BC-SOYABEAN/NAP-2026, 0015-BC-SOYABEAN/NAP-2026",
    )
    n_bc_parts = list(n_bc_sel)
    if n_bc_libre.strip():
        for part in re.split(r"[,;&]+", n_bc_libre):
            part = part.strip()
            if part and part not in n_bc_parts:
                n_bc_parts.append(part)
    n_bc_str = " & ".join(n_bc_parts)

    # --- Déchargements liés (comme les BC) ---
    labels_dech, map_dech = charger_liste_dechargements()
    # + déjà saisis dans N_Dechargement
    if "N_Dechargement" in df.columns:
        for v in df["N_Dechargement"].dropna().astype(str):
            for part in re.split(r"[,;&|/]+", v):
                part = part.strip()
                if part and part not in map_dech.values() and part not in labels_dech:
                    labels_dech.append(part)
                    map_dech[part] = part
    n_dech_sel = st.multiselect(
        "Déchargement(s) lié(s) (N° transaction pesée)",
        options=labels_dech,
        help="Sélectionnez un ou plusieurs déchargements couverts par ce paiement "
             "(camion / ticket / fournisseur).",
    )
    n_dech_libre = st.text_input(
        "Autre N° déchargement non listé (séparés par , ou &)",
        placeholder="ex. 4521, 4522",
    )
    n_dech_parts = []
    for lab in n_dech_sel:
        tr = map_dech.get(lab, lab)
        if tr and tr not in n_dech_parts:
            n_dech_parts.append(tr)
    if n_dech_libre.strip():
        for part in re.split(r"[,;&]+", n_dech_libre):
            part = part.strip()
            if part and part not in n_dech_parts:
                n_dech_parts.append(part)
    n_dech_str = " & ".join(n_dech_parts)

    submitted = st.form_submit_button("💾 Enregistrer + générer le bon", width="stretch")

    if submitted:
        if not no_trans.strip() or not prestataire.strip() or not beneficiaire.strip():
            st.error("N° transaction, prestataire et bénéficiaire obligatoires.")
        elif montant <= 0:
            st.error("Montant invalide.")
        elif not description.strip():
            st.error("Description obligatoire.")
        else:
            lettres = montant_en_lettres(montant)
            nouvelle = {
                "Date": date_p.strftime("%d/%m/%Y"),
                "No de transaction": no_trans.strip(),
                "Type d'operation": type_op,
                "Produit": produit,
                "Prestataires/Fournisseurs": prestataire.strip(),
                "Nom du bénéficiaire": beneficiaire.strip(),
                "Num. Pieces d'identite": piece_id.strip(),
                "Date d'expiration": date_exp.strip(),
                "Addresse /Telephone": telephone.strip(),
                "Ref doc": ref_doc.strip(),
                "Description/Objet du paiement": description.strip(),
                "Montant": montant,
                "Observation": observation,
                "Montant en lettre": lettres,
                "Mode paiement": mode_pay,
                "Numero de cheque": num_cheque.strip() or "-",
                "Banque": banque,
                "N_BC": n_bc_str,
                "N_Dechargement": n_dech_str,
            }
            try:
                sauvegarder_ligne(nouvelle)
            except Exception as e:
                st.error(f"Erreur sauvegarde Excel : {e}")
            else:
                # Sync automatique vers Caisse NAP si mode Espèces / Mobile Money / Caisse
                try:
                    import caisse_data as _cx
                    mode_u = str(mode_pay or "").upper()
                    if any(k in mode_u for k in ("ESPECE", "MOBILE", "MOMO", "WAVE", "CAISSE", "CASH")):
                        # une seule ligne : utiliser import global (anti-doublon SP-)
                        _ok, _msg, _n = _cx.importer_depuis_suivi_paiement(
                            seulement_caisse=True, user=name
                        )
                        if _ok and _n:
                            st.info(f"💵 Caisse NAP : {_msg}")
                except Exception as _e:
                    st.caption(f"(Sync caisse non effectuée : {_e})")
                # Génération Word
                try:
                    from generer_bon_paiement import generer_bon_paiement
                    data = {
                        "Date": nouvelle["Date"],
                        "No_transaction": nouvelle["No de transaction"],
                        "Type_operation": type_op,
                        "Produit": produit,
                        "Prestataire": prestataire.strip(),
                        "Beneficiaire": beneficiaire.strip(),
                        "Piece_identite": piece_id.strip(),
                        "Date_expiration": date_exp.strip(),
                        "Telephone": telephone.strip(),
                        "Ref_doc": ref_doc.strip(),
                        "Description": description.strip(),
                        "Montant": montant,
                        "Observation": observation,
                        "Montant_lettres": lettres,
                        "Mode_paiement": mode_pay,
                        "Numero_cheque": num_cheque.strip() or "-",
                        "Banque": banque,
                        "Utilisateur": name,
                        "N_BC": n_bc_str,
                        "N_Dechargement": n_dech_str,
                    }
                    chemin = generer_bon_paiement(data)
                    st.success(f"✅ Enregistré + bon généré : `{os.path.basename(chemin)}`")
                except Exception as e:
                    st.warning(f"Enregistré dans Excel, mais erreur Word : {e}")
                    import traceback
                    st.code(traceback.format_exc())
                st.rerun()

# ---------- Main ----------
st.title("💳 Suivi des paiements & bons de paiement")
st.caption(f"Fichier : `{FICHIER_PAIEMENTS}` · Feuille : `{FEUILLE_DEPENSES}` · Bons Word : `{DOSSIER_BONS}/`")

df = charger_depenses()

# KPI
if not df.empty:
    montants = pd.to_numeric(df["Montant"], errors="coerce").fillna(0)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Nb opérations", len(df))
    k2.metric("Total payé (FCFA)", f"{montants.sum():,.0f}".replace(",", " "))
    k3.metric("Chèques", len(df[df["Mode paiement"].astype(str).str.upper().str.contains("CHEQUE", na=False)]))
    k4.metric("Espèces / MM", len(df[df["Mode paiement"].astype(str).str.upper().str.contains("ESPECES|MOBILE", na=False, regex=True)]))
else:
    st.info("Aucune dépense chargée. Vérifiez que le fichier `Suivi de paiement.xlsx` est dans le dossier du projet.")

st.markdown("---")

# Filtres
f1, f2, f3, f4 = st.columns(4)
with f1:
    ft = st.selectbox("Type", ["Tous"] + TYPES_OP)
with f2:
    fp = st.selectbox("Produit", ["Tous"] + sorted(df["Produit"].dropna().astype(str).unique().tolist()) if not df.empty else ["Tous"])
with f3:
    fm = st.selectbox("Mode", ["Tous"] + MODES)
with f4:
    q = st.text_input("Recherche", placeholder="N°, fournisseur, réf…")

df_f = df.copy()
if not df_f.empty:
    if ft != "Tous":
        df_f = df_f[df_f["Type d'operation"].astype(str) == ft]
    if fp != "Tous":
        df_f = df_f[df_f["Produit"].astype(str) == fp]
    if fm != "Tous":
        df_f = df_f[df_f["Mode paiement"].astype(str).str.upper() == fm]
    if q.strip():
        qq = q.strip().lower()
        texte = df_f.fillna("").astype(str).agg(" ".join, axis=1).str.lower()
        df_f = df_f[texte.str.contains(qq, regex=False, na=False)]

st.subheader("📋 Journal des dépenses")
if df_f.empty:
    st.warning("Aucune ligne à afficher.")
else:
    st.dataframe(df_f, width="stretch", height=400, hide_index=True)


# ---------- SYNC CAISSE NAP ----------
st.markdown("---")
st.subheader("💵 Synchroniser les paiements caisse → Caisse NAP")
st.caption(
    "Importe dans **Caisse NAP** les paiements en **ESPECES / MOBILE MONEY / CAISSE** "
    "présents dans ce Suivi de paiement (anti-doublon par N° transaction `SP-…`)."
)
try:
    import caisse_data as _cx_sync
    _cand, _stats, _err = _cx_sync.apercu_import_suivi(seulement_caisse=True)
    if _err:
        st.warning(_err)
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Lignes suivi", _stats.get("total_suivi", 0))
        c2.metric("Modes caisse", _stats.get("candidats", 0))
        c3.metric("Déjà en caisse", _stats.get("deja_importes", 0))
        c4.metric("Nouveaux", _stats.get("nouveaux", 0))
        if st.button("📥 Synchroniser maintenant vers Caisse NAP", type="primary", key="sync_caisse_from_pay"):
            _ok, _msg, _n = _cx_sync.importer_depuis_suivi_paiement(seulement_caisse=True, user=name)
            if _ok:
                st.success(_msg)
            else:
                st.error(_msg)
except Exception as e:
    st.caption(f"Module caisse_data indisponible : {e}")

# ---------- MODIFIER / SUPPRIMER UN PAIEMENT ----------
st.markdown("---")
st.subheader("✏️ Modifier ou 🗑️ supprimer un paiement")

df_mod = charger_depenses()
if df_mod.empty:
    st.info("Aucun paiement à modifier.")
else:
    options_trans = []
    seen_t = set()
    for o in reversed(df_mod["No de transaction"].astype(str).tolist()):
        if o and o not in seen_t and o.lower() not in ("nan", "none"):
            seen_t.add(o)
            options_trans.append(o)
    options_trans = list(reversed(options_trans))

    trans_sel = st.selectbox(
        "N° de transaction",
        options=[""] + options_trans,
        key="mod_pay_sel",
        format_func=lambda x: "— Choisir —" if x == "" else x
    )
    if trans_sel:
        lignes_m = df_mod[df_mod["No de transaction"].astype(str) == str(trans_sel)]
        if not lignes_m.empty:
            row_m = lignes_m.iloc[-1]
            with st.form("form_mod_pay"):
                st.caption(f"Modification de **{trans_sel}**")
                c1, c2 = st.columns(2)
                with c1:
                    e_date = st.text_input("Date (JJ/MM/AAAA)", value=_propre(row_m.get("Date", "")))
                    e_type = st.text_input("Type d'opération", value=_propre(row_m.get("Type d'operation", "")))
                    e_prod = st.text_input("Produit", value=_propre(row_m.get("Produit", "")))
                    e_presta = st.text_input("Prestataire *", value=_propre(row_m.get("Prestataires/Fournisseurs", "")))
                    e_benef = st.text_input("Bénéficiaire *", value=_propre(row_m.get("Nom du bénéficiaire", "")))
                    e_piece = st.text_input("N° pièce d'identité", value=_propre(row_m.get("Num. Pieces d'identite", "")))
                    e_exp = st.text_input("Date expiration", value=_propre(row_m.get("Date d'expiration", "")))
                    e_tel = st.text_input("Adresse / Téléphone", value=_propre(row_m.get("Addresse /Telephone", "")))
                with c2:
                    e_ref = st.text_input("Réf. Doc", value=_propre(row_m.get("Ref doc", "")))
                    e_desc = st.text_input("Description *", value=_propre(row_m.get("Description/Objet du paiement", "")))
                    try:
                        mt0 = float(pd.to_numeric(row_m.get("Montant"), errors="coerce") or 0)
                    except Exception:
                        mt0 = 0.0
                    e_montant = st.number_input("Montant (FCFA) *", min_value=0.0, value=mt0, step=100.0, format="%.0f")
                    e_obs = st.selectbox("Observation", OBSERVATIONS,
                                         index=OBSERVATIONS.index(row_m["Observation"]) if _propre(row_m.get("Observation")) in OBSERVATIONS else 0)
                    e_mode = st.selectbox("Mode paiement *", MODES,
                                          index=next((i for i, m in enumerate(MODES) if m.upper() in str(row_m.get("Mode paiement", "")).upper()), 0))
                    e_cheque = st.text_input("N° chèque", value=_propre(row_m.get("Numero de cheque", "")))
                    e_banque = st.text_input("Banque", value=_propre(row_m.get("Banque", "")))

                b_save = st.form_submit_button("💾 Enregistrer les modifications", type="primary", width="stretch")
                if b_save:
                    if not e_presta.strip() or not e_benef.strip() or not e_desc.strip() or e_montant <= 0:
                        st.error("Prestataire, bénéficiaire, description et montant > 0 obligatoires.")
                    else:
                        lettres = montant_en_lettres(e_montant)
                        try:
                            from openpyxl import load_workbook
                            if os.path.exists(FICHIER_PAIEMENTS):
                                wb = load_workbook(FICHIER_PAIEMENTS)
                                ws = wb[FEUILLE_DEPENSES] if FEUILLE_DEPENSES in wb.sheetnames else wb.active
                                hdr_row = None
                                for r in range(1, 15):
                                    if any(str(c.value or "").strip().lower() == "no de transaction" for c in ws[r]):
                                        hdr_row = r
                                        break
                                if hdr_row is None:
                                    hdr_row = 1
                                col_of = {str(c.value or "").strip().lower(): c.column for c in ws[hdr_row] if c.value}
                                col_trans = col_of.get("no de transaction", 2)
                                excel_row = None
                                for r in range(hdr_row + 1, ws.max_row + 1):
                                    if str(ws.cell(r, col_trans).value or "").strip() == str(trans_sel):
                                        excel_row = r
                                if excel_row:
                                    mapping = {
                                        "date": e_date.strip(),
                                        "type d'operation": e_type.strip(),
                                        "type d'opération": e_type.strip(),
                                        "produit": e_prod.strip(),
                                        "prestataires/fournisseurs": e_presta.strip(),
                                        "nom du bénéficiaire": e_benef.strip(),
                                        "nom du beneficiaire": e_benef.strip(),
                                        "num. pieces d'identite": e_piece.strip(),
                                        "date d'expiration": e_exp.strip(),
                                        "addresse /telephone": e_tel.strip(),
                                        "ref doc": e_ref.strip(),
                                        "description/objet du paiement": e_desc.strip(),
                                        "montant": e_montant,
                                        "observation": e_obs,
                                        "montant en lettre": lettres,
                                        "mode paiement": e_mode,
                                        "numero de cheque": e_cheque.strip() or "-",
                                        "banque": e_banque.strip(),
                                    }
                                    for h, col in col_of.items():
                                        if h in mapping:
                                            val = mapping[h]
                                            if h == "date" and isinstance(val, str):
                                                try:
                                                    val = datetime.strptime(val, "%d/%m/%Y")
                                                except ValueError:
                                                    pass
                                            ws.cell(excel_row, col, val if val != "" else None)
                                    wb.save(FICHIER_PAIEMENTS)
                                    st.success(f"✅ Paiement **{trans_sel}** modifié.")
                                    st.rerun()
                                else:
                                    st.error("Ligne introuvable dans Excel.")
                            else:
                                st.error("Fichier Excel introuvable.")
                        except Exception as e:
                            st.error(f"Erreur : {e}")

            st.markdown("##### 🗑️ Supprimer ce paiement")
            conf_del = st.checkbox(f"Je confirme la suppression de **{trans_sel}**", key="conf_del_pay")
            if conf_del and st.button("🗑️ Supprimer ce paiement", type="primary", key="btn_del_pay"):
                try:
                    from openpyxl import load_workbook
                    wb = load_workbook(FICHIER_PAIEMENTS)
                    ws = wb[FEUILLE_DEPENSES] if FEUILLE_DEPENSES in wb.sheetnames else wb.active
                    hdr_row = None
                    for r in range(1, 15):
                        if any(str(c.value or "").strip().lower() == "no de transaction" for c in ws[r]):
                            hdr_row = r
                            break
                    if hdr_row is None:
                        st.error("En-têtes introuvables.")
                    else:
                        col_of = {str(c.value or "").strip().lower(): c.column for c in ws[hdr_row] if c.value}
                        col_trans = col_of.get("no de transaction", 2)
                        rows_to_del = []
                        for r in range(hdr_row + 1, ws.max_row + 1):
                            if str(ws.cell(r, col_trans).value or "").strip() == str(trans_sel):
                                rows_to_del.append(r)
                        for r in reversed(rows_to_del):
                            ws.delete_rows(r, 1)
                        for t in list(ws.tables.values()):
                            try:
                                end_col = t.ref.split(":")[1].rstrip("0123456789")
                                new_end = max(hdr_row, ws.max_row)
                                t.ref = f"{t.ref.split(':')[0]}:{end_col}{new_end}"
                            except Exception:
                                pass
                        wb.save(FICHIER_PAIEMENTS)
                        st.success(f"✅ Paiement **{trans_sel}** supprimé ({len(rows_to_del)} ligne(s)).")
                        st.rerun()
                except Exception as e:
                    st.error(f"Erreur suppression : {e}")

# Annuaire : listes déroulantes alimentées par l'Excel (lecture seule)
st.markdown("---")
st.subheader("🔎 Annuaire prestataires (données issues de l'Excel)")
annuaire = charger_annuaire(df)
if not annuaire:
    st.info("Aucun prestataire dans l'Excel pour le moment.")
else:
    a1, a2, a3, a4, a5 = st.columns(5)
    with a1:
        presta_sel = st.selectbox("Prestataire / Fournisseur", list(annuaire.keys()), key="ann_presta")
    lignes_p = annuaire[presta_sel]
    with a2:
        benefs = _uniques_recents(lignes_p[COL_BENEF]) or ["—"]
        benef_sel = st.selectbox("Nom du bénéficiaire", benefs, key="ann_benef")
    lignes_b = lignes_p[lignes_p[COL_BENEF].map(_propre) == benef_sel] if benef_sel != "—" else lignes_p
    with a3:
        pieces = _uniques_recents(lignes_b[COL_PIECE]) or ["—"]
        st.selectbox("N° pièce d'identité", pieces, key="ann_piece")
    with a4:
        exps = _uniques_recents(lignes_b[COL_EXP]) or ["—"]
        st.selectbox("Date d'expiration", exps, key="ann_exp")
    with a5:
        tels = _uniques_recents(lignes_b[COL_TEL]) or ["—"]
        st.selectbox("Adresse / Téléphone", tels, key="ann_tel")
    st.caption(f"{len(annuaire)} prestataires · {len(lignes_p)} opération(s) pour « {presta_sel} ». "
               "Les listes se mettent à jour automatiquement à chaque nouvel enregistrement.")

# Régénérer un bon depuis une ligne existante
st.markdown("---")
st.subheader("📄 Générer / régénérer un bon de paiement Word")
if not df.empty:
    options = df["No de transaction"].astype(str).tolist()
    # dedupe preserve order
    seen = set()
    options_u = []
    for o in reversed(options):
        if o not in seen and o and o != "nan":
            seen.add(o)
            options_u.append(o)
    options_u = list(reversed(options_u))

    sel = st.selectbox("N° transaction", options=options_u)
    if st.button("📄 Générer le bon Word", type="primary"):
        lignes = df[df["No de transaction"].astype(str) == str(sel)]
        if lignes.empty:
            st.error("Transaction introuvable.")
        else:
            # Si plusieurs lignes même N°, prendre la première / sommer ?
            row = lignes.iloc[-1]
            try:
                montant = float(pd.to_numeric(row["Montant"], errors="coerce") or 0)
            except Exception:
                montant = 0
            lettres = str(row.get("Montant en lettre", "") or "") or montant_en_lettres(montant)
            data = {
                "Date": str(row.get("Date", "")),
                "No_transaction": str(sel),
                "Type_operation": str(row.get("Type d'operation", "")),
                "Produit": str(row.get("Produit", "")),
                "Prestataire": str(row.get("Prestataires/Fournisseurs", "")),
                "Beneficiaire": str(row.get("Nom du bénéficiaire", "")),
                "Piece_identite": str(row.get("Num. Pieces d'identite", "")),
                "Date_expiration": str(row.get("Date d'expiration", "")),
                "Telephone": str(row.get("Addresse /Telephone", "")),
                "Ref_doc": str(row.get("Ref doc", "")),
                "Description": str(row.get("Description/Objet du paiement", "")),
                "Montant": montant,
                "Observation": str(row.get("Observation", "")),
                "Montant_lettres": lettres,
                "Mode_paiement": str(row.get("Mode paiement", "")),
                "Numero_cheque": str(row.get("Numero de cheque", "")),
                "Banque": str(row.get("Banque", "")),
                "Utilisateur": name,
                "N_BC": str(row.get("N_BC", "") or ""),
                "N_Dechargement": str(row.get("N_Dechargement", "") or ""),
            }
            try:
                from generer_bon_paiement import generer_bon_paiement
                chemin = generer_bon_paiement(data)
                st.success(f"Bon généré : `{chemin}`")
                tb.afficher_apres_generation(chemin, "nouveau_" + os.path.basename(chemin))
            except Exception as e:
                st.error(f"Erreur : {e}")
                import traceback
                st.code(traceback.format_exc())

# Liste des bons générés
st.markdown("---")
st.subheader("📥 Bons de paiement disponibles (Word et PDF)")
import glob
tb.afficher_moteur()
tb.afficher_erreur_pdf()
fichiers = sorted(glob.glob(os.path.join(DOSSIER_BONS, "Bon_Paiement_*.docx")), reverse=True)
if not fichiers:
    st.info("Aucun bon dans le dossier bons_paiement/.")
else:
    tb.afficher_tout_creer(fichiers[:30], "bp")
    for fpath in fichiers[:30]:
        c1, c2, c3 = st.columns([4, 1, 1])
        with c1:
            st.write(f"📄 **{os.path.basename(fpath)}**")
        tb.afficher_boutons(fpath, f"bp_{fpath}", c2, c3)

st.caption("NAP SARL — Suivi paiements + génération bon de paiement (QR) · Placez `Suivi de paiement.xlsx` à côté de l'app")
