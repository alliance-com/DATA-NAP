"""
NAP SARL — Liste des FOURNISSEURS (fournisseurs.xlsx)
  - création automatique d'un fournisseur (code FRS00001, FRS00002…) quand un bon de commande est créé pour un nom inconnu
  - mise à jour des informations (nom, téléphone, commune, actif/inactif)
  - soldes par fournisseur (camions, dû, payé, reste à payer) tous produits confondus

À placer à la RACINE du projet (à côté de dechargement_data.py).
Le fichier fournisseurs.xlsx existant est respecté : toutes ses colonnes et feuilles sont conservées.
"""
import os
import re
import unicodedata
from datetime import datetime

import pandas as pd

import dechargement_data as dd

VERSION_FRS = 2

FICHIER = dd.FICHIER_FRS
COL_CODE, COL_NOM, COL_ACTIF = "Code", "Nom_Fournisseur", "Actif"
ALIAS_TEL = ["Telephone", "Téléphone", "Tel", "Telephone_Fournisseur", "Contact"]
ALIAS_COMMUNE = ["Commune", "Commune_Origine", "Ville", "Localite", "Localité"]
EXTRAS = ["Date_Creation", "Cree_Par", "Date_Modification", "Modifie_Par"]


# ---------- Outils ----------
def _norm(t):
    """Nom comparable : sans accents, majuscules, espaces simples."""
    t = unicodedata.normalize("NFKD", dd._s(t)).encode("ascii", "ignore").decode().upper()
    return re.sub(r"\s+", " ", t).strip()


def nom_propre(t):
    """Nom enregistré : espaces simples + majuscules (comme les noms déjà présents)."""
    return re.sub(r"\s+", " ", dd._s(t)).upper()


def _trouver(df, alias, defaut):
    for a in alias:
        for c in df.columns:
            if str(c).strip().lower() == a.lower():
                return c
    return defaut


def _maintenant():
    return datetime.now().strftime("%d/%m/%Y %H:%M")


# ---------- Lecture / écriture (toutes colonnes et feuilles conservées) ----------
def _charger():
    """-> (df complet, colonnes utiles, toutes les feuilles, nom de la feuille principale)"""
    if os.path.exists(FICHIER):
        feuilles = pd.read_excel(FICHIER, sheet_name=None, dtype=str)
        nom0 = next(iter(feuilles))
        df = feuilles[nom0].copy()
    else:
        nom0 = "Fournisseurs"
        df = pd.DataFrame(columns=[COL_CODE, COL_NOM, "Telephone", "Commune", COL_ACTIF] + EXTRAS)
        feuilles = {nom0: df}
    for c in df.columns:
        df[c] = df[c].map(dd._s)
    for c in (COL_CODE, COL_NOM):
        if c not in df.columns:
            df[c] = ""
    if COL_ACTIF not in df.columns:
        df[COL_ACTIF] = "OUI"
    cols = {"tel": _trouver(df, ALIAS_TEL, "Telephone"), "commune": _trouver(df, ALIAS_COMMUNE, "Commune")}
    for c in list(cols.values()) + EXTRAS:
        if c not in df.columns:
            df[c] = ""
    return df, cols, feuilles, nom0


def _sauver(df, feuilles, nom0):
    feuilles = dict(feuilles)
    feuilles[nom0] = df
    tmp = "fournisseurs.tmp.xlsx"
    with pd.ExcelWriter(tmp, engine="openpyxl") as w:
        for nom, d in feuilles.items():
            d.to_excel(w, sheet_name=str(nom)[:31], index=False)
            ws = w.sheets[str(nom)[:31]]
            for col in ws.columns:
                largeur = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(12, largeur + 2), 45)
            ws.freeze_panes = "A2"
    os.replace(tmp, FICHIER)


def charger_fournisseurs(actifs_seulement=False):
    """Vue simple : Code, Nom, Telephone, Commune, Actif (triée par nom)."""
    df, cols, _, _ = _charger()
    v = pd.DataFrame({"Code": df[COL_CODE], "Nom": df[COL_NOM], "Telephone": df[cols["tel"]],
                      "Commune": df[cols["commune"]], "Actif": df[COL_ACTIF].map(lambda x: dd._s(x).upper()),
                      "Cree_Par": df["Cree_Par"], "Date_Creation": df["Date_Creation"],
                      "Modifie_Par": df["Modifie_Par"], "Date_Modification": df["Date_Modification"]})
    v = v[(v["Code"] != "") & (v["Nom"] != "")]
    if actifs_seulement:
        v = v[v["Actif"] == "OUI"]
    return v.sort_values("Nom", key=lambda s: s.map(_norm)).reset_index(drop=True)


def options_selection():
    """{ 'FRS00006 — NOM' : (code, nom) } des fournisseurs actifs (pour les listes déroulantes)."""
    v = charger_fournisseurs(actifs_seulement=True)
    return {dd.label_fournisseur(r.Code, r.Nom): (r.Code, r.Nom) for r in v.itertuples()}


def prochain_code(df):
    nums = [int(m.group(1)) for c in df[COL_CODE] for m in [re.search(r"(\d+)$", c)] if m]
    return f"FRS{(max(nums) + 1) if nums else 1:05d}"


# ---------- Recherche ----------
def chercher_exact(nom):
    """(code, nom) du fournisseur portant ce nom (accents/majuscules/espaces ignorés), sinon None."""
    cible = _norm(nom)
    if not cible:
        return None
    df, _, _, _ = _charger()
    for code, n in zip(df[COL_CODE], df[COL_NOM]):
        if code and _norm(n) == cible:
            return code, n
    return None


def fournisseurs_proches(nom, limite=5):
    """Noms voisins (hors nom identique) : pour éviter de créer deux fois le même fournisseur."""
    cible = _norm(nom)
    if not cible:
        return []
    df, _, _, _ = _charger()
    out = [(c, n) for c, n in zip(df[COL_CODE], df[COL_NOM]) if c and _norm(n) != cible and dd.meme_nom(n, nom)]
    return out[:limite]


# ---------- Création ----------
def creer_fournisseur(nom, telephone="", commune="", utilisateur=""):
    """-> (ok, message, code)"""
    nom = nom_propre(nom)
    if not nom:
        return False, "Le nom du fournisseur est obligatoire.", ""
    existant = chercher_exact(nom)
    if existant:
        return False, f"Ce fournisseur existe déjà : {existant[0]} — {existant[1]}.", existant[0]
    df, cols, feuilles, nom0 = _charger()
    code = prochain_code(df)
    ligne = {c: "" for c in df.columns}
    ligne.update({COL_CODE: code, COL_NOM: nom, cols["tel"]: dd._s(telephone), cols["commune"]: dd._s(commune),
                  COL_ACTIF: "OUI", "Date_Creation": _maintenant(), "Cree_Par": utilisateur})
    _sauver(pd.concat([df, pd.DataFrame([ligne])], ignore_index=True), feuilles, nom0)
    return True, f"Fournisseur {code} — {nom} ajouté à la liste.", code


def assurer_fournisseur(nom, telephone="", commune="", utilisateur=""):
    """Retourne (code, nom, cree). Le fournisseur est créé automatiquement s'il n'est pas dans la liste."""
    existant = chercher_exact(nom)
    if existant:
        return existant[0], existant[1], False
    ok, msg, code = creer_fournisseur(nom, telephone, commune, utilisateur)
    return code, nom_propre(nom), ok


# ---------- Mise à jour ----------
def _propager_nom(code, ancien, nouveau):
    """Répercute un changement de nom sur les bons de commande et les déchargements (tous produits)."""
    nb_bons = nb_dech = 0
    if os.path.exists(dd.FICHIER_BONS):
        b = pd.read_excel(dd.FICHIER_BONS)
        if "Nom_Fournisseur" in b.columns and not b.empty:
            codes = b["Code_Fournisseur"].map(dd._s) if "Code_Fournisseur" in b.columns else pd.Series([""] * len(b))
            masque = (codes == code) | b["Nom_Fournisseur"].map(lambda x: _norm(x) == _norm(ancien))
            if masque.any():
                b.loc[masque, "Nom_Fournisseur"] = nouveau
                if "Code_Fournisseur" in b.columns:
                    b.loc[masque & (codes == ""), "Code_Fournisseur"] = code
                b.to_excel(dd.FICHIER_BONS, index=False)
                nb_bons = int(masque.sum())
    ancien_produit = dd.get_produit()
    try:
        for p in dd.PRODUITS:
            dd.set_produit(p)
            if not os.path.exists(dd.fichier_actif()):
                continue
            dech, aff = dd.charger_dechargements(), dd.charger_affectations()
            masque = (dech["Code_Frs"] == code) | ((dech["Code_Frs"] == "") &
                                                   dech["Fournisseur"].map(lambda x: _norm(x) == _norm(ancien)))
            if masque.any():
                dech.loc[masque, "Fournisseur"] = nouveau
                dech.loc[masque, "Code_Frs"] = code
                dd.sauver(dech, aff)
                nb_dech += int(masque.sum())
    finally:
        dd.set_produit(ancien_produit)
    return nb_bons, nb_dech


def modifier_fournisseur(code, nom=None, telephone=None, commune=None, actif=None, utilisateur="",
                         propager_nom=True):
    """Met à jour les informations d'un fournisseur. -> (ok, message)"""
    df, cols, feuilles, nom0 = _charger()
    idx = df.index[df[COL_CODE] == code]
    if len(idx) == 0:
        return False, "Fournisseur introuvable."
    i = idx[0]
    ancien_nom = df.at[i, COL_NOM]
    changements = []
    nouveau_nom = nom_propre(nom) if nom is not None else ancien_nom
    if nouveau_nom != ancien_nom:
        if not nouveau_nom:
            return False, "Le nom du fournisseur ne peut pas être vide."
        autre = df[(df[COL_CODE] != code) & (df[COL_NOM].map(_norm) == _norm(nouveau_nom))]
        if not autre.empty:
            return False, f"Un autre fournisseur porte déjà ce nom : {autre.iloc[0][COL_CODE]}."
        df.at[i, COL_NOM] = nouveau_nom
        changements.append("nom")
    for champ, valeur, col in (("téléphone", telephone, cols["tel"]), ("commune", commune, cols["commune"])):
        if valeur is not None and dd._s(valeur) != df.at[i, col]:
            df.at[i, col] = dd._s(valeur)
            changements.append(champ)
    if actif is not None and ("OUI" if actif else "NON") != df.at[i, COL_ACTIF].upper():
        df.at[i, COL_ACTIF] = "OUI" if actif else "NON"
        changements.append("statut actif")
    if not changements:
        return False, "Aucun changement à enregistrer."
    df.at[i, "Date_Modification"], df.at[i, "Modifie_Par"] = _maintenant(), utilisateur
    _sauver(df, feuilles, nom0)
    msg = f"{code} mis à jour ({', '.join(changements)})."
    if "nom" in changements and propager_nom:
        nb_bons, nb_dech = _propager_nom(code, ancien_nom, nouveau_nom)
        msg += f" Nom répercuté sur {nb_bons} bon(s) de commande et {nb_dech} déchargement(s)."
    return True, msg


# ---------- Initialisation à partir des données existantes ----------
def _fournisseurs_trouves():
    """Fournisseurs déjà présents dans les déchargements (avec leur code FRS) puis dans les bons de commande."""
    trouves = []
    ancien = dd.get_produit()
    try:
        for p in dd.PRODUITS:
            dd.set_produit(p)
            if os.path.exists(dd.fichier_actif()):
                d = dd.charger_dechargements()
                for code, nom in d[["Code_Frs", "Fournisseur"]].drop_duplicates().itertuples(index=False):
                    if dd._s(nom):
                        trouves.append({"code": dd._s(code), "nom": nom_propre(nom), "tel": "", "commune": ""})
    finally:
        dd.set_produit(ancien)
    if os.path.exists(dd.FICHIER_BONS):
        b = pd.read_excel(dd.FICHIER_BONS, dtype=str)
        for c in ("Code_Fournisseur", "Nom_Fournisseur", "Telephone_Fournisseur", "Commune_Origine"):
            if c not in b.columns:
                b[c] = ""
        for r in b.itertuples(index=False):
            if dd._s(r.Nom_Fournisseur):
                trouves.append({"code": dd._s(r.Code_Fournisseur), "nom": nom_propre(r.Nom_Fournisseur),
                                "tel": dd._s(r.Telephone_Fournisseur), "commune": dd._s(r.Commune_Origine)})
    return trouves


def completer_depuis_donnees(utilisateur=""):
    """Ajoute à la liste les fournisseurs trouvés dans les déchargements et les bons de commande (sans doublon),
    puis renseigne le code FRS des bons qui n'en ont pas.
    -> {"ajoutes": n, "codes_bons": m, "a_verifier": [(code1, nom1, code2, nom2), ...]}"""
    df, cols, feuilles, nom0 = _charger()
    par_nom = {_norm(n): c for c, n in zip(df[COL_CODE], df[COL_NOM]) if c}
    codes = {c for c in df[COL_CODE] if c}
    nouveaux = []

    def ajouter(code, nom, tel="", commune=""):
        nonlocal df
        ligne = {c: "" for c in df.columns}
        ligne.update({COL_CODE: code, COL_NOM: nom, cols["tel"]: tel, cols["commune"]: commune, COL_ACTIF: "OUI",
                      "Date_Creation": _maintenant(), "Cree_Par": f"{utilisateur} (création automatique)"})
        df = pd.concat([df, pd.DataFrame([ligne])], ignore_index=True)
        par_nom[_norm(nom)] = code
        codes.add(code)
        nouveaux.append(code)

    trouves = _fournisseurs_trouves()
    # 1) ceux qui ont déjà un code FRS (déchargements, anciens bons) : on garde leur code
    for t in trouves:
        if t["code"] and t["code"] not in codes and _norm(t["nom"]) not in par_nom:
            ajouter(t["code"], t["nom"], t["tel"], t["commune"])
    # 2) les autres : nouveau code à la suite
    for t in trouves:
        if _norm(t["nom"]) not in par_nom:
            ajouter(prochain_code(df), t["nom"], t["tel"], t["commune"])
    if nouveaux:
        _sauver(df, feuilles, nom0)

    # doublons possibles à vérifier (noms proches, codes différents)
    a_verifier = []
    for c in nouveaux:
        n = df.loc[df[COL_CODE] == c, COL_NOM].iloc[0]
        for c2, n2 in zip(df[COL_CODE], df[COL_NOM]):
            if c2 and c2 != c and dd.meme_nom(n, n2) and (c2, n2, c, n) not in a_verifier:
                a_verifier.append((c, n, c2, n2))

    # code FRS renseigné sur les bons de commande qui n'en ont pas
    codes_bons = 0
    if os.path.exists(dd.FICHIER_BONS) and par_nom:
        b = pd.read_excel(dd.FICHIER_BONS)
        if "Nom_Fournisseur" in b.columns and not b.empty:
            if "Code_Fournisseur" not in b.columns:
                b["Code_Fournisseur"] = ""
            actuel = b["Code_Fournisseur"].map(dd._s)
            voulu = b["Nom_Fournisseur"].map(lambda x: par_nom.get(_norm(x), ""))
            masque = (actuel == "") & (voulu != "")
            if masque.any():
                b.loc[masque, "Code_Fournisseur"] = voulu[masque]
                b.to_excel(dd.FICHIER_BONS, index=False)
                codes_bons = int(masque.sum())
    return {"ajoutes": len(nouveaux), "codes_bons": codes_bons, "a_verifier": a_verifier}


def doublons_possibles():
    """Paires de fournisseurs aux noms voisins : [(code1, nom1, code2, nom2), ...] (à vérifier à la main)."""
    df, _, _, _ = _charger()
    lignes = [(c, n) for c, n in zip(df[COL_CODE], df[COL_NOM]) if c and n]
    out = []
    for i, (c1, n1) in enumerate(lignes):
        for c2, n2 in lignes[i + 1:]:
            if dd.meme_nom(n1, n2):
                out.append((c1, n1, c2, n2))
    return out


def initialiser_si_absent(utilisateur=""):
    """Si fournisseurs.xlsx n'existe pas encore : le crée à partir des données existantes. -> résumé ou None"""
    if os.path.exists(FICHIER):
        return None
    return completer_depuis_donnees(utilisateur)


# ---------- Soldes ----------
def soldes_par_fournisseur():
    """Code, Camions, Montant_Du, Montant_Paye, Reste_A_Payer — tous produits confondus."""
    ancien = dd.get_produit()
    frames = []
    try:
        for p in dd.PRODUITS:
            dd.set_produit(p)
            if os.path.exists(dd.fichier_actif()):
                d, _ = dd.charger_complet()
                if not d.empty:
                    frames.append(d)
    finally:
        dd.set_produit(ancien)
    if not frames:
        return pd.DataFrame(columns=["Code", "Camions", "Montant_Du", "Montant_Paye", "Reste_A_Payer"])
    d = pd.concat(frames, ignore_index=True)
    d["Cle"] = [c or n for c, n in zip(d["Code_Frs"], d["Fournisseur"])]
    g = d.groupby("Cle").agg(Camions=("Transaction", "count"), Montant_Du=("Montant_Du", "sum"),
                             Montant_Paye=("Montant_Paye", "sum")).reset_index().rename(columns={"Cle": "Code"})
    g["Reste_A_Payer"] = g["Montant_Du"] - g["Montant_Paye"]
    return g
