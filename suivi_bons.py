"""
NAP SARL — Suivi des BONS DE COMMANDE : émis / déchargés / reste à décharger + recherche d'un camion.
À placer à la RACINE du projet (à côté de dechargement_data.py). Utilisé par la page 2_Bons_de_Commande.py

Un bon est :
  EN ATTENTE  = aucun camion encore déchargé
  EN COURS    = livraison commencée mais pas terminée
  LIVRE       = poids déchargé >= 98 % de la quantité commandée  (seuil : dd.SEUIL_LIVRE)
"""
import os
import re
from datetime import date

import pandas as pd

import dechargement_data as dd

VERSION_SUIVI = 1
JOURS_ALERTE = 3      # un bon non déchargé depuis plus de N jours est en RETARD (alerte rouge)


# ---------- Outils ----------
def normaliser_plaque(texte):
    """'be 5458-rb / AY0833 RB' -> 'BE5458RBAY0833RB' (espaces, tirets et barres ignorés)."""
    return re.sub(r"[^A-Z0-9]", "", dd._s(texte).upper())


def _produit(reference):
    m = re.search(r"-BC-([A-Za-z]+)", dd._s(reference))
    return m.group(1).upper() if m else ""


def _numero_bon(reference):
    m = re.match(r"\s*(\d+)", dd._s(reference))
    return int(m.group(1)) if m else None


# ---------- Chargement ----------
def charger_bons_tous():
    """Tous les bons de commande (tous produits), avec date, produit et quantité en kg."""
    cols = ["Reference", "Date_Emission", "Heure_Creation", "Utilisateur", "Nom_Fournisseur", "Quantite",
            "Numero_Tracteur", "Chauffeur"]
    if os.path.exists(dd.FICHIER_BONS):
        df = pd.read_excel(dd.FICHIER_BONS, dtype=str)
    else:
        df = pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].map(dd._s)
    df = df[df["Reference"] != ""].copy()
    out = pd.DataFrame({
        "Reference": df["Reference"],
        "Date": pd.to_datetime(df["Date_Emission"], dayfirst=True, errors="coerce"),
        "Heure": df["Heure_Creation"],
        "Produit": [_produit(r) for r in df["Reference"]],
        "Fournisseur": df["Nom_Fournisseur"], "Quantite": df["Quantite"],
        "Qte_KG": [dd.quantite_kg(q) for q in df["Quantite"]],
        "Camion_Bon": df["Numero_Tracteur"], "Chauffeur": df["Chauffeur"], "Cree_Par": df["Utilisateur"],
    })
    return out.reset_index(drop=True)


def charger_dechargements_tous():
    """Tous les déchargements (tous produits). Le produit actif de la session est restauré."""
    ancien = dd.get_produit()
    frames = []
    try:
        for p in dd.PRODUITS:
            dd.set_produit(p)
            if os.path.exists(dd.fichier_actif()):
                d = dd.charger_dechargements()
                if not d.empty:
                    frames.append(d.assign(Produit=p))
    finally:
        dd.set_produit(ancien)
    if frames:
        return pd.concat(frames, ignore_index=True)
    return pd.DataFrame(columns=dd.COLS_DECH + ["Produit"])


# ---------- Suivi par bon ----------
def suivi_par_bon(bons, dech, produit=None, aujourdhui=None):
    """Une ligne par bon : quantité commandée, déchargée, reste, statut, ancienneté."""
    aujourdhui = aujourdhui or date.today()
    b = bons if not produit else bons[bons["Produit"] == produit]
    d = dech[dech["Ref_BC"] != ""]
    if d.empty:
        g = pd.DataFrame(columns=["Nb_Camions", "Livre_KG", "Dernier_Dech"], index=pd.Index([], dtype="object"))
    else:
        g = d.groupby("Ref_BC").agg(Nb_Camions=("Transaction", "count"), Livre_KG=("Poids_Net_KG", "sum"),
                                    Dernier_Dech=("Date_Dechargement", "max"))
    m = b.merge(g, left_on="Reference", right_index=True, how="left")
    m["Nb_Camions"] = pd.to_numeric(m["Nb_Camions"], errors="coerce").fillna(0).astype(int)
    m["Livre_KG"] = pd.to_numeric(m["Livre_KG"], errors="coerce").fillna(0.0)
    m["Reste_KG"] = [max(q - l, 0.0) if q > 0 else 0.0 for q, l in zip(m["Qte_KG"], m["Livre_KG"])]

    def statut(r):
        if r["Nb_Camions"] == 0:
            return "EN ATTENTE"
        if r["Qte_KG"] <= 0 or r["Livre_KG"] >= dd.SEUIL_LIVRE * r["Qte_KG"]:
            return "LIVRE"
        return "EN COURS"

    m["Statut"] = [statut(r) for _, r in m.iterrows()] if not m.empty else []
    m["Age_Jours"] = [(pd.Timestamp(aujourdhui) - d_).days if pd.notna(d_) else None for d_ in m["Date"]]
    m["Qte_T"], m["Livre_T"], m["Reste_T"] = m["Qte_KG"] / 1000, m["Livre_KG"] / 1000, m["Reste_KG"] / 1000
    return m.sort_values(["Date", "Reference"], ascending=[False, False]).reset_index(drop=True)


def indicateurs(suivi, aujourdhui=None):
    """Chiffres du mini tableau de bord."""
    aujourdhui = aujourdhui or date.today()
    a_decharger = suivi[suivi["Statut"] != "LIVRE"]
    en_retard = a_decharger[dd.masque(a_decharger, [(a is not None and not pd.isna(a) and a > JOURS_ALERTE)
                                                    for a in a_decharger["Age_Jours"]])]
    return {
        "emis": len(suivi),
        "emis_aujourdhui": int(sum(1 for d_ in suivi["Date"] if pd.notna(d_) and d_.date() == aujourdhui)),
        "decharges": int((suivi["Nb_Camions"] > 0).sum()),
        "livres": int((suivi["Statut"] == "LIVRE").sum()),
        "camions": int(suivi["Nb_Camions"].sum()),
        "a_decharger": len(a_decharger),
        "en_retard": len(en_retard),
        "qte_T": float(suivi["Qte_T"].sum()),
        "livre_T": float(suivi["Livre_T"].sum()),
        "reste_T": float(suivi["Reste_T"].sum()),
    }


def bons_a_decharger(suivi):
    """Bons pas encore (entièrement) déchargés, du plus ancien au plus récent."""
    a = suivi[suivi["Statut"] != "LIVRE"].copy()
    a = a.sort_values("Age_Jours", ascending=False, na_position="last")

    def alerte(age):
        if age is None or pd.isna(age):
            return ""
        return "🔴 En retard" if age > JOURS_ALERTE else ("🟠 À suivre" if age >= 1 else "🟢 Du jour")

    a.insert(0, "Alerte", [alerte(x) for x in a["Age_Jours"]])
    return a


# ---------- Recherche d'un camion / d'un bon ----------
def rechercher(texte, bons, dech):
    """Cherche un N° de camion (ou de bon) dans les bons de commande ET les déchargements.
    -> {"bons": df, "dech": df, "requete": ...} ou None si la saisie est trop courte."""
    q = normaliser_plaque(texte)
    if len(q) < 2:
        return None
    num = int(q) if q.isdigit() and len(q) <= 6 else None

    def bon_correspond(ref):
        n = _numero_bon(ref)
        return (num is not None and n == num) or normaliser_plaque(ref).startswith(q)

    # 1) la saisie est un numéro de bon (ex. 0013) : on ne montre que ce bon et ses déchargements
    par_numero = [bon_correspond(r) for r in bons["Reference"]]
    if any(par_numero):
        r_bons = bons[dd.masque(bons, par_numero)]
        refs = set(r_bons["Reference"])
        return {"bons": r_bons.reset_index(drop=True),
                "dech": dech[dech["Ref_BC"].isin(refs)].reset_index(drop=True), "requete": q}

    # 2) sinon : c'est un numéro de camion (plaque, même partielle)
    m_dech = [q in normaliser_plaque(c) or q in normaliser_plaque(cb) or q in normaliser_plaque(tk)
              for c, cb, tk in zip(dech["N_Camion"], dech["Camion_BC"], dech["Weighment_Ticket"])]
    r_dech = dech[dd.masque(dech, m_dech)]
    # un bon est trouvé si le camion y est prévu OU s'il a livré pour ce bon (même à la place du camion prévu)
    m_bons = [q in normaliser_plaque(c) for c in bons["Camion_Bon"]]
    r_bons = bons[dd.masque(bons, m_bons)]
    liees = bons[bons["Reference"].isin(set(r_dech["Ref_BC"]))]
    r_bons = pd.concat([r_bons, liees]).drop_duplicates(subset="Reference")
    # déchargements de ces bons réalisés par un autre camion : utiles pour comprendre
    r_dech = pd.concat([r_dech, dech[dech["Ref_BC"].isin(set(r_bons["Reference"]))]]).drop_duplicates(
        subset=["Produit", "Transaction"])
    return {"bons": r_bons.reset_index(drop=True), "dech": r_dech.reset_index(drop=True), "requete": q}
