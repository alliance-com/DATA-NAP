# -*- coding: utf-8 -*-
"""
URHC Jeunesse — Présences (données)
Fichier Excel à côté de ce module : presence_urhc.xlsx
"""
from __future__ import annotations

import hashlib
import os
import re
from datetime import datetime, date, timedelta
from io import BytesIO
from pathlib import Path

import pandas as pd

# Toujours le dossier du projet (pas le cwd Streamlit)
_BASE = Path(__file__).resolve().parent
FICHIER = str(_BASE / "presence_urhc.xlsx")

FEUILLE_MEMBRES = "MEMBRES"
FEUILLE_PRESENCES = "PRESENCES"

COLS_MEMBRES = [
    "ID", "Nom", "Prenom", "Telephone", "Email", "Profession",
    "Est_Etudiant", "Universite", "Faculte", "Filiere", "Annee_Etude",
    "Role", "Actif", "Mot_de_passe_hash", "Date_Inscription",
]
COLS_PRESENCES = [
    "ID", "ID_Membre", "Date_Dimanche", "Statut",
    "Signale_Le", "Valide_Le", "Valide_Par", "Notes",
]

ROLES = ["MEMBRE", "ADMIN"]


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


def _hash(pwd: str) -> str:
    return hashlib.sha256(("urhc|" + (pwd or "")).encode("utf-8")).hexdigest()


def normaliser_tel(tel: str) -> str:
    t = re.sub(r"\D", "", _s(tel))
    if len(t) == 11 and t.startswith("0"):
        t = t[1:]
    return t


def tel_variantes(tel: str):
    t = normaliser_tel(tel)
    out = {t}
    if len(t) == 10 and t.startswith("0"):
        out.add(t[1:])
    if len(t) == 9:
        out.add("0" + t)
    if len(t) == 10 and not t.startswith("0"):
        out.add("0" + t)
    return {x for x in out if x}


def tel_valide(tel: str) -> bool:
    t = normaliser_tel(tel)
    return t.isdigit() and len(t) in (9, 10)


def email_valide(email: str) -> bool:
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", _s(email).lower()))


def ensure_file():
    if os.path.exists(FICHIER):
        # migrer colonnes manquantes
        try:
            df = pd.read_excel(FICHIER, sheet_name=FEUILLE_MEMBRES)
            changed = False
            for c in COLS_MEMBRES:
                if c not in df.columns:
                    df[c] = ""
                    changed = True
            if changed:
                _ecrire(FEUILLE_MEMBRES, df[COLS_MEMBRES])
        except Exception:
            pass
        return
    with pd.ExcelWriter(FICHIER, engine="openpyxl") as w:
        pd.DataFrame(columns=COLS_MEMBRES).to_excel(w, sheet_name=FEUILLE_MEMBRES, index=False)
        pd.DataFrame(columns=COLS_PRESENCES).to_excel(w, sheet_name=FEUILLE_PRESENCES, index=False)


def _lire(nom, cols=None):
    ensure_file()
    try:
        df = pd.read_excel(FICHIER, sheet_name=nom)
    except Exception:
        return pd.DataFrame(columns=cols or [])
    if cols:
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[[c for c in cols if c in df.columns]]
    return df


def _ecrire(nom, df):
    ensure_file()
    sheets = {}
    try:
        with pd.ExcelFile(FICHIER) as xls:
            sheets = {s: xls.parse(s) for s in xls.sheet_names}
    except Exception:
        sheets = {}
    sheets[nom] = df
    if FEUILLE_MEMBRES not in sheets:
        sheets[FEUILLE_MEMBRES] = pd.DataFrame(columns=COLS_MEMBRES)
    if FEUILLE_PRESENCES not in sheets:
        sheets[FEUILLE_PRESENCES] = pd.DataFrame(columns=COLS_PRESENCES)
    tmp = FICHIER.replace(".xlsx", ".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as w:
        for s, d in sheets.items():
            d.to_excel(w, sheet_name=s, index=False)
    os.replace(tmp, FICHIER)


def charger_membres(actifs_only=False):
    df = _lire(FEUILLE_MEMBRES, COLS_MEMBRES)
    if df.empty:
        return df
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
    for c in COLS_MEMBRES:
        if c not in ("ID",):
            df[c] = df[c].map(_s)
    df["Telephone"] = df["Telephone"].map(lambda x: normaliser_tel(x) if _s(x) else "")
    df["Email"] = df["Email"].str.lower()
    if actifs_only:
        df = df[df["Actif"].str.upper() != "NON"]
    return df.sort_values(["Nom", "Prenom"]).reset_index(drop=True)


def prochain_id_membre(df=None):
    df = df if df is not None else charger_membres()
    return 1 if df.empty else int(df["ID"].max()) + 1


def inscrire_membre(
    nom, prenom, telephone, email,
    profession="", est_etudiant=False,
    universite="", faculte="", filiere="", annee_etude="",
    role="MEMBRE", mot_de_passe="",
):
    if not _s(nom) or not _s(prenom):
        return False, "Nom et prénom obligatoires.", None
    if not tel_valide(telephone):
        return False, "Téléphone invalide (9 ou 10 chiffres).", None
    if not email_valide(email):
        return False, "Email invalide.", None
    if not _s(profession):
        return False, "Profession obligatoire.", None
    if est_etudiant:
        if not _s(universite) or not _s(filiere):
            return False, "Pour un étudiant : université et filière obligatoires.", None

    tel = normaliser_tel(telephone)
    email = _s(email).lower()
    df = charger_membres()
    if not df.empty:
        if df["Telephone"].isin(tel_variantes(tel)).any():
            return False, "Ce numéro est déjà inscrit. Allez à la connexion.", None
        if (df["Email"].str.lower() == email).any():
            return False, "Cet email est déjà inscrit. Allez à la connexion.", None

    role = _s(role).upper() or "MEMBRE"
    if role not in ROLES:
        role = "MEMBRE"
    # Premier membre = ADMIN
    if df.empty or not (df["Role"].str.upper() == "ADMIN").any():
        role = "ADMIN"

    pwd = _s(mot_de_passe) or tel[-4:]
    row = {
        "ID": prochain_id_membre(df),
        "Nom": _s(nom).upper(),
        "Prenom": _s(prenom).title(),
        "Telephone": tel,
        "Email": email,
        "Profession": _s(profession),
        "Est_Etudiant": "OUI" if est_etudiant else "NON",
        "Universite": _s(universite) if est_etudiant else "",
        "Faculte": _s(faculte) if est_etudiant else "",
        "Filiere": _s(filiere) if est_etudiant else "",
        "Annee_Etude": _s(annee_etude) if est_etudiant else "",
        "Role": role,
        "Actif": "OUI",
        "Mot_de_passe_hash": _hash(pwd),
        "Date_Inscription": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    _ecrire(FEUILLE_MEMBRES, df[COLS_MEMBRES])
    msg = (
        f"Inscription réussie — {row['Prenom']} {row['Nom']}. "
        f"Code de connexion : {pwd}"
    )
    return True, msg, int(row["ID"])


def authentifier(telephone, email, mot_de_passe=""):
    if not tel_valide(telephone):
        return None, "Téléphone invalide (9 ou 10 chiffres)."
    if not email_valide(email):
        return None, "Email invalide."
    email = _s(email).lower()
    variants = tel_variantes(telephone)
    df = charger_membres(actifs_only=True)
    if df.empty:
        return None, "Aucun membre inscrit. Utilisez la page Inscription."

    m = df[df["Telephone"].isin(variants) & (df["Email"].str.lower() == email)]
    if m.empty:
        par_tel = df[df["Telephone"].isin(variants)]
        par_mail = df[df["Email"].str.lower() == email]
        if not par_tel.empty:
            return None, "Téléphone reconnu, mais l’email ne correspond pas."
        if not par_mail.empty:
            return None, "Email reconnu, mais le numéro ne correspond pas."
        return None, "Aucun compte pour ce téléphone + email."

    row = m.iloc[0]
    tel_stocke = _s(row["Telephone"])
    hash_ok = _s(row["Mot_de_passe_hash"])
    candidats = []
    if _s(mot_de_passe):
        candidats.append(_s(mot_de_passe))
    candidats.append(normaliser_tel(telephone)[-4:])
    if tel_stocke:
        candidats.append(tel_stocke[-4:])
    # uniques
    vu, cands = set(), []
    for c in candidats:
        if c and c not in vu:
            vu.add(c)
            cands.append(c)
    if hash_ok and not any(hash_ok == _hash(c) for c in cands):
        return None, "Mot de passe incorrect (code d’inscription ou 4 derniers chiffres du tél.)."
    return row.to_dict(), None


def dimanches_du_mois(annee, mois):
    d = date(annee, mois, 1)
    while d.weekday() != 6:
        d += timedelta(days=1)
        if d.month != mois:
            return []
    out = []
    while d.month == mois:
        out.append(d)
        d += timedelta(days=7)
    return out


def dimanche_courant_ou_dernier(ref=None):
    ref = ref or date.today()
    if ref.weekday() == 6:
        return ref
    d = ref
    while d.weekday() != 6:
        d -= timedelta(days=1)
    return d


def prochain_dimanche(ref=None):
    ref = ref or date.today()
    d = ref
    while d.weekday() != 6:
        d += timedelta(days=1)
    return d


def charger_presences(id_membre=None, date_debut=None, date_fin=None):
    df = _lire(FEUILLE_PRESENCES, COLS_PRESENCES)
    if df.empty:
        return df
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
    df["ID_Membre"] = pd.to_numeric(df["ID_Membre"], errors="coerce").fillna(0).astype(int)
    df["Date_Dimanche"] = pd.to_datetime(df["Date_Dimanche"], errors="coerce", dayfirst=True)
    for c in ["Statut", "Signale_Le", "Valide_Le", "Valide_Par", "Notes"]:
        df[c] = df[c].map(_s)
    if id_membre is not None:
        df = df[df["ID_Membre"] == int(id_membre)]
    if date_debut is not None:
        df = df[df["Date_Dimanche"].notna() & (df["Date_Dimanche"].dt.date >= date_debut)]
    if date_fin is not None:
        df = df[df["Date_Dimanche"].notna() & (df["Date_Dimanche"].dt.date <= date_fin)]
    return df.sort_values(["Date_Dimanche", "ID_Membre"], ascending=[False, True]).reset_index(drop=True)


def signaler_presence(id_membre, date_dimanche, notes=""):
    if isinstance(date_dimanche, datetime):
        date_dimanche = date_dimanche.date()
    if date_dimanche.weekday() != 6:
        return False, "La date doit être un dimanche."
    df = charger_presences()
    if not df.empty:
        m = (df["ID_Membre"] == int(id_membre)) & (df["Date_Dimanche"].dt.date == date_dimanche)
        if m.any():
            return False, f"Déjà signalé (statut : {df.loc[m, 'Statut'].iloc[0]})."
    nid = 1 if df.empty else int(df["ID"].max()) + 1
    row = {
        "ID": nid,
        "ID_Membre": int(id_membre),
        "Date_Dimanche": pd.Timestamp(date_dimanche),
        "Statut": "EN_ATTENTE",
        "Signale_Le": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Valide_Le": "",
        "Valide_Par": "",
        "Notes": _s(notes),
    }
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    _ecrire(FEUILLE_PRESENCES, df[COLS_PRESENCES])
    return True, f"Présence signalée pour le {date_dimanche.strftime('%d/%m/%Y')} — en attente de validation."


def valider_presence(id_presence, admin_nom, accepter=True, notes=""):
    df = charger_presences()
    m = df["ID"] == int(id_presence)
    if not m.any():
        return False, "Présence introuvable."
    df.loc[m, "Statut"] = "VALIDE" if accepter else "REFUSE"
    df.loc[m, "Valide_Le"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    df.loc[m, "Valide_Par"] = _s(admin_nom)
    if notes:
        old = _s(df.loc[m, "Notes"].iloc[0])
        df.loc[m, "Notes"] = (old + " | " if old else "") + _s(notes)
    _ecrire(FEUILLE_PRESENCES, df[COLS_PRESENCES])
    return True, f"Présence {'validée' if accepter else 'refusée'}."


def presences_en_attente():
    df = charger_presences()
    if df.empty:
        return df
    return df[df["Statut"] == "EN_ATTENTE"].reset_index(drop=True)


def matrice_mois(annee, mois):
    dims = dimanches_du_mois(annee, mois)
    membres = charger_membres(actifs_only=True)
    fin = date(annee, mois + 1, 1) - timedelta(days=1) if mois < 12 else date(annee, 12, 31)
    pres = charger_presences(date_debut=date(annee, mois, 1), date_fin=fin)
    lignes = []
    for _, mb in membres.iterrows():
        row = {
            "ID": int(mb["ID"]),
            "Nom": _s(mb["Nom"]),
            "Prenom": _s(mb["Prenom"]),
            "Telephone": _s(mb["Telephone"]),
            "Profession": _s(mb.get("Profession")),
            "Est_Etudiant": _s(mb.get("Est_Etudiant")),
        }
        for d in dims:
            cell = ""
            if not pres.empty:
                m = (pres["ID_Membre"] == int(mb["ID"])) & (pres["Date_Dimanche"].dt.date == d)
                if m.any():
                    st = pres.loc[m, "Statut"].iloc[0]
                    cell = {"VALIDE": "V", "EN_ATTENTE": "A", "REFUSE": "R"}.get(st, "")
            row[d.isoformat()] = cell
        row["Total_Valide"] = sum(1 for d in dims if row.get(d.isoformat()) == "V")
        lignes.append(row)
    return dims, lignes


def to_excel_bytes_mois(annee, mois):
    dims, lignes = matrice_mois(annee, mois)
    cols = ["Nom", "Prenom", "Profession"] + [d.strftime("%d/%m/%Y") for d in dims] + ["Total_Valide"]
    rows = []
    for L in lignes:
        r = {"Nom": L["Nom"], "Prenom": L["Prenom"], "Profession": L.get("Profession", "")}
        for d in dims:
            r[d.strftime("%d/%m/%Y")] = {"V": "✓", "A": "…", "R": "✗"}.get(L.get(d.isoformat(), ""), "")
        r["Total_Valide"] = L["Total_Valide"]
        rows.append(r)
    df = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=f"{mois:02d}-{annee}")
    buf.seek(0)
    return buf.getvalue()


def appliquer_appel(date_dimanche, ids_presents, ids_affiches, admin_nom):
    """Appel fait par l'admin (sans internet côté membres).
    - coché   -> présence VALIDÉE immédiatement (créée ou mise à jour)
    - décoché -> une présence validée est retirée (les autres statuts sont laissés tels quels)
    Seuls les membres de `ids_affiches` sont concernés.
    """
    if isinstance(date_dimanche, datetime):
        date_dimanche = date_dimanche.date()
    if date_dimanche.weekday() != 6:
        return False, "La date doit être un dimanche."

    presents = {int(i) for i in ids_presents}
    affiches = {int(i) for i in ids_affiches}
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    df = charger_presences()
    if df.empty:
        df = pd.DataFrame(columns=COLS_PRESENCES)

    existant = {}  # id_membre -> (index ligne, statut)
    if not df.empty:
        jour = df[df["Date_Dimanche"].dt.date == date_dimanche]
        for idx, r in jour.iterrows():
            existant[int(r["ID_Membre"])] = (idx, r["Statut"])

    nid = 1 if df.empty else int(df["ID"].max()) + 1
    nouvelles, a_retirer = [], []
    n_ajout = n_valide = n_retire = 0

    for mid in affiches:
        if mid in presents:
            if mid in existant:
                idx, statut = existant[mid]
                if statut != "VALIDE":
                    df.loc[idx, "Statut"] = "VALIDE"
                    df.loc[idx, "Valide_Le"] = now
                    df.loc[idx, "Valide_Par"] = _s(admin_nom)
                    old = _s(df.loc[idx, "Notes"])
                    df.loc[idx, "Notes"] = (old + " | " if old else "") + "Appel admin"
                    n_valide += 1
            else:
                nouvelles.append({
                    "ID": nid,
                    "ID_Membre": mid,
                    "Date_Dimanche": pd.Timestamp(date_dimanche),
                    "Statut": "VALIDE",
                    "Signale_Le": now,
                    "Valide_Le": now,
                    "Valide_Par": _s(admin_nom),
                    "Notes": "Appel admin",
                })
                nid += 1
                n_ajout += 1
        else:
            if mid in existant and existant[mid][1] == "VALIDE":
                a_retirer.append(int(df.loc[existant[mid][0], "ID"]))
                n_retire += 1

    if a_retirer:
        df = df[~df["ID"].isin(a_retirer)]
    if nouvelles:
        df = pd.concat([df, pd.DataFrame(nouvelles)], ignore_index=True)
    df["Date_Dimanche"] = pd.to_datetime(df["Date_Dimanche"], errors="coerce")

    _ecrire(FEUILLE_PRESENCES, df[COLS_PRESENCES])
    return True, (
        f"Appel du {date_dimanche.strftime('%d/%m/%Y')} enregistré — "
        f"{n_ajout + n_valide} présence(s) validée(s), {n_retire} retirée(s)."
    )
