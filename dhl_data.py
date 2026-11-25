# -*- coding: utf-8 -*-
"""Données envois DHL — stockage Excel local."""
import os
import re
import shutil
from datetime import datetime, date
from io import BytesIO

import pandas as pd

FICHIER_DHL = "envois_dhl.xlsx"
FEUILLE = "ENVOIS"
DOSSIER_DOCS = "documents_dhl"

COLS = [
    "ID",
    "Date_Envoi",
    "Waybill",
    "Type_Document",
    "Nb_Documents",
    "Poids_KG",
    "Montant",
    "Mode_Paiement",
    "Destinataire",
    "Pays_Destination",
    "Description",
    "Fichiers",
    "Cree_Par",
    "Heure_Creation",
    "Remarque",
]

TYPES_DOC = [
    "BL / Connaissement",
    "Certificat phytosanitaire",
    "Certificat d'origine",
    "Facture commerciale",
    "Packing list",
    "Documents douaniers",
    "Courrier divers",
    "Sets Documents",
    "Echantillon du Produit",
    "Autre",
]

MODES_PAIEMENT = ["Espèce", "Mobile Money", "Virement"]


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


def _n(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(str(v).replace(" ", "").replace(",", "."))
    except Exception:
        return 0.0


def _ensure_dirs():
    os.makedirs(DOSSIER_DOCS, exist_ok=True)


def charger():
    _ensure_dirs()
    if not os.path.exists(FICHIER_DHL):
        return pd.DataFrame(columns=COLS)
    try:
        df = pd.read_excel(FICHIER_DHL, sheet_name=FEUILLE)
    except ValueError:
        try:
            df = pd.read_excel(FICHIER_DHL)
        except Exception:
            return pd.DataFrame(columns=COLS)
    except Exception:
        return pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in df.columns:
            df[c] = "" if c not in ("ID", "Nb_Documents", "Poids_KG", "Montant") else 0
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
    df["Nb_Documents"] = pd.to_numeric(df["Nb_Documents"], errors="coerce").fillna(0).astype(int)
    df["Poids_KG"] = df["Poids_KG"].map(_n)
    df["Montant"] = df["Montant"].map(_n)
    for c in ["Waybill", "Type_Document", "Mode_Paiement", "Destinataire", "Pays_Destination",
              "Description", "Fichiers", "Cree_Par", "Heure_Creation", "Remarque"]:
        df[c] = df[c].map(_s)
    df["Date_Envoi"] = pd.to_datetime(df["Date_Envoi"], errors="coerce", dayfirst=True)
    return df.sort_values(["Date_Envoi", "ID"], ascending=[False, False]).reset_index(drop=True)


def sauver(df):
    _ensure_dirs()
    out = df.copy()
    for c in COLS:
        if c not in out.columns:
            out[c] = ""
    out = out[COLS]
    tmp = FICHIER_DHL.replace(".xlsx", ".tmp.xlsx")
    out.to_excel(tmp, sheet_name=FEUILLE, index=False)
    os.replace(tmp, FICHIER_DHL)


def prochain_id(df=None):
    if df is None:
        df = charger()
    if df.empty:
        return 1
    return int(df["ID"].max()) + 1


def dossier_envoi(envoi_id, waybill=""):
    safe = re.sub(r'[\\/*?:"<>|]', "", _s(waybill)) or f"ID{envoi_id}"
    path = os.path.join(DOSSIER_DOCS, f"{envoi_id:04d}_{safe}")
    os.makedirs(path, exist_ok=True)
    return path


def enregistrer_fichiers(envoi_id, waybill, uploaded_files):
    """Sauvegarde les fichiers uploadés Streamlit. Retourne liste de chemins relatifs."""
    if not uploaded_files:
        return []
    dossier = dossier_envoi(envoi_id, waybill)
    sauves = []
    for uf in uploaded_files:
        nom = os.path.basename(uf.name)
        nom = re.sub(r'[\\/*?:"<>|]', "_", nom)
        dest = os.path.join(dossier, nom)
        # éviter écrasement
        base, ext = os.path.splitext(nom)
        n = 1
        while os.path.exists(dest):
            dest = os.path.join(dossier, f"{base}_{n}{ext}")
            n += 1
        with open(dest, "wb") as f:
            f.write(uf.getbuffer())
        sauves.append(os.path.relpath(dest).replace("\\", "/"))
    return sauves


def ajouter(data, uploaded_files=None, utilisateur=""):
    df = charger()
    waybill = _s(data.get("Waybill"))
    if waybill and waybill in set(df["Waybill"]):
        return False, f"Le waybill {waybill} existe déjà.", None

    eid = prochain_id(df)
    fichiers = enregistrer_fichiers(eid, waybill, uploaded_files or [])

    ligne = {
        "ID": eid,
        "Date_Envoi": pd.Timestamp(data.get("Date_Envoi") or date.today()),
        "Waybill": waybill,
        "Type_Document": _s(data.get("Type_Document")),
        "Nb_Documents": int(_n(data.get("Nb_Documents"))),
        "Poids_KG": _n(data.get("Poids_KG")),
        "Montant": _n(data.get("Montant")),
        "Mode_Paiement": _s(data.get("Mode_Paiement")),
        "Destinataire": _s(data.get("Destinataire")),
        "Pays_Destination": _s(data.get("Pays_Destination")),
        "Description": _s(data.get("Description")),
        "Fichiers": " | ".join(fichiers),
        "Cree_Par": _s(utilisateur),
        "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Remarque": _s(data.get("Remarque")),
    }
    df = pd.concat([df, pd.DataFrame([ligne])], ignore_index=True)
    sauver(df)
    return True, f"Envoi DHL {waybill or eid} enregistré ({len(fichiers)} fichier(s)).", eid


def modifier(envoi_id, data, uploaded_files=None, utilisateur=""):
    df = charger()
    envoi_id = int(envoi_id)
    mask = df["ID"] == envoi_id
    if not mask.any():
        return False, "Envoi introuvable."

    waybill = _s(data.get("Waybill"))

    # doublon waybill autre ligne
    autres = df.loc[~mask, "Waybill"].map(_s)
    if waybill and waybill in set(autres):
        return False, f"Le waybill {waybill} est déjà utilisé par un autre envoi."

    nouveaux = enregistrer_fichiers(envoi_id, waybill, uploaded_files or [])
    anciens = _s(df.loc[mask, "Fichiers"].iloc[0])
    if nouveaux:
        fichiers = (anciens + " | " if anciens else "") + " | ".join(nouveaux)
    else:
        fichiers = anciens

    df.loc[mask, "Date_Envoi"] = pd.Timestamp(data.get("Date_Envoi") or date.today())
    df.loc[mask, "Waybill"] = waybill
    df.loc[mask, "Type_Document"] = _s(data.get("Type_Document"))
    df.loc[mask, "Nb_Documents"] = int(_n(data.get("Nb_Documents")))
    df.loc[mask, "Poids_KG"] = _n(data.get("Poids_KG"))
    df.loc[mask, "Montant"] = _n(data.get("Montant"))
    df.loc[mask, "Mode_Paiement"] = _s(data.get("Mode_Paiement"))
    df.loc[mask, "Destinataire"] = _s(data.get("Destinataire"))
    df.loc[mask, "Pays_Destination"] = _s(data.get("Pays_Destination"))
    df.loc[mask, "Description"] = _s(data.get("Description"))
    df.loc[mask, "Fichiers"] = fichiers
    df.loc[mask, "Remarque"] = _s(data.get("Remarque"))
    sauver(df)
    return True, f"Envoi {waybill or envoi_id} mis à jour."


def supprimer(envoi_id):
    df = charger()
    envoi_id = int(envoi_id)
    mask = df["ID"] == envoi_id
    if not mask.any():
        return False, "Envoi introuvable."
    waybill = _s(df.loc[mask, "Waybill"].iloc[0])
    df = df.loc[~mask].reset_index(drop=True)
    sauver(df)
    # dossier docs conservé volontairement (historique)
    return True, f"Envoi {waybill or envoi_id} supprimé du registre."


def liste_fichiers(fichiers_str):
    parts = [p.strip() for p in _s(fichiers_str).split("|") if p.strip()]
    return [p for p in parts if os.path.exists(p)]


def to_excel_bytes(df):
    buf = BytesIO()
    out = df.copy()
    if "Date_Envoi" in out.columns:
        out["Date_Envoi"] = pd.to_datetime(out["Date_Envoi"], errors="coerce").dt.strftime("%d/%m/%Y")
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="ENVOIS_DHL")
    buf.seek(0)
    return buf.getvalue()
