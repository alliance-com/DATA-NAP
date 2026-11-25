# -*- coding: utf-8 -*-
"""
export_data.py — Base export maritime NAP SARL
Fichier : export_data.xlsx
"""
from __future__ import annotations

import functools
import gc
import os
import time
from datetime import datetime, date, timedelta
from io import BytesIO

import pandas as pd

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FICHIER = os.path.join(_BASE_DIR, "export_data.xlsx")
VERSION_EXPORT = 1

# Feuilles
REF_BUYERS = "REF_Buyers"
REF_FWD = "REF_Forwarders"
REF_LINERS = "REF_Liners"
REF_COMM = "REF_Commodities"
REF_PORTS = "REF_Ports"
REF_BANKS = "REF_Banks"
REF_TANKS = "REF_Tanks"
REF_DOCS = "REF_DocTypes"
REF_STATUS = "REF_Status"

T_CONTRACTS = "T_Contracts"
T_BOOKINGS = "T_Bookings"
T_CONTAINERS = "T_Containers"
T_DOCUMENTS = "T_Documents"
T_INVOICES = "T_Invoices"
T_PAYMENTS = "T_Payments_Buyer"
T_BFU = "T_BFU"
T_CUSTOMS = "T_Customs"
T_FREIGHT = "T_Freight"
T_AGENTS = "T_Agent_Invoices"

STUFFING_STATUSES = ["PENDING", "IN PROGRESS", "COMPLETED"]
SHIPPING_STATUSES = ["PENDING", "BOOKED ON VESSEL", "SHIPPED"]
BFU_STATUSES = ["NO INITIATED", "PENDING", "COMPLETED"]
PAYMENT_STATUSES = ["PENDING", "PARTIAL", "RECEIVED"]
DOC_STATUSES = ["PENDING", "DONE"]

JOURS_ALERTE_ETD = 7
JOURS_ALERTE_STUFFING = 10


def _s(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, float) and v.is_integer():
        return str(int(v))  # n° de contrat / booking lus en float par pandas (ex. 4600063618.0)
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none", "nat") else t


def _n(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(str(v).replace(" ", "").replace(",", "."))
    except Exception:
        return 0.0


def _i(v):
    try:
        return int(float(_n(v)))
    except Exception:
        return 0


def _d(v):
    if v is None or v == "":
        return pd.NaT
    if isinstance(v, date) and not isinstance(v, datetime):
        return pd.Timestamp(v)
    ts = pd.to_datetime(v, errors="coerce", dayfirst=True)
    return ts


def ensure_file():
    if os.path.exists(FICHIER):
        return
    raise FileNotFoundError(
        f"{FICHIER} introuvable. Placez export_data.xlsx à la racine du projet."
    )


def lire_feuille(nom, cols=None):
    ensure_file()
    try:
        df = pd.read_excel(FICHIER, sheet_name=nom)
    except ValueError:
        return pd.DataFrame(columns=cols or [])
    if cols:
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[[c for c in cols if c in df.columns]]
    return df


def _remplacer_avec_retry(tmp, cible, essais=8, pause=0.6):
    """os.replace avec nouvelles tentatives : OneDrive / antivirus / Excel verrouillent
    parfois le fichier un court instant sous Windows."""
    gc.collect()  # libère d'éventuels handles pandas/openpyxl encore ouverts
    for i in range(essais):
        try:
            os.replace(tmp, cible)
            return
        except PermissionError:
            if i == essais - 1:
                try:
                    os.remove(tmp)
                except OSError:
                    pass
                raise PermissionError(
                    f"Impossible d'enregistrer « {os.path.basename(cible)} » : le fichier est verrouillé. "
                    "Ferme-le dans Excel (vérifie aussi le Gestionnaire des tâches) ou attends la fin de la "
                    "synchronisation OneDrive, puis réessaie."
                )
            time.sleep(pause)


def _securise(fn):
    """Transforme une PermissionError en (False, message) : la page affiche un message clair."""
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except PermissionError as e:
            return False, str(e)
    return wrapper


def ecrire_feuille(nom, df):
    ensure_file()
    # recharge tout le classeur (fichier refermé AVANT le remplacement) et réécrit la feuille
    with pd.ExcelFile(FICHIER) as xls:
        sheets = {s: xls.parse(s) for s in xls.sheet_names}
    sheets[nom] = df
    tmp = FICHIER.replace(".xlsx", ".tmp.xlsx")
    try:
        with pd.ExcelWriter(tmp, engine="openpyxl") as w:
            for s, d in sheets.items():
                d.to_excel(w, sheet_name=s, index=False)
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise
    _remplacer_avec_retry(tmp, FICHIER)


# ---------- Référentiels ----------
def liste_buyers(actifs_only=True):
    df = lire_feuille(REF_BUYERS)
    if df.empty:
        return []
    if actifs_only and "Actif" in df.columns:
        df = df[df["Actif"].astype(str).str.upper() == "OUI"]
    return sorted({_s(x) for x in df.get("Nom", []) if _s(x)})


def liste_forwarders(actifs_only=True):
    df = lire_feuille(REF_FWD)
    if df.empty:
        return []
    if actifs_only and "Actif" in df.columns:
        df = df[df["Actif"].astype(str).str.upper() == "OUI"]
    return sorted({_s(x) for x in df.get("Nom", []) if _s(x)})


def liste_liners(actifs_only=True):
    df = lire_feuille(REF_LINERS)
    if df.empty:
        return []
    if actifs_only and "Actif" in df.columns:
        df = df[df["Actif"].astype(str).str.upper() == "OUI"]
    return sorted({_s(x) for x in df.get("Nom", []) if _s(x)})


def liste_commodities(actifs_only=True):
    df = lire_feuille(REF_COMM)
    if df.empty:
        return []
    if actifs_only and "Actif" in df.columns:
        df = df[df["Actif"].astype(str).str.upper() == "OUI"]
    return sorted({_s(x) for x in df.get("Libelle", []) if _s(x)})


def liste_ports():
    df = lire_feuille(REF_PORTS)
    if df.empty:
        return []
    return sorted({_s(x) for x in df.get("POD", []) if _s(x)})


def liste_countries():
    df = lire_feuille(REF_PORTS)
    if df.empty:
        return []
    return sorted({_s(x) for x in df.get("Country", []) if _s(x)})


def liste_banks():
    df = lire_feuille(REF_BANKS)
    if df.empty:
        return []
    return sorted({_s(x) for x in df.get("Banque", []) if _s(x)})


def liste_tanks():
    df = lire_feuille(REF_TANKS)
    if df.empty:
        return []
    if "Actif" in df.columns:
        df = df[df["Actif"].astype(str).str.upper() == "OUI"]
    return [_s(x) for x in df.get("Tank", []) if _s(x)]


def liste_doc_types():
    df = lire_feuille(REF_DOCS)
    if df.empty:
        return []
    if "Ordre" in df.columns:
        df = df.sort_values("Ordre")
    return [_s(x) for x in df.get("Code", []) if _s(x)]


# ---------- Contracts ----------
COLS_CONTRACTS = [
    "Contract_No", "Date_Contract", "Code_Buyer", "Buyer", "Commodity",
    "Qty_Contract_MT", "Price_Unit", "Currency", "POD", "Country", "Notes", "Actif",
]


def charger_contracts():
    df = lire_feuille(T_CONTRACTS, COLS_CONTRACTS)
    if df.empty:
        return df
    df["Date_Contract"] = df["Date_Contract"].map(_d)
    df["Qty_Contract_MT"] = df["Qty_Contract_MT"].map(_n)
    df["Price_Unit"] = df["Price_Unit"].map(_n)
    for c in ["Contract_No", "Code_Buyer", "Buyer", "Commodity", "Currency", "POD", "Country", "Notes", "Actif"]:
        df[c] = df[c].map(_s)
    return df


def sauver_contracts(df):
    ecrire_feuille(T_CONTRACTS, df[COLS_CONTRACTS] if not df.empty else pd.DataFrame(columns=COLS_CONTRACTS))


@_securise
def ajouter_contract(data):
    df = charger_contracts()
    no = _s(data.get("Contract_No"))
    if not no:
        return False, "N° contrat obligatoire."
    if not df.empty and no in df["Contract_No"].values:
        return False, f"Contrat {no} existe déjà."
    row = {c: data.get(c, "") for c in COLS_CONTRACTS}
    row["Contract_No"] = no
    row["Date_Contract"] = _d(data.get("Date_Contract"))
    row["Qty_Contract_MT"] = _n(data.get("Qty_Contract_MT"))
    row["Price_Unit"] = _n(data.get("Price_Unit"))
    row["Actif"] = _s(data.get("Actif") or "OUI")
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    sauver_contracts(df)
    return True, f"Contrat {no} enregistré."


# ---------- Bookings ----------
COLS_BOOKINGS = [
    "Booking_No", "Contract_No", "BL_No", "Commodity", "Buyer", "POD", "Country",
    "Liner", "Forwarder", "Ctr_Booked", "Ctr_Loaded", "Balance", "Not_Allocated",
    "Net_Qty_KG", "Gross_Qty_KG", "Nb_Bags", "Storage_Tank", "Rotation",
    "Stuffing_Status", "Vessel", "VOY", "ETD", "ETA", "Shipping_Status",
    "Date_Creation", "User", "Notes",
]


def charger_bookings():
    df = lire_feuille(T_BOOKINGS, COLS_BOOKINGS)
    if df.empty:
        return df
    for c in ["ETD", "ETA", "Date_Creation"]:
        df[c] = df[c].map(_d)
    for c in ["Ctr_Booked", "Ctr_Loaded", "Balance", "Not_Allocated", "Nb_Bags"]:
        df[c] = df[c].map(_i)
    for c in ["Net_Qty_KG", "Gross_Qty_KG"]:
        df[c] = df[c].map(_n)
    for c in COLS_BOOKINGS:
        if c not in ("ETD", "ETA", "Date_Creation", "Ctr_Booked", "Ctr_Loaded",
                     "Balance", "Not_Allocated", "Nb_Bags", "Net_Qty_KG", "Gross_Qty_KG"):
            df[c] = df[c].map(_s)
    # recalcul balance
    df["Balance"] = (df["Ctr_Booked"] - df["Ctr_Loaded"]).clip(lower=0)
    return df


def sauver_bookings(df):
    out = df.copy()
    if not out.empty:
        out["Balance"] = (out["Ctr_Booked"].map(_i) - out["Ctr_Loaded"].map(_i)).clip(lower=0)
    ecrire_feuille(T_BOOKINGS, out[COLS_BOOKINGS] if not out.empty else pd.DataFrame(columns=COLS_BOOKINGS))


@_securise
def ajouter_booking(data, user=""):
    df = charger_bookings()
    bno = _s(data.get("Booking_No"))
    if not bno:
        return False, "N° Booking obligatoire."
    if not df.empty and bno in df["Booking_No"].astype(str).values:
        return False, f"Booking {bno} existe déjà."
    bl = _s(data.get("BL_No") or bno)
    booked = _i(data.get("Ctr_Booked"))
    loaded = _i(data.get("Ctr_Loaded"))
    row = {c: data.get(c, "") for c in COLS_BOOKINGS}
    row.update({
        "Booking_No": bno,
        "BL_No": bl,
        "Ctr_Booked": booked,
        "Ctr_Loaded": loaded,
        "Balance": max(booked - loaded, 0),
        "Not_Allocated": _i(data.get("Not_Allocated")),
        "Net_Qty_KG": _n(data.get("Net_Qty_KG")),
        "Gross_Qty_KG": _n(data.get("Gross_Qty_KG")),
        "Nb_Bags": _i(data.get("Nb_Bags")),
        "Stuffing_Status": _s(data.get("Stuffing_Status") or "PENDING"),
        "Shipping_Status": _s(data.get("Shipping_Status") or "PENDING"),
        "ETD": _d(data.get("ETD")),
        "ETA": _d(data.get("ETA")),
        "Date_Creation": pd.Timestamp(datetime.now()),
        "User": _s(user),
    })
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    sauver_bookings(df)
    return True, f"Booking {bno} enregistré."


@_securise
def maj_booking(booking_no, updates: dict):
    df = charger_bookings()
    if df.empty or booking_no not in df["Booking_No"].astype(str).values:
        return False, "Booking introuvable."
    m = df["Booking_No"].astype(str) == str(booking_no)
    for k, v in updates.items():
        if k in df.columns:
            if k in ("ETD", "ETA", "Date_Creation"):
                df.loc[m, k] = _d(v)
            elif k in ("Ctr_Booked", "Ctr_Loaded", "Not_Allocated", "Nb_Bags"):
                df.loc[m, k] = _i(v)
            elif k in ("Net_Qty_KG", "Gross_Qty_KG"):
                df.loc[m, k] = _n(v)
            else:
                df.loc[m, k] = _s(v)
    sauver_bookings(df)
    return True, f"Booking {booking_no} mis à jour."


# ---------- Containers ----------
COLS_CTR = [
    "Container_No", "Booking_No", "BL_No", "Type_Ctr", "Seal", "Weight_KG",
    "Status_Load", "Tank", "Date_Stuffing", "Notes",
]


def charger_containers():
    df = lire_feuille(T_CONTAINERS, COLS_CTR)
    if df.empty:
        return df
    df["Date_Stuffing"] = df["Date_Stuffing"].map(_d)
    df["Weight_KG"] = df["Weight_KG"].map(_n)
    for c in ["Container_No", "Booking_No", "BL_No", "Type_Ctr", "Seal", "Status_Load", "Tank", "Notes"]:
        df[c] = df[c].map(_s)
    return df


def sauver_containers(df):
    ecrire_feuille(T_CONTAINERS, df[COLS_CTR] if not df.empty else pd.DataFrame(columns=COLS_CTR))


@_securise
def ajouter_container(data):
    df = charger_containers()
    cno = _s(data.get("Container_No"))
    if not cno:
        return False, "N° conteneur obligatoire."
    if not df.empty and cno in df["Container_No"].values:
        return False, f"Conteneur {cno} existe déjà."
    row = {c: data.get(c, "") for c in COLS_CTR}
    row["Container_No"] = cno
    row["Weight_KG"] = _n(data.get("Weight_KG"))
    row["Date_Stuffing"] = _d(data.get("Date_Stuffing"))
    row["Status_Load"] = _s(data.get("Status_Load") or "LOADED")
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    sauver_containers(df)
    # maj ctr loaded on booking
    bno = _s(data.get("Booking_No"))
    if bno:
        n = len(df[df["Booking_No"] == bno])
        maj_booking(bno, {"Ctr_Loaded": n, "Stuffing_Status": "COMPLETED" if n else "IN PROGRESS"})
    return True, f"Conteneur {cno} enregistré."


# ---------- KPI & alertes ----------
def indicateurs(df_b=None):
    if df_b is None:
        df_b = charger_bookings()
    if df_b.empty:
        return {
            "bookings": 0, "stuffed": 0, "pending_stuff": 0,
            "ctr_booked": 0, "ctr_loaded": 0, "qty_mt": 0.0,
            "pct_load": 0.0, "shipped": 0, "buyers": 0, "forwarders": 0,
        }
    ctr_b = int(df_b["Ctr_Booked"].sum())
    ctr_l = int(df_b["Ctr_Loaded"].sum())
    return {
        "bookings": len(df_b),
        "stuffed": int((df_b["Stuffing_Status"] == "COMPLETED").sum()),
        "pending_stuff": int((df_b["Stuffing_Status"] != "COMPLETED").sum()),
        "ctr_booked": ctr_b,
        "ctr_loaded": ctr_l,
        "qty_mt": float(df_b["Net_Qty_KG"].sum()) / 1000.0,
        "pct_load": (ctr_l / ctr_b) if ctr_b else 0.0,
        "shipped": int((df_b["Shipping_Status"] == "SHIPPED").sum()),
        "buyers": df_b["Buyer"].nunique(),
        "forwarders": df_b["Forwarder"].nunique(),
    }


def alertes(df_b=None):
    if df_b is None:
        df_b = charger_bookings()
    today = date.today()
    items = []
    if df_b.empty:
        return items
    for _, r in df_b.iterrows():
        bno = r["Booking_No"]
        st = r["Stuffing_Status"]
        sh = r["Shipping_Status"]
        etd = r["ETD"]
        bal = _i(r["Balance"])
        # conteneurs manquants
        if bal > 0 and st != "COMPLETED":
            items.append({
                "Niveau": "WARNING",
                "Type": "STUFFING",
                "Booking": bno,
                "Message": f"{bal} conteneur(s) restant(s) à empoter — {r['Forwarder']}",
            })
        # ETD proche
        if pd.notna(etd):
            d_etd = pd.Timestamp(etd).date()
            delta = (d_etd - today).days
            if delta <= JOURS_ALERTE_ETD and st != "COMPLETED":
                items.append({
                    "Niveau": "CRITIQUE" if delta < 0 else "CRITIQUE",
                    "Type": "ETD",
                    "Booking": bno,
                    "Message": f"ETD {d_etd.strftime('%d/%m/%Y')} ({delta} j) — stuffing {st}",
                })
        # shipped but not stuffed
        if sh == "SHIPPED" and st != "COMPLETED":
            items.append({
                "Niveau": "CRITIQUE",
                "Type": "COHERENCE",
                "Booking": bno,
                "Message": "Statut SHIPPED alors que stuffing non COMPLETED",
            })
    return items


def bookings_pending_stuffing(df_b=None):
    if df_b is None:
        df_b = charger_bookings()
    if df_b.empty:
        return df_b
    return df_b[df_b["Stuffing_Status"] != "COMPLETED"].copy()


def to_excel_bytes(df):
    buf = BytesIO()
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[c]):
            out[c] = out[c].dt.strftime("%d/%m/%Y")
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="EXPORT")
    buf.seek(0)
    return buf.getvalue()
