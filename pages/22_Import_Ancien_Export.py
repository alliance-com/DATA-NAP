# -*- coding: utf-8 -*-
"""
Page Streamlit — Import de l'ancien fichier Export
(« Export General Report 2026-2027 (UPDATE).xlsx » -> export_data.xlsx)

Charge automatiquement :
  - T_Contracts  : un contrat par n° de contrat trouvé dans le master tracker
  - T_Bookings   : une ligne par booking du master tracker (feuille « Export master tracker_ »)
  - T_Containers : un conteneur par ligne des rapports de chargement (« SHEA BUTTER », « SHEA DOC »…)

Sécurités : copie de sauvegarde avant import, on n'ajoute QUE ce qui n'existe pas déjà,
aucune ligne existante n'est modifiée ni supprimée.

À placer dans DATA/pages/   (le package DATA/export/ doit exister).
"""
import io
import os
import re
import shutil
import sys
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd

try:
    from export import export_data as ex
except ImportError:                                   # lancement depuis un autre dossier
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from export import export_data as ex


# ====================== OUTILS DE LECTURE ======================
def _norm(s) -> str:
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", s)


def _txt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        if pd.isna(v):
            return ""
        if v.is_integer():
            return str(int(v))
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none", "nat", "-") or t.startswith("#") else t


def _num(v) -> float:
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(str(v).replace(" ", "").replace(",", "."))
    except Exception:
        return 0.0


def _date(v):
    if v is None or v == "":
        return pd.NaT
    return pd.to_datetime(v, errors="coerce", dayfirst=True)


def _entetes(ws, requis, max_lignes=15):
    """Cherche la ligne d'en-têtes contenant toutes les colonnes `requis` (noms normalisés)."""
    for r in range(1, min(ws.max_row, max_lignes) + 1):
        cols = {}
        for c in ws[r]:
            k = _norm(c.value)
            if k and k not in cols:
                cols[k] = c.column
        if all(k in cols for k in requis):
            return r, cols
    return None, {}


def _val(ws, r, cols, *noms):
    for n in noms:
        if n in cols:
            return ws.cell(r, cols[n]).value
    return None


# ====================== DÉTECTION DES FEUILLES ======================
REQ_TRACKER = ["contractno", "bookings", "containerbooked", "containerloaded"]
REQ_CONTENEURS = ["containernumber", "bkgno"]


def charger_classeur(data: bytes):
    from openpyxl import load_workbook
    return load_workbook(io.BytesIO(data), data_only=True)


def detecter_feuilles(wb):
    trackers, conteneurs = [], []
    for ws in wb.worksheets:
        r, cols = _entetes(ws, REQ_TRACKER)
        if r:
            score = sum(k in cols for k in ("netqtykg", "stuffingstatus", "etd", "eta", "vessel",
                                            "shippedstatus", "priceunit", "currency", "storagelocation", "rotation"))
            n = sum(1 for i in range(r + 1, ws.max_row + 1) if _txt(_val(ws, i, cols, "bookings")))
            trackers.append({"feuille": ws.title, "visible": ws.sheet_state == "visible", "score": score, "lignes": n})
        r, cols = _entetes(ws, REQ_CONTENEURS)
        if r:
            n = sum(1 for i in range(r + 1, ws.max_row + 1) if _txt(_val(ws, i, cols, "containernumber")))
            conteneurs.append({"feuille": ws.title, "visible": ws.sheet_state == "visible", "lignes": n})
    trackers.sort(key=lambda t: (not t["visible"], -t["score"], -t["lignes"]))
    conteneurs.sort(key=lambda t: (not t["visible"], -t["lignes"]))
    return trackers, conteneurs


# ====================== EXTRACTION ======================
def _statut_stuffing(v) -> str:
    s = _txt(v).upper()
    if s in ("COMPLETED", "COMPLETE", "DONE"):
        return "COMPLETED"
    if "PROGRESS" in s:
        return "IN PROGRESS"
    return "PENDING"


def _statut_expedition(v) -> str:
    s = _txt(v).upper()
    if s in ("DONE", "SHIPPED", "COMPLETED", "OK"):
        return "SHIPPED"
    if "VESSEL" in s or "BOOKED" in s:
        return "BOOKED ON VESSEL"
    return "PENDING"


def lire_tracker(ws) -> pd.DataFrame:
    r0, cols = _entetes(ws, REQ_TRACKER)
    lignes = []
    for r in range(r0 + 1, ws.max_row + 1):
        bk = _txt(_val(ws, r, cols, "bookings"))
        if not bk:
            continue
        lignes.append({
            "Booking_No": bk,
            "Contract_No": _txt(_val(ws, r, cols, "contractno")),
            "BL_No": _txt(_val(ws, r, cols, "bls")) or bk,
            "Commodity": _txt(_val(ws, r, cols, "commodity")),
            "Buyer": _txt(_val(ws, r, cols, "buyer")),
            "POD": _txt(_val(ws, r, cols, "pod")),
            "Country": _txt(_val(ws, r, cols, "country")),
            "Liner": _txt(_val(ws, r, cols, "liner")),
            "Forwarder": _txt(_val(ws, r, cols, "forwarder")),
            "Ctr_Booked": int(_num(_val(ws, r, cols, "containerbooked"))),
            "Ctr_Loaded": int(_num(_val(ws, r, cols, "containerloaded"))),
            "Not_Allocated": int(_num(_val(ws, r, cols, "notyetallocated"))),
            "Net_Qty_KG": _num(_val(ws, r, cols, "netqtykg")),
            "Storage_Tank": _txt(_val(ws, r, cols, "storagelocation")),
            "Rotation": _txt(_val(ws, r, cols, "rotation")),
            "Stuffing_Status": _statut_stuffing(_val(ws, r, cols, "stuffingstatus")),
            "Vessel": _txt(_val(ws, r, cols, "vessel")),
            "VOY": _txt(_val(ws, r, cols, "voy")),
            "ETD": _date(_val(ws, r, cols, "etd")),
            "ETA": _date(_val(ws, r, cols, "eta")),
            "Shipping_Status": _statut_expedition(_val(ws, r, cols, "shippedstatus")),
            "_price": _num(_val(ws, r, cols, "priceunit")),
            "_currency": _txt(_val(ws, r, cols, "currency")).upper(),
        })
    return pd.DataFrame(lignes)


def lire_conteneurs(ws) -> pd.DataFrame:
    r0, cols = _entetes(ws, REQ_CONTENEURS)
    lignes = []
    for r in range(r0 + 1, ws.max_row + 1):
        cno = _txt(_val(ws, r, cols, "containernumber"))
        if not cno:
            continue
        poids = _num(_val(ws, r, cols, "productweight")) or _num(_val(ws, r, cols, "netweight"))
        lignes.append({
            "Container_No": cno.upper(),
            "bkg": _txt(_val(ws, r, cols, "bkgno")),
            "bl": _txt(_val(ws, r, cols, "blno")),
            "buyer": _txt(_val(ws, r, cols, "buyername")),
            "Seal": _txt(_val(ws, r, cols, "plombno")),
            "Weight_KG": poids,
            "sacs": int(_num(_val(ws, r, cols, "noofsacs"))),
            "Tank": _txt(_val(ws, r, cols, "loadingwh")),
            "Date_Stuffing": _date(_val(ws, r, cols, "date")),
            "camion": _txt(_val(ws, r, cols, "camionnumber")),
            "Feuille": ws.title,
        })
    return pd.DataFrame(lignes)


def affecter_bookings(df_ctr: pd.DataFrame, df_trk: pd.DataFrame):
    """Rattache chaque conteneur à un booking du tracker (au besoin via le nom de l'acheteur)."""
    connus = set(df_trk["Booking_No"])
    par_acheteur = {}
    for _, r in df_trk.iterrows():
        par_acheteur.setdefault(_norm(r["Buyer"]), set()).add(r["Booking_No"])
    finals, notes, infos = [], [], []
    for _, r in df_ctr.iterrows():
        bk, note = r["bkg"], ""
        if bk not in connus:
            b = _norm(r["buyer"])
            cand = {x for nb, bks in par_acheteur.items() if nb and b and (nb.startswith(b) or b.startswith(nb))
                    for x in bks}
            if len(cand) == 1:
                nouveau = next(iter(cand))
                note = f"BKG d'origine {bk}"
                infos.append((r["Container_No"], bk, nouveau))
                bk = nouveau
        finals.append(bk)
        notes.append(note)
    out = df_ctr.copy()
    out["Booking_No"] = finals
    out["_note"] = notes
    return out, infos


def construire_tout(trk: pd.DataFrame, ctr: pd.DataFrame, aligner_loaded: bool = False):
    """Retourne (contrats, bookings, conteneurs, controles) au format des tables export_data."""
    controles = []
    ctr, remaps = affecter_bookings(ctr, trk) if not ctr.empty else (ctr, [])
    for cno, ancien, nouveau in remaps:
        controles.append({"Niveau": "INFO", "Sujet": f"Conteneur {cno}",
                          "Détail": f"BKG « {ancien} » absent du tracker → rattaché au booking {nouveau} (même acheteur)"})

    # --- conteneurs
    bl_par_bk = dict(zip(trk["Booking_No"], trk["BL_No"]))
    cont = pd.DataFrame({
        "Container_No": ctr.get("Container_No", pd.Series(dtype=str)),
        "Booking_No": ctr.get("Booking_No", pd.Series(dtype=str)),
        "BL_No": [(b or bl_par_bk.get(k, "") or k) for b, k in zip(ctr.get("bl", []), ctr.get("Booking_No", []))],
        "Type_Ctr": "",
        "Seal": ctr.get("Seal", ""),
        "Weight_KG": ctr.get("Weight_KG", 0.0),
        "Status_Load": "LOADED",
        "Tank": ctr.get("Tank", ""),
        "Date_Stuffing": ctr.get("Date_Stuffing", pd.NaT),
        "Notes": [" · ".join(x for x in (f"Camion {c}" if c else "", n) if x)
                  for c, n in zip(ctr.get("camion", []), ctr.get("_note", []))],
    })[ex.COLS_CTR] if not ctr.empty else pd.DataFrame(columns=ex.COLS_CTR)

    # --- bookings
    nb_ctr = Counter(cont["Booking_No"]) if not cont.empty else Counter()
    sacs = ctr.groupby("Booking_No")["sacs"].sum().to_dict() if not ctr.empty else {}
    book = trk.drop(columns=["_price", "_currency"]).copy()
    for bk_no, charges, n_found in ((r["Booking_No"], r["Ctr_Loaded"], nb_ctr.get(r["Booking_No"], 0)) for _, r in book.iterrows()):
        if n_found and n_found != charges:
            controles.append({"Niveau": "ATTENTION", "Sujet": f"Booking {bk_no}",
                              "Détail": f"{charges} conteneur(s) chargé(s) dans le tracker, {n_found} trouvé(s) dans les rapports de chargement"})
    if aligner_loaded:
        for i, r in book.iterrows():
            n_found = nb_ctr.get(r["Booking_No"], 0)
            if n_found and n_found != r["Ctr_Loaded"]:
                book.at[i, "Ctr_Loaded"] = n_found
    for bk_no in sorted(set(nb_ctr) - set(book["Booking_No"])):
        controles.append({"Niveau": "ATTENTION", "Sujet": f"Booking {bk_no}",
                          "Détail": f"{nb_ctr[bk_no]} conteneur(s) rattaché(s) à un booking absent du tracker (non importés avec un booking)"})
    book["Balance"] = (book["Ctr_Booked"] - book["Ctr_Loaded"]).clip(lower=0)
    book["Gross_Qty_KG"] = 0.0
    book["Nb_Bags"] = [int(sacs.get(b, 0)) for b in book["Booking_No"]]
    book["Date_Creation"] = pd.Timestamp(datetime.now())
    book["User"] = "IMPORT"
    book["Notes"] = "Importé de l'ancien fichier"
    book = book[ex.COLS_BOOKINGS]
    for _, r in book.iterrows():
        if not r["Contract_No"]:
            controles.append({"Niveau": "INFO", "Sujet": f"Booking {r['Booking_No']}",
                              "Détail": "Aucun n° de contrat dans l'ancien fichier (booking importé sans contrat)"})
        if r["Shipping_Status"] == "SHIPPED" and r["Stuffing_Status"] != "COMPLETED":
            controles.append({"Niveau": "ATTENTION", "Sujet": f"Booking {r['Booking_No']}",
                              "Détail": "SHIPPED alors que le stuffing n'est pas COMPLETED"})

    # --- contrats
    lignes = []
    for no, g in trk[trk["Contract_No"] != ""].groupby("Contract_No", sort=False):
        prix = [p for p in g["_price"] if p]
        dev = [c for c in g["_currency"] if c]
        premier = lambda col: next((x for x in g[col] if x), "")
        mt = float(book[book["Contract_No"] == no]["Net_Qty_KG"].sum()) / 1000.0
        lignes.append({
            "Contract_No": no, "Date_Contract": pd.NaT, "Code_Buyer": "",
            "Buyer": premier("Buyer"), "Commodity": premier("Commodity"),
            "Qty_Contract_MT": 0.0,
            "Price_Unit": Counter(prix).most_common(1)[0][0] if prix else 0.0,
            "Currency": Counter(dev).most_common(1)[0][0] if dev else "",
            "POD": premier("POD"), "Country": premier("Country"),
            "Notes": "Importé de l'ancien fichier — quantité et date du contrat à compléter",
            "Actif": "OUI", "_mt_charge": round(mt, 3),
        })
    contrats = pd.DataFrame(lignes, columns=ex.COLS_CONTRACTS + ["_mt_charge"])
    return contrats, book.reset_index(drop=True), cont.reset_index(drop=True), controles


# ====================== IMPORT DANS export_data.xlsx ======================
def _ajouter(existant: pd.DataFrame, nouveau: pd.DataFrame, cle: str, cols) -> tuple:
    """Ajoute à `existant` les lignes de `nouveau` dont la clé n'y est pas. Retourne (df, nb_ajoutes)."""
    deja = set(existant[cle].astype(str)) if not existant.empty else set()
    a_ajouter = nouveau[~nouveau[cle].astype(str).isin(deja)][cols]
    morceaux = [d for d in (existant[cols] if not existant.empty else None, a_ajouter if not a_ajouter.empty else None)
                if d is not None]
    df = pd.concat(morceaux, ignore_index=True) if morceaux else pd.DataFrame(columns=cols)
    return df, len(a_ajouter)


def importer(contrats, bookings, conteneurs, sauvegarde: bool = True) -> dict:
    """Écrit les 3 tables (ajout seulement). Peut lever PermissionError si le fichier est verrouillé."""
    ex.ensure_file()
    res = {"sauvegarde": ""}
    if sauvegarde:
        dest = ex.FICHIER.replace(".xlsx", f"_sauvegarde_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        shutil.copy2(ex.FICHIER, dest)
        res["sauvegarde"] = os.path.basename(dest)
    c_df, res["contrats"] = _ajouter(ex.charger_contracts(), contrats, "Contract_No", ex.COLS_CONTRACTS)
    ex.sauver_contracts(c_df)
    b_df, res["bookings"] = _ajouter(ex.charger_bookings(), bookings, "Booking_No", ex.COLS_BOOKINGS)
    ex.sauver_bookings(b_df)
    k_df, res["conteneurs"] = _ajouter(ex.charger_containers(), conteneurs, "Container_No", ex.COLS_CTR)
    ex.sauver_containers(k_df)
    return res


# ====================== INTERFACE ======================
if __name__ != "__main__" or True:   # (toujours exécuté par Streamlit)
    import streamlit as st

    st.set_page_config(page_title="Import ancien export", page_icon="📥", layout="wide")
    ss = st.session_state

    st.title("📥 Import de l'ancien fichier Export")
    st.caption("« Export General Report » → contrats, bookings et conteneurs de la base `export_data.xlsx`")

    flash = ss.pop("imp_flash", None)
    if flash:
        getattr(st, flash[0])(flash[1])

    fichier = st.file_uploader("Dépose l'ancien fichier Excel (Export General Report …)", type=["xlsx", "xlsm"])
    if not fichier:
        st.info("Dépose le fichier : les feuilles sont reconnues automatiquement, tu vérifies, puis tu importes. "
                "Rien n'est écrit avant que tu cliques sur « Importer ».")
        st.stop()

    try:
        wb = charger_classeur(fichier.getvalue())
        trackers, conteneurs = detecter_feuilles(wb)
    except Exception as e:                                     # noqa
        st.error(f"Fichier illisible : {e}")
        st.stop()
    if not trackers:
        st.error("Aucune feuille « master tracker » reconnue (colonnes Contract No, Bookings, Container Booked, "
                 "Container loaded).")
        st.stop()

    with st.expander("🗂️ Feuilles utilisées", expanded=False):
        feuille_trk = st.selectbox("Master tracker (bookings)", [t["feuille"] for t in trackers])
        feuilles_ctr = st.multiselect("Rapports de chargement (conteneurs)", [c["feuille"] for c in conteneurs],
                                      default=[c["feuille"] for c in conteneurs if c["visible"]])
        aligner = st.checkbox("Si le nombre de conteneurs du rapport diffère du tracker, utiliser celui du rapport "
                              "pour « Ctr_Loaded »", value=False)

    trk = lire_tracker(wb[feuille_trk])
    ctr = (pd.concat([lire_conteneurs(wb[f]) for f in feuilles_ctr], ignore_index=True)
           if feuilles_ctr else pd.DataFrame())
    if trk.empty:
        st.error("Le master tracker ne contient aucun booking.")
        st.stop()
    contrats, bookings, conteneurs_df, controles = construire_tout(trk, ctr, aligner)

    # --- déjà dans la base ?
    try:
        deja_c = set(ex.charger_contracts()["Contract_No"]) if not ex.charger_contracts().empty else set()
        deja_b = set(ex.charger_bookings()["Booking_No"].astype(str)) if not ex.charger_bookings().empty else set()
        deja_k = set(ex.charger_containers()["Container_No"]) if not ex.charger_containers().empty else set()
    except FileNotFoundError as e:
        st.error(str(e))
        st.stop()
    n_c = int((~contrats["Contract_No"].isin(deja_c)).sum())
    n_b = int((~bookings["Booking_No"].astype(str).isin(deja_b)).sum())
    n_k = int((~conteneurs_df["Container_No"].isin(deja_k)).sum()) if not conteneurs_df.empty else 0

    m1, m2, m3 = st.columns(3)
    m1.metric("Contrats", len(contrats), f"{n_c} nouveau(x)", delta_color="off")
    m2.metric("Bookings", len(bookings), f"{n_b} nouveau(x)", delta_color="off")
    m3.metric("Conteneurs", len(conteneurs_df), f"{n_k} nouveau(x)", delta_color="off")

    t1, t2, t3, t4 = st.tabs(["📑 Contrats (à compléter)", "🚢 Bookings", "📦 Conteneurs",
                              f"⚠️ Contrôles ({len(controles)})"])
    with t1:
        st.caption("L'ancien fichier ne contient ni la quantité ni la date des contrats : complète-les ici "
                   "avant d'importer (colonne « Chargé (MT) » = repère). Les autres colonnes restent modifiables.")
        aff_c = contrats.copy()
        aff_c.insert(0, "Statut", ["Déjà dans la base" if c in deja_c else "Nouveau" for c in aff_c["Contract_No"]])
        contrats_edit = st.data_editor(
            aff_c, key="imp_contrats", hide_index=True, width="stretch", num_rows="fixed",
            disabled=["Statut", "Contract_No", "Buyer", "Commodity", "_mt_charge"],
            column_config={
                "Date_Contract": st.column_config.DateColumn("Date_Contract", format="DD/MM/YYYY"),
                "Qty_Contract_MT": st.column_config.NumberColumn("Qty_Contract_MT (MT)", format="%.3f"),
                "Price_Unit": st.column_config.NumberColumn("Price_Unit", format="%.2f"),
                "_mt_charge": st.column_config.NumberColumn("Chargé (MT)", format="%.3f"),
            })
    with t2:
        aff_b = bookings.copy()
        aff_b.insert(0, "Statut", ["Déjà dans la base" if str(b) in deja_b else "Nouveau" for b in aff_b["Booking_No"]])
        for c in ("ETD", "ETA", "Date_Creation"):
            aff_b[c] = pd.to_datetime(aff_b[c], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(aff_b, width="stretch", hide_index=True, height=420)
    with t3:
        if conteneurs_df.empty:
            st.info("Aucun rapport de chargement sélectionné : seuls les contrats et bookings seront importés.")
        else:
            aff_k = conteneurs_df.copy()
            aff_k.insert(0, "Statut", ["Déjà dans la base" if c in deja_k else "Nouveau" for c in aff_k["Container_No"]])
            aff_k["Date_Stuffing"] = pd.to_datetime(aff_k["Date_Stuffing"], errors="coerce").dt.strftime("%d/%m/%Y")
            st.dataframe(aff_k, width="stretch", hide_index=True, height=420)
    with t4:
        if controles:
            st.dataframe(pd.DataFrame(controles), width="stretch", hide_index=True)
        else:
            st.success("Aucune anomalie détectée.")

    st.divider()
    sauvegarde = st.checkbox("Faire une copie de sauvegarde d'export_data.xlsx avant l'import", value=True)
    st.caption("Seuls les éléments « Nouveau » sont ajoutés. Rien d'existant n'est modifié ni supprimé. "
               "Ferme export_data.xlsx dans Excel avant d'importer.")
    if st.button("📥 Importer dans export_data.xlsx", type="primary", disabled=(n_c + n_b + n_k) == 0):
        try:
            c_final = contrats_edit.drop(columns=["Statut", "_mt_charge"])
            c_final["Date_Contract"] = pd.to_datetime(c_final["Date_Contract"], errors="coerce")
            res = importer(c_final, bookings, conteneurs_df, sauvegarde)
        except PermissionError as e:
            st.error(str(e) or "Impossible d'écrire : ferme export_data.xlsx dans Excel puis réessaie.")
        except Exception as e:                                 # noqa
            st.error(f"Import interrompu : {e}")
        else:
            msg = (f"Import terminé : {res['contrats']} contrat(s), {res['bookings']} booking(s), "
                   f"{res['conteneurs']} conteneur(s) ajouté(s).")
            if res["sauvegarde"]:
                msg += f" Sauvegarde : {res['sauvegarde']}"
            ss["imp_flash"] = ("success", msg)
            st.rerun()
