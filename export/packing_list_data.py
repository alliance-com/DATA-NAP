# -*- coding: utf-8 -*-
"""
Packing List NAP SARL — base Excel + liaison Bookings / Containers
À placer dans : DATA/export/packing_list_data.py
Fichier données : DATA/packing_lists.xlsx (racine projet)
"""
from __future__ import annotations

import os
from datetime import datetime, date
from io import BytesIO

import pandas as pd

# Dossier export/ → racine projet = parent
_EXPORT_DIR = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_EXPORT_DIR)  # DATA/
FICHIER = os.path.join(_ROOT, "packing_lists.xlsx")

FEUILLE = "PACKING"
FEUILLE_CTR = "PACKING_CTR"
VERSION = 1

COLS = [
    "ID", "Packing_No", "Date_PL",
    "Booking_No", "BL_No", "Dated_BL",
    "Shipper", "Shipper_Address",
    "Consignee", "Consignee_Address",
    "Notify", "Notify_Address",
    "Port_Loading", "Port_Discharge", "Country_Origin", "Final_Destination",
    "Vessel_Name", "Voyage",
    "Commodity", "Product_Lines",
    "Hs_Code",
    "Heating_Pad",
    "Notes",
    "Statut", "Cree_Par", "Heure_Creation",
]

COLS_CTR = [
    "ID", "Packing_ID", "S_No", "Container_No", "Seal_No",
    "Tan_No", "Rotation", "Bags", "Gross_KG", "Net_KG",
]

DEFAULT_SHIPPER = "STE NAP SARL"
DEFAULT_SHIPPER_ADDR = "PLOT NO 14 ,GODJI , BENIN"
DEFAULT_POL = "COTONOU, BENIN"
DEFAULT_ORIGIN = "BENIN"


def _s(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none", "nat") else t


def _n(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(str(v).replace(" ", "").replace(",", "."))
    except Exception:
        return 0.0


def ensure_file():
    if os.path.exists(FICHIER):
        return
    with pd.ExcelWriter(FICHIER, engine="openpyxl") as w:
        pd.DataFrame(columns=COLS).to_excel(w, sheet_name=FEUILLE, index=False)
        pd.DataFrame(columns=COLS_CTR).to_excel(w, sheet_name=FEUILLE_CTR, index=False)


def charger():
    ensure_file()
    try:
        df = pd.read_excel(FICHIER, sheet_name=FEUILLE)
    except Exception:
        return pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in df.columns:
            df[c] = 0 if c == "ID" else ""
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
    df["Date_PL"] = pd.to_datetime(df["Date_PL"], errors="coerce", dayfirst=True)
    df["Dated_BL"] = pd.to_datetime(df["Dated_BL"], errors="coerce", dayfirst=True)
    for c in COLS:
        if c not in ("ID", "Date_PL", "Dated_BL"):
            df[c] = df[c].map(_s)
    return df.sort_values(["Date_PL", "ID"], ascending=[False, False]).reset_index(drop=True)


def charger_containers_pl(packing_id=None):
    ensure_file()
    try:
        df = pd.read_excel(FICHIER, sheet_name=FEUILLE_CTR)
    except Exception:
        return pd.DataFrame(columns=COLS_CTR)
    for c in COLS_CTR:
        if c not in df.columns:
            df[c] = 0 if c in ("ID", "Packing_ID", "S_No", "Bags") else (0.0 if c in ("Gross_KG", "Net_KG") else "")
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
    df["Packing_ID"] = pd.to_numeric(df["Packing_ID"], errors="coerce").fillna(0).astype(int)
    df["S_No"] = pd.to_numeric(df["S_No"], errors="coerce").fillna(0).astype(int)
    df["Bags"] = pd.to_numeric(df["Bags"], errors="coerce").fillna(0).astype(int)
    for c in ("Gross_KG", "Net_KG"):
        df[c] = df[c].map(_n)
    for c in ("Container_No", "Seal_No", "Tan_No", "Rotation"):
        df[c] = df[c].map(_s)
    if packing_id is not None:
        df = df[df["Packing_ID"] == int(packing_id)]
    return df.reset_index(drop=True)


def sauver(df, ctr_df=None):
    ensure_file()
    out = df.copy()
    for c in COLS:
        if c not in out.columns:
            out[c] = ""
    out = out[COLS]
    if ctr_df is None:
        ctr_df = charger_containers_pl()
    else:
        for c in COLS_CTR:
            if c not in ctr_df.columns:
                ctr_df[c] = ""
        ctr_df = ctr_df[COLS_CTR]
    tmp = FICHIER.replace(".xlsx", ".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as w:
        out.to_excel(w, sheet_name=FEUILLE, index=False)
        ctr_df.to_excel(w, sheet_name=FEUILLE_CTR, index=False)
    import time, gc
    gc.collect()
    last = None
    for _ in range(6):
        try:
            os.replace(tmp, FICHIER)
            return
        except PermissionError as e:
            last = e
            time.sleep(0.5)
    raise PermissionError(str(last))


def prochain_id(df=None):
    if df is None:
        df = charger()
    return 1 if df.empty else int(df["ID"].max()) + 1


def prochain_numero(annee=None):
    annee = annee or date.today().year
    df = charger()
    n = 1
    if not df.empty:
        for r in df["Packing_No"].astype(str):
            parts = r.replace(" ", "").split("-")
            if len(parts) >= 3 and parts[1].isdigit() and int(parts[1]) == int(annee):
                try:
                    n = max(n, int(parts[2]) + 1)
                except Exception:
                    pass
    return f"PL-{annee}-{n:03d}"


def _import_export_data():
    """Charge export_data depuis le même dossier export/."""
    try:
        from export import export_data as ed
        return ed
    except ImportError:
        pass
    try:
        import export_data as ed
        return ed
    except ImportError:
        pass
    # même dossier que ce fichier
    import importlib.util
    path = os.path.join(_EXPORT_DIR, "export_data.py")
    if os.path.exists(path):
        spec = importlib.util.spec_from_file_location("export_data", path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod
    return None




# Description produit type SI / Packing List (Shea Butter etc.)
DESC_TEMPLATES = {
    "CRUDE SHEA BUTTER": (
        "{n} CONTAINER EACH CONTAINER WITH HEATING PAD\n"
        "RAW CRUDE SHEA BUTTER IN FLEXITANK OF REP OF BENIN\n"
        "ORIGIN WEST AFRICA MATERIAL DESCRIPTION : FA010C\n"
        "CRUDE SHB\n"
        "HS CODE : {hs}"
    ),
    "SHEA BUTTER": (
        "{n} CONTAINER EACH CONTAINER WITH HEATING PAD\n"
        "RAW CRUDE SHEA BUTTER IN FLEXITANK OF REP OF BENIN\n"
        "ORIGIN WEST AFRICA MATERIAL DESCRIPTION : FA010C\n"
        "CRUDE SHB\n"
        "HS CODE : {hs}"
    ),
    "SHEA NUT": (
        "{n} CONTAINERS\n"
        "SHEA NUT BENIN ORIGIN\n"
        "ORIGIN WEST AFRICA\n"
        "HS CODE : {hs}"
    ),
    "SHEA MEAL": (
        "{n} CONTAINERS\n"
        "SHEA NUT MEAL BENIN ORIGIN\n"
        "ORIGIN WEST AFRICA\n"
        "HS CODE : {hs}"
    ),
}

HS_DEFAULT = {
    "CRUDE SHEA BUTTER": "151590990",
    "SHEA BUTTER": "151590990",
    "SHEA NUT": "120799",
    "SHEA MEAL": "230690",
}


def construire_description_pl(commodity, nb_containers, hs_code="", heating=True):
    """
    Bloc description Packing List / SI.
    Ex. :
      5 CONTAINER EACH CONTAINER WITH HEATING PAD
      RAW CRUDE SHEA BUTTER IN FLEXITANK OF REP OF BENIN
      ORIGIN WEST AFRICA MATERIAL DESCRIPTION : FA010C
      CRUDE SHB
      HS CODE : 151590990
    """
    n = int(nb_containers or 0) or 1
    cu = (commodity or "").strip().upper()
    hs = (hs_code or "").strip()
    # match template
    tpl_key = None
    for k in DESC_TEMPLATES:
        if k in cu:
            tpl_key = k
            break
    if not hs and tpl_key:
        hs = HS_DEFAULT.get(tpl_key, "")
    if tpl_key:
        text = DESC_TEMPLATES[tpl_key].format(n=n, hs=hs or "—")
        if not heating and "HEATING PAD" in text:
            text = text.replace(" EACH CONTAINER WITH HEATING PAD", "")
            text = text.replace("CONTAINER EACH CONTAINER", "CONTAINERS")
        return text.strip()
    # générique
    lines = [f"{n} CONTAINERS"]
    if heating:
        lines[0] = f"{n} CONTAINER EACH CONTAINER WITH HEATING PAD"
    if cu:
        lines.append(cu)
    lines.append("ORIGIN WEST AFRICA MATERIAL DESCRIPTION")
    if hs:
        lines.append(f"HS CODE : {hs}")
    return "\n".join(lines)



def prefill_from_booking(booking_no):
    """
    Préremplit la Packing List depuis UN booking :
    - en-tête (BL, navire, POD, buyer, commodity, rotation, quantités booking)
    - TOUS les conteneurs liés à ce Booking_No (ou BL_No)
    """
    ed = _import_export_data()
    if ed is None:
        return None, "Module export_data introuvable (dossier export/)."

    bdf = ed.charger_bookings()
    if bdf.empty:
        return None, "Aucun booking dans export_data.xlsx."

    key = str(booking_no).strip()
    m = bdf["Booking_No"].astype(str).str.strip() == key
    if not m.any():
        m = bdf["BL_No"].astype(str).str.strip() == key
    if not m.any():
        return None, f"Booking / BL « {booking_no} » introuvable."

    b = bdf.loc[m].iloc[0]
    bno = ed._s(b["Booking_No"])
    bl = ed._s(b["BL_No"]) or bno

    # --- TOUS les conteneurs du booking (T_Containers) ---
    # Colonnes typiques export_data : Container_No, Booking_No, BL_No, Seal, Weight_KG, Tank, ...
    cdf = ed.charger_containers()
    ctrs = []
    rot_booking = ed._s(b["Rotation"]) if "Rotation" in bdf.columns else ""

    def _col(df, *names):
        for n in names:
            if n in df.columns:
                return n
        return None

    if cdf is not None and not cdf.empty:
        # Normaliser clés texte
        cdf = cdf.copy()
        if "Booking_No" in cdf.columns:
            cdf["_bno"] = cdf["Booking_No"].map(lambda x: ed._s(x))
        else:
            cdf["_bno"] = ""
        if "BL_No" in cdf.columns:
            cdf["_bl"] = cdf["BL_No"].map(lambda x: ed._s(x))
        else:
            cdf["_bl"] = ""

        sub = cdf[cdf["_bno"] == bno]
        if sub.empty and bl:
            sub = cdf[cdf["_bl"] == bl]
        # parfois Booking_No stocké = BL
        if sub.empty:
            sub = cdf[cdf["_bno"] == bl]
        if sub.empty:
            sub = cdf[cdf["_bl"] == bno]

        c_seal = _col(sub if not sub.empty else cdf, "Seal", "Seal_No", "Scelle", "SEAL")
        c_tank = _col(sub if not sub.empty else cdf, "Tank", "Tan_No", "Tan", "TANK")
        c_w = _col(sub if not sub.empty else cdf, "Weight_KG", "Weight", "Gross_KG", "Net_KG")
        c_bags = _col(sub if not sub.empty else cdf, "Bags", "Nb_Bags", "BAGS")
        c_rot = _col(sub if not sub.empty else cdf, "Rotation", "ROTATION")
        c_no = _col(sub if not sub.empty else cdf, "Container_No", "Container", "Ctr_No")

        if not sub.empty and c_no:
            sub = sub.sort_values(c_no)
            for i, (_, row) in enumerate(sub.iterrows(), start=1):
                seal = ed._s(row[c_seal]) if c_seal else ""
                tan = ed._s(row[c_tank]) if c_tank else ""
                w = ed._n(row[c_w]) if c_w else 0.0
                bags = ed._i(row[c_bags]) if c_bags else 0
                rot = ed._s(row[c_rot]) if c_rot else rot_booking
                if not rot:
                    rot = rot_booking
                # Poids : privilégier colonnes distinctes Gross / Net si présentes
                c_gross = _col(sub, "Gross_KG", "Gross_Weight", "Weight_Gross", "Poids_Brut")
                c_net = _col(sub, "Net_KG", "Net_Weight", "Weight_Net", "Poids_Net")
                # Weight_KG seul = souvent le net cargo (ou brut selon saisie) — ne pas forcer égalité
                gross_v = ed._n(row[c_gross]) if c_gross else 0.0
                net_v = ed._n(row[c_net]) if c_net else 0.0
                w_one = ed._n(row[c_w]) if c_w else 0.0
                if gross_v <= 0 and net_v <= 0 and w_one > 0:
                    # Une seule colonne Weight_KG : on la met en NET (poids marchandise)
                    # GROSS reste à saisir ou sera complété via ratio booking si dispo
                    net_v = w_one
                    gross_v = 0.0
                elif gross_v <= 0 and net_v > 0:
                    gross_v = 0.0  # ne pas inventer
                elif net_v <= 0 and gross_v > 0:
                    net_v = 0.0
                ctrs.append({
                    "S_No": i,
                    "Container_No": ed._s(row[c_no]),
                    "Seal_No": seal,
                    "Tan_No": tan,
                    "Rotation": rot,
                    "Bags": bags,
                    "Gross_KG": gross_v,
                    "Net_KG": net_v,
                })

    # Répartir Nb_Bags du booking si les lignes n'ont pas de bags
    nb_bags = ed._i(b["Nb_Bags"]) if "Nb_Bags" in bdf.columns else 0
    if nb_bags and ctrs and all(int(c.get("Bags") or 0) == 0 for c in ctrs):
        per = nb_bags // len(ctrs)
        rem = nb_bags % len(ctrs)
        for i, c in enumerate(ctrs):
            c["Bags"] = per + (1 if i < rem else 0)

    # Totaux booking (référence — ne remplace PAS un poids déjà saisi sur un conteneur)
    net_b = ed._n(b["Net_Qty_KG"]) if "Net_Qty_KG" in bdf.columns else 0.0
    gross_b = ed._n(b["Gross_Qty_KG"]) if "Gross_Qty_KG" in bdf.columns else 0.0

    # Si NET manquant partout mais Weight était en NET : déjà fait plus haut
    # Si GROSS manquant mais NET présent + ratio booking connu → estimer gross par conteneur
    # UNIQUEMENT si booking a les deux totaux et chaque ligne a un net > 0
    if ctrs and gross_b > 0 and net_b > 0:
        ratio = gross_b / net_b if net_b else 1.0
        for c in ctrs:
            n = float(c.get("Net_KG") or 0)
            g = float(c.get("Gross_KG") or 0)
            if n > 0 and g <= 0:
                c["Gross_KG"] = round(n * ratio, 0)
    # Si aucun net sur les lignes mais booking net total : répartir (dernier recours)
    if ctrs and net_b > 0 and all(float(c.get("Net_KG") or 0) == 0 for c in ctrs):
        per = net_b / len(ctrs)
        for c in ctrs:
            c["Net_KG"] = round(per, 0)
        if gross_b > 0 and all(float(c.get("Gross_KG") or 0) == 0 for c in ctrs):
            per_g = gross_b / len(ctrs)
            for c in ctrs:
                c["Gross_KG"] = round(per_g, 0)
    # Si gross booking total et nets individuels OK mais gross lignes vides sans ratio utilisable
    elif ctrs and gross_b > 0 and all(float(c.get("Gross_KG") or 0) == 0 for c in ctrs):
        # ne pas inventer un gross = net : laisser 0 pour saisie manuelle
        pass

    # Si aucun conteneur en base mais Ctr_Loaded / Ctr_Booked > 0 : lignes vides à compléter
    n_expected = 0
    for col in ("Ctr_Loaded", "Ctr_Booked"):
        if col in bdf.columns:
            n_expected = max(n_expected, ed._i(b[col]))
    if not ctrs and n_expected > 0:
        rot_booking = ed._s(b["Rotation"]) if "Rotation" in bdf.columns else ""
        per_net = (net_b / n_expected) if net_b else 0
        per_gross = (gross_b / n_expected) if gross_b else 0
        per_bags = (nb_bags // n_expected) if nb_bags else 0
        rem_bags = (nb_bags % n_expected) if nb_bags else 0
        for i in range(1, n_expected + 1):
            ctrs.append({
                "S_No": i,
                "Container_No": "",
                "Seal_No": "",
                "Tan_No": "",
                "Rotation": rot_booking,
                "Bags": per_bags + (1 if i <= rem_bags else 0),
                "Gross_KG": round(per_gross, 0),
                "Net_KG": round(per_net, 0),
            })

    commodity = ed._s(b["Commodity"]) if "Commodity" in bdf.columns else ""
    n_ctr_desc = len(ctrs) if ctrs else (
        ed._i(b["Ctr_Loaded"]) if "Ctr_Loaded" in bdf.columns and ed._i(b["Ctr_Loaded"])
        else (ed._i(b["Ctr_Booked"]) if "Ctr_Booked" in bdf.columns else 1)
    )
    # HS : colonne si un jour ajoutée, sinon défaut produit
    hs_code = ""
    for col in ("Hs_Code", "HS_Code", "HS"):
        if col in bdf.columns:
            hs_code = ed._s(b[col])
            if hs_code:
                break
    heating = True
    if commodity:
        cu = commodity.upper()
        heating = ("SHEA" in cu and "BUTTER" in cu) or "HEATING" in cu
    product_lines_txt = construire_description_pl(commodity, n_ctr_desc, hs_code=hs_code, heating=heating)

    # Date BL : ETD si dispo
    dated = date.today()
    if "ETD" in bdf.columns and pd.notna(b["ETD"]):
        try:
            dated = pd.Timestamp(b["ETD"]).date()
        except Exception:
            dated = date.today()

    data = {
        "Booking_No": bno,
        "BL_No": bl,
        "Dated_BL": dated,
        "Shipper": DEFAULT_SHIPPER,
        "Shipper_Address": DEFAULT_SHIPPER_ADDR,
        "Consignee": ed._s(b["Buyer"]) if "Buyer" in bdf.columns else "",
        "Consignee_Address": "",
        "Notify": ed._s(b["Buyer"]) if "Buyer" in bdf.columns else "",
        "Notify_Address": "",
        "Port_Loading": DEFAULT_POL,
        "Port_Discharge": ed._s(b["POD"]) if "POD" in bdf.columns else "",
        "Country_Origin": DEFAULT_ORIGIN,
        "Final_Destination": ed._s(b["POD"]) if "POD" in bdf.columns else "",
        "Vessel_Name": ed._s(b["Vessel"]) if "Vessel" in bdf.columns else "",
        "Voyage": ed._s(b["VOY"]) if "VOY" in bdf.columns else "",
        "Commodity": commodity,
        "Product_Lines": product_lines_txt,
        "Hs_Code": hs_code or HS_DEFAULT.get(
            next((k for k in HS_DEFAULT if k in commodity.upper()), ""),
            "151590990" if "SHEA" in commodity.upper() and "BUTTER" in commodity.upper() else "",
        ),
        "Heating_Pad": "OUI" if heating else "NON",
        "containers": ctrs,
        # infos utiles UI
        "_nb_containers": len(ctrs),
        "_net_booking_kg": net_b,
        "_gross_booking_kg": gross_b,
        "_ctr_booked": ed._i(b["Ctr_Booked"]) if "Ctr_Booked" in bdf.columns else 0,
        "_ctr_loaded": ed._i(b["Ctr_Loaded"]) if "Ctr_Loaded" in bdf.columns else 0,
    }
    if not ctrs:
        return data, (
            f"Booking {bno} trouvé, mais **aucun conteneur** dans T_Containers. "
            "Ajoutez les conteneurs sur le booking ou saisissez-les dans le formulaire."
        )
    return data, None



def ajouter(data, containers, user=""):
    df = charger()
    pid = prochain_id(df)
    num = _s(data.get("Packing_No")) or prochain_numero()
    row = {
        "ID": pid,
        "Packing_No": num,
        "Date_PL": pd.Timestamp(data.get("Date_PL") or date.today()),
        "Booking_No": _s(data.get("Booking_No")),
        "BL_No": _s(data.get("BL_No")),
        "Dated_BL": pd.to_datetime(data.get("Dated_BL"), errors="coerce", dayfirst=True),
        "Shipper": _s(data.get("Shipper") or DEFAULT_SHIPPER),
        "Shipper_Address": _s(data.get("Shipper_Address") or DEFAULT_SHIPPER_ADDR),
        "Consignee": _s(data.get("Consignee")),
        "Consignee_Address": _s(data.get("Consignee_Address")),
        "Notify": _s(data.get("Notify")),
        "Notify_Address": _s(data.get("Notify_Address")),
        "Port_Loading": _s(data.get("Port_Loading") or DEFAULT_POL),
        "Port_Discharge": _s(data.get("Port_Discharge")),
        "Country_Origin": _s(data.get("Country_Origin") or DEFAULT_ORIGIN),
        "Final_Destination": _s(data.get("Final_Destination")),
        "Vessel_Name": _s(data.get("Vessel_Name")),
        "Voyage": _s(data.get("Voyage")),
        "Commodity": _s(data.get("Commodity")),
        "Product_Lines": _s(data.get("Product_Lines")),
        "Hs_Code": _s(data.get("Hs_Code")),
        "Heating_Pad": _s(data.get("Heating_Pad") or "NON"),
        "Notes": _s(data.get("Notes")),
        "Statut": _s(data.get("Statut") or "EMIS"),
        "Cree_Par": _s(user),
        "Heure_Creation": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    ctr_all = charger_containers_pl()
    new_rows = []
    nid = int(ctr_all["ID"].max()) + 1 if not ctr_all.empty else 1
    for i, c in enumerate(containers or [], start=1):
        new_rows.append({
            "ID": nid, "Packing_ID": pid, "S_No": int(c.get("S_No") or i),
            "Container_No": _s(c.get("Container_No")),
            "Seal_No": _s(c.get("Seal_No")),
            "Tan_No": _s(c.get("Tan_No")),
            "Rotation": _s(c.get("Rotation")),
            "Bags": int(_n(c.get("Bags"))),
            "Gross_KG": _n(c.get("Gross_KG")),
            "Net_KG": _n(c.get("Net_KG")),
        })
        nid += 1
    if new_rows:
        ctr_all = pd.concat([ctr_all, pd.DataFrame(new_rows)], ignore_index=True)
    try:
        sauver(df, ctr_all)
    except PermissionError as e:
        return False, str(e), None
    return True, f"Packing List {num} enregistrée.", pid


def get_full(packing_id):
    df = charger()
    m = df["ID"] == int(packing_id)
    if not m.any():
        return None, None
    return df.loc[m].iloc[0].to_dict(), charger_containers_pl(packing_id)
