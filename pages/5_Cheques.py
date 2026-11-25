# -*- coding: utf-8 -*-
"""
Page Streamlit — Suivi des chèques PAR BANQUE (feuilles séparées)
- 1 feuille Excel par banque dans registre_cheques.xlsx
- Synchronise depuis Suivi de paiement.xlsx (DEPENSES, Mode = CHEQUE)
- Ajout manuel + annulation par banque
"""
import streamlit as st
import pandas as pd
from datetime import datetime, date
import os
import re

st.set_page_config(page_title="Chèques", page_icon="🧾", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

# ==================== CONFIG ====================
FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"
FEUILLE_DEPENSES = "DEPENSES"
FICHIER_CHEQUES = "registre_cheques.xlsx"

BANQUES = [
    "ECOBANK BENIN",
    "BIIC",
    "ATLANTIC BANK BENIN",
    "BGFI",
    "CAISSE",
    "AUTRE",
]
STATUTS = ["EMIS", "ANNULE", "UTILISE", "EN ATTENTE"]

COLONNES = [
    "Date", "N° chèque", "N° transaction", "Type opération", "Produit",
    "Prestataire / Fournisseur", "Bénéficiaire", "Montant (FCFA)", "Statut",
    "Motif annulation / observation", "Source",
]
# =================================================


def _propre(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, (datetime, date, pd.Timestamp)):
        return v.strftime("%d/%m/%Y")
    t = str(v).strip()
    if t.lower() in ("", "nan", "none", "nat", "-"):
        return ""
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", t)
    if m:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    return t


def extraire_num_cheque(val):
    if val is None:
        return None
    s = str(val).strip().replace(" ", "").replace(",", "")
    if not s or s in ("-", "nan", "None"):
        return None
    m = re.search(r"(\d{4,})", s)
    return int(m.group(1)) if m else None


def normaliser_banque(nom):
    t = _propre(nom).upper()
    if not t:
        return "ECOBANK BENIN"
    for b in BANQUES:
        if b.upper() in t or t in b.upper():
            return b
    if "ECOBANK" in t:
        return "ECOBANK BENIN"
    if "ATLANTIC" in t:
        return "ATLANTIC BANK BENIN"
    if "BIIC" in t:
        return "BIIC"
    if "BGFI" in t:
        return "BGFI"
    if "CAISSE" in t:
        return "CAISSE"
    return "AUTRE"


def sheet_name(banque):
    return banque[:31]


def charger_depenses_cheques():
    if not os.path.exists(FICHIER_PAIEMENTS):
        return pd.DataFrame()
    try:
        xl = pd.ExcelFile(FICHIER_PAIEMENTS)
        sheet = FEUILLE_DEPENSES if FEUILLE_DEPENSES in xl.sheet_names else xl.sheet_names[0]
        df = None
        for hdr in (2, 1, 0, 3):
            try:
                tmp = pd.read_excel(FICHIER_PAIEMENTS, sheet_name=sheet, header=hdr)
                cols = " ".join(str(c).lower() for c in tmp.columns)
                if "cheque" in cols or "chèque" in cols or "transaction" in cols:
                    df = tmp
                    break
            except Exception:
                continue
        if df is None:
            return pd.DataFrame()

        df = df.dropna(axis=1, how="all")
        df = df.loc[:, [c for c in df.columns if not str(c).startswith("Unnamed")]]

        alias = {
            "date": "Date",
            "no de transaction": "N° transaction",
            "n° de transaction": "N° transaction",
            "type d'operation": "Type opération",
            "type d'opération": "Type opération",
            "produit": "Produit",
            "prestataires/fournisseurs": "Prestataire / Fournisseur",
            "nom du bénéficiaire": "Bénéficiaire",
            "montant": "Montant (FCFA)",
            "mode paiement": "Mode paiement",
            "numero de cheque": "N° chèque",
            "numéro de chèque": "N° chèque",
            "banque": "Banque",
            "observation": "Observation",
        }
        rename = {}
        for c in df.columns:
            k = str(c).strip().lower()
            if k in alias:
                rename[c] = alias[k]
        df = df.rename(columns=rename)

        if "Mode paiement" not in df.columns or "N° chèque" not in df.columns:
            return pd.DataFrame()

        mask = df["Mode paiement"].astype(str).str.upper().str.contains("CHEQUE", na=False)
        df = df[mask].copy()
        df["N° chèque"] = df["N° chèque"].map(_propre)
        df = df[df["N° chèque"] != ""]

        out = pd.DataFrame()
        n = len(df)
        empty = pd.Series([""] * n)
        out["Date"] = df.get("Date", empty).map(_propre)
        out["N° chèque"] = df["N° chèque"]
        out["Banque"] = df.get("Banque", empty).map(normaliser_banque)
        out["N° transaction"] = df.get("N° transaction", empty).map(_propre)
        out["Type opération"] = df.get("Type opération", empty).map(_propre)
        out["Produit"] = df.get("Produit", empty).map(_propre)
        out["Prestataire / Fournisseur"] = df.get("Prestataire / Fournisseur", empty).map(_propre)
        out["Bénéficiaire"] = df.get("Bénéficiaire", empty).map(_propre)
        out["Montant (FCFA)"] = pd.to_numeric(df.get("Montant (FCFA)", 0), errors="coerce").fillna(0)
        out["Statut"] = "EMIS"
        out["Motif annulation / observation"] = df.get("Observation", empty).map(_propre)
        out["Source"] = "AUTO"
        return out.reset_index(drop=True)
    except Exception as e:
        st.error(f"Erreur lecture Suivi de paiement : {e}")
        return pd.DataFrame()


def charger_feuille_banque(banque):
    if not os.path.exists(FICHIER_CHEQUES):
        return pd.DataFrame(columns=COLONNES)
    try:
        name = sheet_name(banque)
        xl = pd.ExcelFile(FICHIER_CHEQUES)
        if name not in xl.sheet_names:
            return pd.DataFrame(columns=COLONNES)
        df = pd.read_excel(FICHIER_CHEQUES, sheet_name=name, header=3)
        df = df.dropna(how="all")
        for c in COLONNES:
            if c not in df.columns:
                df[c] = ""
        df = df[COLONNES]
        df = df[df["N° chèque"].map(_propre) != ""]
        return df.reset_index(drop=True)
    except Exception:
        return pd.DataFrame(columns=COLONNES)


def fusion_banque(auto_df, reg_df, banque):
    auto_b = pd.DataFrame(columns=COLONNES)
    if auto_df is not None and not auto_df.empty:
        auto_b = auto_df[auto_df["Banque"].map(normaliser_banque) == banque].copy()
        if not auto_b.empty:
            auto_b = auto_b[COLONNES]

    if reg_df is None or reg_df.empty:
        return auto_b.reset_index(drop=True)

    def cle(row):
        return _propre(row.get("N° chèque", "")).upper()

    auto_keys = set(auto_b.apply(cle, axis=1)) if not auto_b.empty else set()
    manuels = []
    overrides = {}

    for _, row in reg_df.iterrows():
        k = cle(row)
        src = _propre(row.get("Source", "")).upper()
        statut = _propre(row.get("Statut", "")).upper()
        if statut == "ANNULE" and k in auto_keys:
            overrides[k] = row
        elif k not in auto_keys and (src == "MANUEL" or statut):
            manuels.append(row)

    if overrides and not auto_b.empty:
        for i, row in auto_b.iterrows():
            k = cle(row)
            if k in overrides:
                o = overrides[k]
                auto_b.at[i, "Statut"] = _propre(o.get("Statut")) or "ANNULE"
                auto_b.at[i, "Motif annulation / observation"] = _propre(
                    o.get("Motif annulation / observation")
                )
                auto_b.at[i, "Source"] = "AUTO+ANNULE"

    if manuels:
        mdf = pd.DataFrame(manuels)
        for c in COLONNES:
            if c not in mdf.columns:
                mdf[c] = ""
        auto_b = pd.concat([auto_b, mdf[COLONNES]], ignore_index=True)

    return auto_b.reset_index(drop=True)


def analyser_series_banque(df, banque):
    empty = {
        "Banque": banque, "N° min": "", "N° max": "", "Nb chèques": 0,
        "Nb annulés": 0, "Montant total (FCFA)": 0, "Trous (n° manquants)": "—",
    }
    if df is None or df.empty:
        return empty
    tmp = df.copy()
    tmp["_num"] = tmp["N° chèque"].map(extraire_num_cheque)
    tmp = tmp[tmp["_num"].notna()]
    nums = sorted(set(int(n) for n in tmp["_num"].tolist())) if not tmp.empty else []
    annul = df[df["Statut"].astype(str).str.upper() == "ANNULE"]
    montants = pd.to_numeric(df["Montant (FCFA)"], errors="coerce").fillna(0).sum()
    trous = []
    for a, b in zip(nums, nums[1:]):
        if b - a > 1:
            manquants = list(range(a + 1, b))
            if len(manquants) <= 12:
                trous.extend(manquants)
            else:
                trous.append(f"{manquants[0]}…{manquants[-1]} ({len(manquants)})")
    return {
        "Banque": banque,
        "N° min": nums[0] if nums else "",
        "N° max": nums[-1] if nums else "",
        "Nb chèques": len(nums),
        "Nb annulés": len(annul),
        "Montant total (FCFA)": montants,
        "Trous (n° manquants)": ", ".join(str(t) for t in trous) if trous else "—",
    }


def ecrire_feuille_banque(wb, banque, df):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    name = sheet_name(banque)
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)

    header_fill = PatternFill("solid", fgColor="1B4F72")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )

    ws["A1"] = f"CHÈQUES — {banque}"
    ws["A1"].font = Font(bold=True, size=14, color="1B4F72")
    ws["A2"] = f"Maj : {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = Font(name="Arial", size=8, italic=True, color="666666")

    for i, h in enumerate(COLONNES, 1):
        cell = ws.cell(4, i, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin
        cell.alignment = Alignment(wrap_text=True, horizontal="center")

    data = df if df is not None else pd.DataFrame(columns=COLONNES)
    for r_idx, row in enumerate(data.itertuples(index=False), 5):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(r_idx, c_idx, val if val == val else "")
            cell.border = thin
            if COLONNES[c_idx - 1] == "Montant (FCFA)":
                try:
                    cell.value = float(val)
                    cell.number_format = "#,##0"
                except Exception:
                    pass
            if COLONNES[c_idx - 1] == "Statut" and str(val).upper() == "ANNULE":
                cell.fill = PatternFill("solid", fgColor="FFCDD2")
                cell.font = Font(color="B71C1C", bold=True)

    widths = [12, 12, 16, 14, 12, 28, 22, 14, 12, 32, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def ecrire_synthese(wb, series_rows):
    from openpyxl.styles import Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    if "SYNTHESE" in wb.sheetnames:
        del wb["SYNTHESE"]
    ws = wb.create_sheet("SYNTHESE", 1)
    header_fill = PatternFill("solid", fgColor="1B4F72")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    ws["A1"] = "SYNTHÈSE — SÉRIES PAR BANQUE"
    ws["A1"].font = Font(bold=True, size=14, color="1B4F72")
    ws["A2"] = f"Recalculée le {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    h = ["Banque", "N° min", "N° max", "Nb chèques", "Nb annulés", "Montant total (FCFA)", "Trous (n° manquants)"]
    for i, x in enumerate(h, 1):
        cell = ws.cell(4, i, x)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin
    for r_idx, row in enumerate(series_rows, 5):
        for c_idx, key in enumerate(h, 1):
            cell = ws.cell(r_idx, c_idx, row.get(key, ""))
            cell.border = thin
            if key == "Montant total (FCFA)":
                try:
                    cell.value = float(row.get(key) or 0)
                    cell.number_format = "#,##0"
                except Exception:
                    pass
    for i, w in enumerate([22, 12, 12, 12, 12, 18, 45], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def synchroniser_tout():
    from openpyxl import load_workbook, Workbook

    auto = charger_depenses_cheques()
    series_rows = []
    data_par_banque = {}

    for banque in BANQUES:
        reg = charger_feuille_banque(banque)
        fusion = fusion_banque(auto, reg, banque)
        data_par_banque[banque] = fusion
        series_rows.append(analyser_series_banque(fusion, banque))

    if os.path.exists(FICHIER_CHEQUES):
        wb = load_workbook(FICHIER_CHEQUES)
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    for banque in BANQUES:
        ecrire_feuille_banque(wb, banque, data_par_banque[banque])
    ecrire_synthese(wb, series_rows)

    if "GUIDE" not in wb.sheetnames:
        ws = wb.create_sheet("GUIDE", 0)
        ws["A1"] = "Une feuille = une banque. Ne pas fusionner les banques."
        ws["A1"].font = __import__("openpyxl.styles", fromlist=["Font"]).Font(bold=True, size=12)

    wb.save(FICHIER_CHEQUES)
    return data_par_banque, series_rows


def sauver_ligne_manuelle(banque, ligne):
    from openpyxl import load_workbook, Workbook

    reg = charger_feuille_banque(banque)
    ligne = {c: ligne.get(c, "") for c in COLONNES}
    ligne["Source"] = ligne.get("Source") or "MANUEL"
    reg = pd.concat([reg, pd.DataFrame([ligne])], ignore_index=True)

    auto = charger_depenses_cheques()
    fusion = fusion_banque(auto, reg, banque)

    if os.path.exists(FICHIER_CHEQUES):
        wb = load_workbook(FICHIER_CHEQUES)
    else:
        wb = Workbook()
    ecrire_feuille_banque(wb, banque, fusion)
    series_rows = [
        analyser_series_banque(fusion if b == banque else charger_feuille_banque(b), b)
        for b in BANQUES
    ]
    ecrire_synthese(wb, series_rows)
    wb.save(FICHIER_CHEQUES)


# ==================== UI ====================
st.title("🧾 Suivi des chèques — une banque = une feuille")
st.caption(
    f"`{FICHIER_CHEQUES}` · Feuilles : "
    + " · ".join(BANQUES)
    + f" · Source : `{FICHIER_PAIEMENTS}`"
)

c_sync, c_info = st.columns([1, 3])
with c_sync:
    if st.button("🔄 Synchroniser depuis paiements", type="primary", width="stretch"):
        data_par, series = synchroniser_tout()
        total = sum(len(v) for v in data_par.values())
        st.success(f"OK — {total} chèque(s) répartis sur {len(BANQUES)} banques.")
        st.rerun()

data_par = {}
series_rows = []
for b in BANQUES:
    data_par[b] = charger_feuille_banque(b)
    series_rows.append(analyser_series_banque(data_par[b], b))

with c_info:
    total = sum(len(v) for v in data_par.values())
    n_ann = sum(
        len(v[v["Statut"].astype(str).str.upper() == "ANNULE"])
        for v in data_par.values() if not v.empty
    )
    st.info(f"**{total}** chèques · **{n_ann}** annulé(s) · **{len(BANQUES)}** banques séparées")

st.subheader("📊 Synthèse par banque")
st.dataframe(pd.DataFrame(series_rows), width="stretch", hide_index=True)

st.markdown("---")
st.subheader("🏦 Registre par banque")
tabs = st.tabs(BANQUES)

for tab, banque in zip(tabs, BANQUES):
    with tab:
        df_b = data_par.get(banque, pd.DataFrame(columns=COLONNES))
        s = analyser_series_banque(df_b, banque)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Nb chèques", s["Nb chèques"])
        m2.metric("Annulés", s["Nb annulés"])
        m3.metric("N° min → max", f"{s['N° min']} → {s['N° max']}" if s["N° min"] != "" else "—")
        m4.metric("Montant (FCFA)", f"{s['Montant total (FCFA)']:,.0f}".replace(",", " "))

        if s["Trous (n° manquants)"] and s["Trous (n° manquants)"] != "—":
            st.warning(f"Trous de numérotation : **{s['Trous (n° manquants)']}**")

        f1, f2, f3 = st.columns(3)
        with f1:
            fs = st.selectbox("Statut", ["Tous"] + STATUTS, key=f"st_{banque}")
        with f2:
            fsrc = st.selectbox("Source", ["Toutes", "AUTO", "MANUEL", "AUTO+ANNULE"], key=f"src_{banque}")
        with f3:
            q = st.text_input("Recherche", key=f"q_{banque}", placeholder="N° chèque, transaction…")

        df_f = df_b.copy()
        if not df_f.empty:
            if fs != "Tous":
                df_f = df_f[df_f["Statut"].astype(str).str.upper() == fs]
            if fsrc != "Toutes":
                df_f = df_f[df_f["Source"].astype(str).str.upper() == fsrc]
            if q.strip():
                qq = q.strip().lower()
                df_f = df_f[df_f.apply(lambda r: qq in " ".join(r.astype(str)).lower(), axis=1)]

        if df_f.empty:
            st.info(f"Aucun chèque pour **{banque}**. Synchronisez ou ajoutez manuellement.")
        else:
            st.dataframe(df_f, width="stretch", height=320, hide_index=True)

        st.markdown(f"##### 🚫 Annuler un chèque — {banque}")
        if not df_b.empty:
            opts = df_b.apply(
                lambda r: f"{_propre(r['N° chèque'])} | {_propre(r['N° transaction'])} | {_propre(r['Prestataire / Fournisseur'])}",
                axis=1,
            ).tolist()
            choix = st.selectbox("Chèque", [""] + opts, key=f"ann_{banque}")
            motif = st.text_input("Motif d'annulation", key=f"mot_{banque}")
            if st.button("Marquer ANNULE", key=f"btn_ann_{banque}") and choix:
                num = choix.split("|")[0].strip()
                parts = [p.strip() for p in choix.split("|")]
                ligne = {
                    "Date": datetime.now().strftime("%d/%m/%Y"),
                    "N° chèque": num,
                    "N° transaction": parts[1] if len(parts) > 1 else "",
                    "Type opération": "",
                    "Produit": "",
                    "Prestataire / Fournisseur": parts[2] if len(parts) > 2 else "",
                    "Bénéficiaire": "",
                    "Montant (FCFA)": 0,
                    "Statut": "ANNULE",
                    "Motif annulation / observation": motif,
                    "Source": "MANUEL",
                }
                reg = charger_feuille_banque(banque)
                if not reg.empty:
                    reg = reg[reg["N° chèque"].map(_propre) != num]
                reg = pd.concat([reg, pd.DataFrame([ligne])], ignore_index=True)
                auto = charger_depenses_cheques()
                fusion = fusion_banque(auto, reg, banque)
                from openpyxl import load_workbook, Workbook
                wb = load_workbook(FICHIER_CHEQUES) if os.path.exists(FICHIER_CHEQUES) else Workbook()
                ecrire_feuille_banque(wb, banque, fusion)
                series_rows2 = [
                    analyser_series_banque(fusion if b == banque else charger_feuille_banque(b), b)
                    for b in BANQUES
                ]
                ecrire_synthese(wb, series_rows2)
                wb.save(FICHIER_CHEQUES)
                st.success(f"{num} annulé sur {banque}.")
                st.rerun()

        st.markdown(f"##### ➕ Ajouter un chèque — {banque}")
        with st.form(f"form_{banque}", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                date_c = st.date_input("Date", value=date.today(), key=f"d_{banque}")
                num_c = st.text_input("N° chèque *", key=f"n_{banque}")
            with c2:
                trans_c = st.text_input("N° transaction", key=f"t_{banque}")
                montant_c = st.number_input("Montant (FCFA)", min_value=0.0, step=1000.0, key=f"m_{banque}")
            with c3:
                presta_c = st.text_input("Prestataire", key=f"p_{banque}")
                statut_c = st.selectbox("Statut", STATUTS, key=f"s_{banque}")
            motif_c = st.text_input("Motif / observation", key=f"o_{banque}")
            ok = st.form_submit_button("💾 Enregistrer", width="stretch")
            if ok:
                if not num_c.strip():
                    st.error("N° chèque obligatoire.")
                else:
                    ligne = {
                        "Date": date_c.strftime("%d/%m/%Y"),
                        "N° chèque": num_c.strip(),
                        "N° transaction": trans_c.strip(),
                        "Type opération": "",
                        "Produit": "",
                        "Prestataire / Fournisseur": presta_c.strip(),
                        "Bénéficiaire": "",
                        "Montant (FCFA)": montant_c,
                        "Statut": statut_c,
                        "Motif annulation / observation": motif_c.strip(),
                        "Source": "MANUEL",
                    }
                    sauver_ligne_manuelle(banque, ligne)
                    st.success(f"Chèque {num_c} ajouté sur **{banque}**.")
                    st.rerun()

st.markdown("---")
st.caption(
    "NAP SARL — 1 banque = 1 feuille Excel · "
    f"`{FICHIER_CHEQUES}` + `{FICHIER_PAIEMENTS}` · `pages/5_Cheques.py`"
)
