# -*- coding: utf-8 -*-
"""Export — Impression A4 / A3"""
import streamlit as st
import pandas as pd
from datetime import date

import sys
import os
_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
from export import export_data as ex
from generer_export_print import RAPPORTS, generer_rapport, choisir_format

st.set_page_config(page_title="Export Impression", page_icon="🖨️", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

st.title("🖨️ Export — Impression")
st.caption(
    "PDF professionnel · **A4** pour les tableaux courts · **A3** pour les vues larges (bookings, recap). "
    "Le format peut être forcé ou laissé en automatique."
)

try:
    df_b = ex.charger_bookings()
    df_c = ex.charger_containers()
    df_ct = ex.charger_contracts()
except FileNotFoundError as e:
    st.error(str(e))
    st.stop()

# ---- Choix du rapport ----
options = {
    "Bookings / BL (complet)": "bookings",
    "Recap statuts": "recap",
    "Bookings en attente d'empotage": "stuffing_pending",
    "Conteneurs": "containers",
    "Contrats": "contracts",
}

r1, r2, r3 = st.columns([2, 1, 1])
with r1:
    choix = st.selectbox("Données à imprimer", list(options.keys()))
with r2:
    format_sel = st.selectbox(
        "Format feuille",
        ["Auto", "A4", "A3"],
        help="Auto : A3 si ≥ 10 colonnes, sinon A4. Toujours en paysage.",
    )
with r3:
    st.write("")
    st.write("")
    # indicateur format auto
    kind = options[choix]
    n_cols = len(RAPPORTS[kind]["headers"])
    auto_fmt = choisir_format(n_cols, None)
    st.info(f"Colonnes : **{n_cols}** → Auto = **{auto_fmt}**")

# ---- Filtres selon type ----
st.markdown("#### Filtres")
f1, f2, f3 = st.columns(3)
with f1:
    commodities = ["Tous"] + (sorted(df_b["Commodity"].dropna().unique().tolist()) if not df_b.empty else [])
    f_comm = st.selectbox("Commodity", commodities)
with f2:
    buyers = ["Tous"] + (sorted(df_b["Buyer"].dropna().unique().tolist()) if not df_b.empty else [])
    f_buyer = st.selectbox("Buyer", buyers)
with f3:
    fwds = ["Tous"] + (sorted(df_b["Forwarder"].dropna().unique().tolist()) if not df_b.empty else [])
    f_fwd = st.selectbox("Forwarder", fwds)

def _filtre_bookings(df):
    if df is None or df.empty:
        return df
    out = df.copy()
    if f_comm != "Tous":
        out = out[out["Commodity"] == f_comm]
    if f_buyer != "Tous":
        out = out[out["Buyer"] == f_buyer]
    if f_fwd != "Tous":
        out = out[out["Forwarder"] == f_fwd]
    return out

kind = options[choix]
if kind in ("bookings", "recap"):
    data = _filtre_bookings(df_b)
elif kind == "stuffing_pending":
    data = ex.bookings_pending_stuffing(_filtre_bookings(df_b))
elif kind == "containers":
    data = df_c
    if not data.empty and not df_b.empty:
        # optionnel : filtrer conteneurs via bookings filtrés
        bset = set(_filtre_bookings(df_b)["Booking_No"].astype(str))
        if f_comm != "Tous" or f_buyer != "Tous" or f_fwd != "Tous":
            data = data[data["Booking_No"].astype(str).isin(bset)]
elif kind == "contracts":
    data = df_ct
    if not data.empty and f_comm != "Tous":
        data = data[data["Commodity"] == f_comm]
    if not data.empty and f_buyer != "Tous":
        data = data[data["Buyer"] == f_buyer]
else:
    data = df_b

st.markdown("#### Aperçu")
if data is None or data.empty:
    st.warning("Aucune ligne pour ce filtre.")
else:
    st.success(f"**{len(data)} ligne(s)** prêtes pour impression")
    # aperçu limité
    aff = data.head(30).copy()
    for c in aff.columns:
        if "Date" in c or c in ("ETD", "ETA", "Date_Creation", "Date_Contract", "Date_Stuffing"):
            aff[c] = pd.to_datetime(aff[c], errors="coerce").dt.strftime("%d/%m/%Y")
    st.dataframe(aff, hide_index=True, width="stretch", height=280)

st.markdown("---")
fmt_force = None if format_sel == "Auto" else format_sel
filtre_txt = " · ".join(
    x for x in [
        None if f_comm == "Tous" else f_comm,
        None if f_buyer == "Tous" else f_buyer,
        None if f_fwd == "Tous" else f_fwd,
    ] if x
) or "Tous filtres"

col_btn, col_dl = st.columns([1, 2])
with col_btn:
    go = st.button("🖨️ Générer le PDF", type="primary", use_container_width=True, disabled=data is None or data.empty)

if go and data is not None and not data.empty:
    try:
        # generer_export_print peut être à la racine ou dans export/
        try:
            from generer_export_print import generer_rapport as _gr
        except ImportError:
            from export.generer_export_print import generer_rapport as _gr
        pdf_bytes, fmt_used = _gr(kind, data, format_page=fmt_force, filtre_label=filtre_txt)
        nom = f"Export_{kind}_{fmt_used}_{date.today().strftime('%Y%m%d')}.pdf"
        st.success(f"PDF généré en **{fmt_used} paysage** — {len(data)} ligne(s)")
        st.download_button(
            f"⬇️ Télécharger ({fmt_used})",
            data=pdf_bytes,
            file_name=nom,
            mime="application/pdf",
            type="primary",
        )
    except Exception as e:
        st.error(f"Erreur impression : {e}")
        import traceback
        st.code(traceback.format_exc())

st.markdown("---")
st.markdown("""
**Règle de format**
| Situation | Format |
|-----------|--------|
| Tableau ≤ 9 colonnes (contrats, conteneurs, pending) | **A4** paysage |
| Tableau large (bookings complet, recap) | **A3** paysage |
| Choix manuel | Force A4 ou A3 |

Placez aussi `generer_export_print.py` à la **racine** du projet (ou dans `export/`).
""")
st.caption("NAP SARL — Impression Export A4/A3")
