# -*- coding: utf-8 -*-
"""
Liste des paiements par produit / période
Source : Suivi de paiement.xlsx — Impression A3
"""
from __future__ import annotations

from datetime import date, timedelta

import pandas as pd
import streamlit as st

import liste_paiements_data as lpd
import generer_liste_paiements_a3 as glp

st.set_page_config(page_title="Liste paiements", page_icon="💳", layout="wide")

if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

st.title("💳 Liste des paiements par produit")
st.caption("Source : **Suivi de paiement.xlsx** · Filtres produit & période · Impression **A3**")

df, chemin, err = lpd.charger_paiements()
if err and (df is None or df.empty):
    st.error(err)
    st.info("Placez `Suivi de paiement.xlsx` à la racine du projet (à côté de app.py).")
    st.stop()

if chemin:
    st.sidebar.success(f"Fichier : `{chemin}`")

produits = ["Tous"] + lpd.liste_produits(df)

# Période par défaut — bornes saines (évite dates Excel aberrantes type année 1036)
LIMITE_MIN = date(2020, 1, 1)
LIMITE_MAX = date(date.today().year + 2, 12, 31)
today = date.today()

def _borne_date(d, default):
    try:
        if d is None or (hasattr(d, "__float__") and pd.isna(d)):
            return default
        if hasattr(d, "date"):
            d = d.date()
        if d.year < 2000 or d.year > today.year + 5:
            return default
        return d
    except Exception:
        return default

if not df.empty and df["Date"].notna().any():
    raw_min = df["Date"].min()
    raw_max = df["Date"].max()
    dmin = _borne_date(raw_min, date(today.year, 1, 1))
    dmax = _borne_date(raw_max, today)
else:
    dmin = date(today.year, 1, 1)
    dmax = today

# Garantir dmin <= dmax et dans [LIMITE_MIN, LIMITE_MAX]
dmin = max(dmin, LIMITE_MIN)
dmax = min(max(dmax, dmin), LIMITE_MAX)
if dmin > dmax:
    dmin, dmax = date(today.year, 1, 1), today

c1, c2, c3, c4 = st.columns([2, 1.5, 1.5, 1])
with c1:
    produit = st.selectbox("Produit", produits, key="lp_produit")
with c2:
    debut = st.date_input(
        "Du",
        value=dmin,
        min_value=LIMITE_MIN,
        max_value=LIMITE_MAX,
        key="lp_deb",
    )
with c3:
    fin = st.date_input(
        "Au",
        value=dmax,
        min_value=LIMITE_MIN,
        max_value=LIMITE_MAX,
        key="lp_fin",
    )
with c4:
    st.write("")
    st.write("")
    filtrer = st.button("Appliquer les filtres", type="primary")

if debut > fin:
    st.warning("La date de début est après la date de fin — les dates ont été inversées.")
    debut, fin = fin, debut

filtre = lpd.filtrer(df, produit=produit, date_debut=debut, date_fin=fin)

# KPI
k1, k2, k3, k4 = st.columns(4)
k1.metric("Lignes", f"{len(filtre):,}".replace(",", " "))
k2.metric("Montant total", f"{filtre['Montant'].sum():,.0f} FCFA".replace(",", " ") if not filtre.empty else "0")
k3.metric("Produit", produit)
k4.metric("Période", f"{debut.strftime('%d/%m/%Y')} → {fin.strftime('%d/%m/%Y')}")

st.markdown("---")

if filtre.empty:
    st.info("Aucun paiement pour ces filtres.")
else:
    aff = filtre.copy()
    aff["Date"] = pd.to_datetime(aff["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
    cols = [
        "Date", "No_Transaction", "Produit", "Prestataire", "Beneficiaire",
        "Montant", "Mode", "Banque", "Num_Cheque", "Observation", "Description",
    ]
    cols = [c for c in cols if c in aff.columns]
    st.dataframe(
        aff[cols],
        hide_index=True,
        width="stretch",
        height=min(520, 40 + 28 * min(len(aff), 15)),
        column_config={
            "No_Transaction": "N° Transaction",
            "Prestataire": "Prestataire / Fournisseur",
            "Beneficiaire": "Bénéficiaire",
            "Montant": st.column_config.NumberColumn("Montant", format="%d"),
            "Num_Cheque": "N° Chèque",
        },
    )

    periode_txt = f"{debut.strftime('%d/%m/%Y')} au {fin.strftime('%d/%m/%Y')}"
    b1, b2 = st.columns(2)
    with b1:
        st.download_button(
            "⬇️ Télécharger Excel (tableau formaté)",
            data=lpd.to_excel_bytes(filtre, produit=produit, periode=periode_txt),
            file_name=f"Paiements_{produit}_{debut}_{fin}.xlsx".replace(" ", "_"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    with b2:
        try:
            pdf = glp.generer_pdf_a3(
                filtre,
                periode=periode_txt,
                produit=produit,
            )
            st.download_button(
                "🖨️ Imprimer A3 (PDF)",
                data=pdf,
                file_name=f"Liste_Paiements_{produit}_{debut}_{fin}.pdf".replace(" ", "_"),
                mime="application/pdf",
                type="primary",
            )
        except Exception as e:
            st.error(f"PDF A3 : {e}")

st.caption("NAP SARL · Liste paiements · Suivi de paiement.xlsx")
