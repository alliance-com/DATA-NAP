# -*- coding: utf-8 -*-
"""Rapport des paiements filtré par Fournisseur / Bénéficiaire — Suivi de paiement."""
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

import rapport_paiements as rp

st.set_page_config(page_title="Rapport paiements — NAP SARL", page_icon="🧾", layout="wide")

BASE = Path(__file__).resolve().parent.parent  # dossier du projet (où est app.py)


@st.cache_data(show_spinner=False)
def _feuilles(data: bytes):
    return rp.lister_feuilles(data)


@st.cache_data(show_spinner=False)
def _lire(data: bytes, feuille: str):
    return rp.lire_feuille(data, feuille)


def _trouver_fichier():
    for p in sorted(BASE.glob("*.xls*")):
        n = rp.norm(p.name)
        if p.name.startswith("~$"):
            continue
        if "suivi" in n and "paiement" in n:
            return p
    return None


st.title("🧾 Rapport des paiements")
st.caption("Choisissez un ou plusieurs fournisseurs / bénéficiaires, puis imprimez en Excel ou PDF A3.")

# ───────── Fichier source ─────────
fichier = _trouver_fichier()
data, nom_src = None, ""
if fichier:
    data, nom_src = fichier.read_bytes(), fichier.name

with st.expander("📂 Fichier source", expanded=data is None):
    up = st.file_uploader("Choisir un autre fichier Excel (optionnel)", type=["xlsx", "xlsm"])
    if up:
        data, nom_src = up.getvalue(), up.name
    if nom_src:
        st.caption(f"Source utilisée : **{nom_src}**")

if data is None:
    st.warning("Fichier « Suivi de paiement.xlsx » introuvable dans le dossier du projet. "
               "Déposez-le ci-dessus.")
    st.stop()

feuilles = _feuilles(data)
defaut = next((i for i, f in enumerate(feuilles) if rp.norm(f) == "depenses"), 0)
if defaut == 0:
    defaut = next((i for i, f in enumerate(feuilles) if "paiement" in rp.norm(f)), 0)
feuille = st.selectbox("Feuille", feuilles, index=defaut)

brut = _lire(data, feuille)
if brut.empty:
    st.error("Cette feuille est vide.")
    st.stop()

# ───────── Colonnes (détection automatique, modifiable) ─────────
cols_all = list(brut.columns)
AUCUNE = "— aucune —"
d_benef = rp.trouver(cols_all, rp.MOTS_BENEF)
d_date = rp.trouver(cols_all, ["date"], exclure=["expir", "naissance"])
d_mont = rp.trouver(cols_all, ["montant"]) or rp.trouver(cols_all, ["total", "somme"])

with st.expander("⚙️ Colonnes utilisées (détectées automatiquement)"):
    k1, k2, k3 = st.columns(3)
    c_benef = k1.selectbox("Fournisseur / Bénéficiaire", cols_all,
                           index=cols_all.index(d_benef) if d_benef in cols_all else 0)
    opt = [AUCUNE] + cols_all
    c_date = k2.selectbox("Date", opt, index=opt.index(d_date) if d_date in opt else 0)
    c_mont = k3.selectbox("Montant", opt, index=opt.index(d_mont) if d_mont in opt else 0)
    c_date = None if c_date == AUCUNE else c_date
    c_mont = None if c_mont == AUCUNE else c_mont
    cols_aff = st.multiselect("Colonnes à afficher / imprimer", cols_all, default=cols_all)

if not cols_aff:
    st.warning("Choisissez au moins une colonne à afficher.")
    st.stop()
if c_benef not in cols_aff:
    cols_aff = [c_benef] + cols_aff

# ───────── Nettoyage ─────────
df = brut.copy()
df[c_benef] = df[c_benef].map(lambda v: "" if pd.isna(v) else str(v).strip())
df = df[df[c_benef] != ""]
if c_mont:
    df[c_mont] = rp.to_num(df[c_mont])
if c_date:
    df[c_date] = pd.to_datetime(df[c_date], errors="coerce", dayfirst=True)

# ───────── Filtres ─────────
f1, f2, f3 = st.columns([3, 2, 2])
with f1:
    options = sorted(df[c_benef].unique(), key=lambda x: rp.norm(x))
    choix = st.multiselect("Fournisseur / Bénéficiaire", options,
                           placeholder="Tous (ou choisissez un ou plusieurs noms)")
with f2:
    periode = None
    if c_date and df[c_date].notna().any():
        actif = st.checkbox("Filtrer par période")
        dmin, dmax = df[c_date].min().date(), df[c_date].max().date()
        if actif:
            periode = st.date_input("Période", value=(dmin, dmax), format="DD/MM/YYYY")
with f3:
    recherche = st.text_input("🔎 Recherche libre", placeholder="Ex. camion, chèque, ECOBANK…").strip()

dff = df
if choix:
    dff = dff[dff[c_benef].isin(choix)]
filtres = []
filtres.append("Fournisseur/Bénéficiaire : " + (", ".join(choix) if choix else "Tous"))
if periode and len(periode) == 2:
    dff = dff[dff[c_date].dt.date.between(periode[0], periode[1])]
    filtres.append(f"Période : {periode[0].strftime('%d/%m/%Y')} → {periode[1].strftime('%d/%m/%Y')}")
if recherche:
    masque = dff[cols_aff].astype(str).apply(
        lambda s: s.str.contains(recherche, case=False, regex=False)).any(axis=1)
    dff = dff[masque]
    filtres.append(f"Recherche : « {recherche} »")
filtres_txt = "  |  ".join(filtres)

if dff.empty:
    st.info("Aucun paiement ne correspond à ces filtres.")
    st.stop()

tri = [c_benef] + ([c_date] if c_date else [])
dff = dff.sort_values(tri, kind="stable").reset_index(drop=True)

if not choix:
    titre_benef = "Tous les fournisseurs / bénéficiaires"
elif len(choix) <= 3:
    titre_benef = " · ".join(choix)
else:
    titre_benef = f"{len(choix)} fournisseurs / bénéficiaires"

# ───────── Résumé ─────────
m1, m2, m3 = st.columns(3)
m1.metric("Paiements", f"{len(dff)}")
m2.metric("Fournisseurs / Bénéficiaires", f"{dff[c_benef].nunique()}")
if c_mont:
    m3.metric("Total payé (FCFA)", rp.fmt_montant(dff[c_mont].sum()))

synth = rp.synthese(dff, c_benef, c_mont)
if len(synth) > 1:
    st.markdown("#### Synthèse par fournisseur / bénéficiaire")
    s_aff = synth.copy()
    if "Total" in s_aff.columns:
        s_aff["Total"] = s_aff["Total"].map(rp.fmt_montant)
    st.dataframe(s_aff, hide_index=True, width="stretch")

st.markdown("#### Détail des paiements")
st.dataframe(rp.table_texte(dff, cols_aff, c_mont), hide_index=True, width="stretch")

# ───────── Impression ─────────
st.markdown("---")
slug = "".join(ch if ch.isalnum() else "_" for ch in (choix[0] if len(choix) == 1 else "Tous"))[:30]
stamp = date.today().strftime("%Y%m%d")

b1, b2 = st.columns(2)
with b1:
    try:
        st.download_button(
            "⬇️ Imprimer en Excel",
            data=rp.excel_bytes(dff, cols_aff, c_benef, c_mont, titre_benef, filtres_txt, synth),
            file_name=f"Rapport_Paiements_{slug}_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"Excel : {e}")
with b2:
    try:
        st.download_button(
            "🖨️ Imprimer en PDF A3",
            data=rp.pdf_a3(dff, cols_aff, c_benef, c_mont, titre_benef, filtres_txt, synth),
            file_name=f"Rapport_Paiements_{slug}_{stamp}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"PDF : {e}")
