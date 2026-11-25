# -*- coding: utf-8 -*-
from datetime import date
import streamlit as st
import pandas as pd
import presence_data as db
from generer_rapport_presence_a3 import generer_pdf_presence_a3

st.set_page_config(page_title="Rapport — Exposé N° 3", page_icon="📊", layout="wide")

user = st.session_state.get("user")
if not user:
    st.warning("Veuillez vous connecter.")
    st.page_link("pages/1_Connexion.py", label="🔑 Connexion")
    st.stop()

MOIS = ["", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
        "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

st.title("📊 Rapport de présence")

c1, c2 = st.columns(2)
with c1:
    annee = st.number_input("Année", 2024, 2035, date.today().year)
with c2:
    mois = st.selectbox("Mois", list(range(1, 13)), index=date.today().month - 1,
                        format_func=lambda m: MOIS[m])

dims, lignes = db.matrice_mois(int(annee), int(mois))
if not dims:
    st.warning("Aucun dimanche ce mois.")
    st.stop()

st.write("**Dimanches :** " + " · ".join(f"S{i+1} {d.strftime('%d/%m')}" for i, d in enumerate(dims)))

if lignes:
    rows = []
    for L in lignes:
        r = {"Nom": L["Nom"], "Prénom": L["Prenom"], "Profession": L.get("Profession", "")}
        for d in dims:
            r[d.strftime("%d/%m")] = {"V": "✓", "A": "…", "R": "✗"}.get(L.get(d.isoformat(), ""), "")
        r["Total ✓"] = L["Total_Valide"]
        rows.append(r)
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
else:
    st.info("Aucun membre.")

b1, b2 = st.columns(2)
with b1:
    st.download_button(
        "⬇️ Excel",
        data=db.to_excel_bytes_mois(int(annee), int(mois)),
        file_name=f"Presences_Expose3_Zogbadje_{int(mois):02d}_{int(annee)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with b2:
    try:
        pdf = generer_pdf_presence_a3(int(annee), int(mois), dims, lignes)
        st.download_button(
            "🖨️ PDF A3",
            data=pdf,
            file_name=f"Presences_Expose3_Zogbadje_{MOIS[int(mois)]}_{int(annee)}.pdf",
            mime="application/pdf",
            type="primary",
        )
    except Exception as e:
        st.error(f"PDF : {e}")
