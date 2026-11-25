# -*- coding: utf-8 -*-
"""Import Export General Report → export_master.xlsx"""
import streamlit as st
import tempfile
import os
import importlib
import export_data as ed
importlib.reload(ed)

st.set_page_config(page_title="Import Export", page_icon="📥", layout="wide")
if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

st.title("📥 Import — Export General Report")
st.caption(
    "Charge ton fichier Excel (GENERAL + stuffing SHEA BUTTER) dans `export_master.xlsx`. "
    "Les bookings et conteneurs alimentent Dashboard, Bookings et Conteneurs."
)

st.markdown("""
**Feuilles lues automatiquement**
| Feuille Excel | Destination |
|---------------|-------------|
| **GENERAL** | Bookings (1 ligne = 1 booking/BL) |
| **SHEA BUTTER** (stuffing) | Conteneurs (1 ligne = 1 conteneur) |
""")

mode = st.radio(
    "Mode d'import",
    ["Fusionner (garde l'existant, ignore les doublons)", "Remplacer tout (écrase export_master)"],
    index=0,
)
remplacer = mode.startswith("Remplacer")

fichier = st.file_uploader(
    "Fichier Export General Report (.xlsx)",
    type=["xlsx"],
    help="Ex. Export_General_Report_2026-2027.xlsx",
)

if fichier is not None:
    st.write(f"Fichier : **{fichier.name}** ({fichier.size // 1024} Ko)")
    if st.button("🚀 Lancer l'import", type="primary"):
        with st.spinner("Import en cours…"):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(fichier.getbuffer())
                    tmp_path = tmp.name
                ok, msg, n_bk, n_ctr = ed.importer_rapport_general(
                    tmp_path, user=name, remplacer=remplacer
                )
                os.unlink(tmp_path)
                if ok:
                    st.success(msg)
                    st.metric("Bookings en base", n_bk)
                    st.metric("Conteneurs en base", n_ctr)
                    st.info(
                        "Va sur **Export Dashboard**, **Export Bookings** et **Export Conteneurs** "
                        "pour voir les données. Recharge la page (F5) si besoin."
                    )
                else:
                    st.error(msg)
            except Exception as e:
                st.error(f"Erreur import : {e}")
                st.exception(e)

st.markdown("---")
bk = ed.charger_bookings()
ctr = ed.charger_conteneurs()
c1, c2 = st.columns(2)
c1.metric("Bookings actuels", len(bk))
c2.metric("Conteneurs actuels", len(ctr))
if not bk.empty:
    st.dataframe(
        bk[["Booking", "BL", "Forwarder", "Commodity", "Contract_No", "Buyer",
            "Container_Booked", "Container_Loaded", "Stuffing_Status"]].head(20),
        hide_index=True, width="stretch",
    )

st.caption("NAP SARL — Import Export General Report")
