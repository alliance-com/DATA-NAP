import streamlit as st
from datetime import datetime

import dechargement_data as dd
import export_base_excel as ex
import importlib
importlib.reload(dd)   # relit le fichier à chaque affichage (évite l'ancienne version en mémoire)
importlib.reload(ex)

st.set_page_config(page_title="Export Excel", page_icon="📤", layout="wide")

# --- Sécurité : recharge les fichiers modifiés et vérifie qu'ils sont à jour ---
if getattr(dd, "VERSION_DATA", 0) < 9:
    st.error(f"⚠️ Le fichier **dechargement_data.py** n'est pas à jour (fichier utilisé : "
             f"`{getattr(dd, '__file__', '?')}`). Remplacez-le par la dernière version, au même endroit que "
             "`app.py`, puis arrêtez et relancez Streamlit.")
    st.stop()

# ---------- Authentification (comme les autres pages) ----------
if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

st.title("📤 Export Excel — base reliée pour travailler hors du site")
st.markdown(
    "Un **seul classeur Excel** avec toutes les données (tous les produits) : déchargements, paiements, "
    "bons de commande, fournisseurs, journal des bons de paiement. Les feuilles sont **reliées entre elles** "
    "et les totaux sont des **formules** : vous pouvez filtrer, ajouter des lignes ou faire des tableaux croisés "
    "sans le site."
)

c1, c2 = st.columns([1, 2])
with c1:
    if st.button("⚙️ Générer la base Excel", type="primary", width="stretch"):
        with st.spinner("Préparation du classeur…"):
            try:
                st.session_state["base_xlsx"] = ex.construire_base(name)
                st.session_state["base_xlsx_heure"] = datetime.now()
                st.session_state.pop("base_xlsx_erreur", None)
            except Exception as e:
                import traceback
                st.session_state.pop("base_xlsx", None)
                st.session_state["base_xlsx_erreur"] = f"{e}\n\n{traceback.format_exc()}"
    if st.session_state.get("base_xlsx_erreur") and "base_xlsx" not in st.session_state:
        st.error("Erreur pendant l'export. Détail technique ci-dessous (à m'envoyer si besoin).")
        st.code(st.session_state["base_xlsx_erreur"])

    if "base_xlsx" in st.session_state:
        h = st.session_state["base_xlsx_heure"]
        st.success(f"Classeur prêt ({h.strftime('%d/%m/%Y %H:%M')})")
        st.download_button(
            "⬇️ Télécharger la base Excel",
            data=st.session_state["base_xlsx"],
            file_name=f"NAP_Base_Dechargements_{h.strftime('%Y-%m-%d_%Hh%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

with c2:
    st.markdown("#### Contenu du classeur")
    st.markdown(
        "| Feuille | Rôle |\n|---|---|\n"
        "| **LISEZ_MOI** | Mode d'emploi et liste des liaisons |\n"
        "| **Releve_Fournisseur** | Choisissez un fournisseur → totaux + toutes ses livraisons, solde cumulé |\n"
        "| **Dechargements** | 1 ligne = 1 camion (dû, payé, solde, statut : formules) |\n"
        "| **Paiements** | 1 ligne = 1 paiement approuvé sur une livraison |\n"
        "| **Liens_Bons** | Les bons de paiement utilisés par chaque paiement |\n"
        "| **Bons_Commande** | Commandé ↔ livré, écart, statut (EN ATTENTE / EN COURS / LIVRE) |\n"
        "| **Fournisseurs** | Camions, poids, dû, payé, reste à payer par fournisseur |\n"
        "| **Journal_Bons** | Bons de paiement du journal, avec « rattaché à une livraison : OUI/NON » |"
    )

st.markdown("---")
st.info(
    "ℹ️ Le fichier est une **photo** des données au moment de l'export. Ce que vous modifiez dans Excel "
    "**ne remonte pas** dans le site : pour avoir les données à jour, générez un nouvel export."
)
with st.expander("🔗 Liaisons entre les feuilles (pour Power Query / Power Pivot)"):
    st.code(
        "Dechargements[ID_Fournisseur]   →  Fournisseurs[ID_Fournisseur]\n"
        "Dechargements[Ref_BC]           →  Bons_Commande[Reference]\n"
        "Paiements[ID_Livraison]         →  Dechargements[ID_Livraison]\n"
        "Liens_Bons[ID_Livraison]        →  Dechargements[ID_Livraison]\n"
        "Liens_Bons[No_Bon_Paiement]     →  Journal_Bons[No_Bon_Paiement]\n\n"
        "ID_Livraison = PRODUIT-Transaction (ex. SHEA-1004) : clé unique, car les numéros de "
        "transaction peuvent se répéter d'un produit à l'autre.",
        language="text",
    )

st.caption("NAP SARL — Export de la base Excel (tous produits)")
