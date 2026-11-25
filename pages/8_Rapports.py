import streamlit as st
import pandas as pd
from datetime import date

import dechargement_data as dd
import rapport_data as rp
import importlib
importlib.reload(dd)   # relit le fichier à chaque affichage (évite l'ancienne version en mémoire)
importlib.reload(rp)

st.set_page_config(page_title="Rapports", page_icon="📊", layout="wide")

# --- Sécurité : recharge les fichiers modifiés et vérifie qu'ils sont à jour ---
if getattr(dd, "VERSION_DATA", 0) < 9:
    st.error(f"⚠️ Le fichier **dechargement_data.py** n'est pas à jour (fichier utilisé : "
             f"`{getattr(dd, '__file__', '?')}`). Remplacez-le par la dernière version, au même endroit que "
             "`app.py`, puis arrêtez et relancez Streamlit.")
    st.stop()
if getattr(rp, "VERSION_RAPPORT", 0) < 5:
    st.error("⚠️ Le fichier **rapport_data.py** n'est pas à jour (version trop ancienne). Fichier utilisé par "
             f"l'application : `{getattr(rp, '__file__', '?')}`. Remplacez-le par la dernière version, puis arrêtez "
             "et relancez Streamlit.")
    st.stop()

# ---------- Authentification (comme les autres pages) ----------
if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"
username = st.session_state.get("username") or ""
est_admin = username == "admin"


def flash(type_, message):
    st.session_state["_flash"] = (type_, message)


def montrer(df, config=None, hauteur=None, vide="Aucune activité enregistrée sur la période."):
    if df.empty:
        st.info(vide)
    else:
        options = {"height": hauteur} if hauteur else {}
        st.dataframe(df, hide_index=True, width="stretch", column_config=config or {}, **options)


CFG_DATE = {"Date": st.column_config.DateColumn("Date", format="DD/MM/YYYY")}
NUM = lambda titre: st.column_config.NumberColumn(titre, format="%d")

data = rp.charger_donnees()

st.title("📊 Rapports journalier & hebdomadaire")

if "_flash" in st.session_state:
    type_, message = st.session_state.pop("_flash")
    getattr(st, type_)(message)

# ==================== CHOIX DU RAPPORT ====================
produit = st.radio("📦 Produit concerné par le rapport", [None] + dd.PRODUITS, horizontal=True,
                   format_func=lambda p: "Tous les produits" if p is None else dd.libelle_produit(p),
                   help="Le rapport ne contient QUE les activités du produit choisi (bons de commande, "
                        "bons de paiement, déchargements, paiements approuvés, courrier).")
c1, c2, c3 = st.columns([1, 1.4, 1])
type_rap = c1.radio("Type de rapport", ["Journalier", "Hebdomadaire"], horizontal=True)
jour = c2.date_input("Jour concerné (hebdomadaire : n'importe quel jour de la semaine)", value=date.today(),
                     format="DD/MM/YYYY")
langue = c3.radio("🌐 Langue d'impression (PDF et Excel)", list(rp.LANGUES.keys()), horizontal=True,
                  format_func=lambda x: rp.LANGUES[x], key="langue_rapport",
                  help="Le rapport imprimé / téléchargé sera entièrement en anglais si vous choisissez English. "
                       "L'écran reste en français. Les observations sont imprimées telles que vous les saisissez.")

obs = st.text_area("Observations / commentaire à inclure dans le rapport (facultatif)", key="rap_obs", height=90,
                   placeholder="Ex. : démarrage du soja, relance des fournisseurs, difficultés rencontrées…")

rap = rp.generer_rapport(data, type_rap, jour, produit, obs, name)

st.subheader(f"{rap['titre']}")
st.caption(f"{rap['periode'].capitalize()}  ·  {rap['produit']}  ·  comparaison avec : {rap['periode_prec']}")

# ---------- Téléchargements (dans la langue choisie) ----------
rap_imp = rp.traduire_rapport(rap) if langue == "en" else rap
suffixe = "_EN" if langue == "en" else ""
nom_fichier = f"Rapport_{type_rap}_{produit or 'TousProduits'}_{rap['fin']:%Y-%m-%d}{suffixe}"
etiquette = " (English)" if langue == "en" else ""
b1, b2, _ = st.columns([1, 1, 2])
try:
    b1.download_button(f"🖨️ Télécharger le PDF (à imprimer){etiquette}", data=rp.construire_pdf(rap_imp),
                       file_name=nom_fichier + ".pdf", mime="application/pdf", type="primary", width="stretch")
except ImportError:
    b1.warning("Pour le PDF, installez : `pip install reportlab`")
b2.download_button(f"📥 Télécharger le tableau Excel{etiquette}", data=rp.construire_excel(rap_imp),
                   file_name=nom_fichier + ".xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")

# ---------- Synthèse ----------
st.markdown("#### 1. Synthèse de la période")
for debut in range(0, len(rap["kpis"]), 4):
    cols = st.columns(4)
    for col, k in zip(cols, rap["kpis"][debut:debut + 4]):
        col.metric(k["label"], rp._nb(k["cur"], k["dec"]), delta=None if k["evol"] == "=" else k["evol"],
                   help=f"{rap['lib_prec']} : {rp._nb(k['prev'], k['dec'])}")

# ==================== DÉTAILS ====================
t_evo, t_bons, t_pay, t_dech, t_app, t_act = st.tabs([
    "📈 Évolution & situation", "📑 Bons de commande", "💳 Bons de paiement", "🚛 Déchargements",
    "✅ Paiements approuvés", "✉️ Courrier"])

with t_evo:
    ev = rap["evolution"]
    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**Poids net déchargé (tonnes)**")
        st.bar_chart(ev.set_index("Date")["Poids net (T)"])
    with g2:
        st.markdown("**Montant des bons de paiement (FCFA)**")
        st.bar_chart(ev.set_index("Date")["Montant bons de paiement (FCFA)"])
    st.markdown("**Jour par jour** (" + ("la semaine" if type_rap == "Hebdomadaire" else "les 7 derniers jours") + ")")
    montrer(ev, {**CFG_DATE, "Montant livraisons (FCFA)": NUM("Montant livraisons (FCFA)"),
                 "Montant bons de paiement (FCFA)": NUM("Montant bons de paiement (FCFA)"),
                 "Paiements approuvés (FCFA)": NUM("Paiements approuvés (FCFA)")})
    if not rap["par_produit"].empty:
        st.markdown("**Répartition par produit**")
        montrer(rap["par_produit"], {"Montant livraisons (FCFA)": NUM("Montant livraisons (FCFA)"),
                                     "Montant bons de paiement (FCFA)": NUM("Montant bons de paiement (FCFA)")})
    st.markdown("**Situation à la fin de la période**")
    st.dataframe(pd.DataFrame(rap["etat"], columns=["Situation", "Valeur"]), hide_index=True, width="stretch")

with t_bons:
    st.markdown("**2. Bons de commande émis**")
    montrer(rap["bons"], CFG_DATE)

with t_pay:
    st.markdown("**3. Bons de paiement émis**")
    montrer(rap["paiements"], {**CFG_DATE, "Montant (FCFA)": NUM("Montant (FCFA)")})

with t_dech:
    st.markdown("**4. Déchargements**")
    montrer(rap["dech"], {**CFG_DATE, "Poids net (kg)": NUM("Poids net (kg)"), "Sacs": NUM("Sacs"), "PU": NUM("PU"),
                          "Montant dû (FCFA)": NUM("Montant dû (FCFA)")})
    if not rap["par_frs"].empty:
        st.markdown("**Par fournisseur**")
        montrer(rap["par_frs"], {"Poids net (kg)": NUM("Poids net (kg)"), "Montant dû (FCFA)": NUM("Montant dû (FCFA)")})

with t_app:
    st.markdown("**5. Paiements approuvés sur les livraisons**")
    montrer(rap["appro"], {**CFG_DATE, "Montant (FCFA)": NUM("Montant (FCFA)")})

with t_act:
    st.markdown("**6. Courrier (arrivée et départ) de la période**")
    st.caption("📨 Ces courriers viennent automatiquement de la page **Courrier** : c'est là que vous les saisissez "
               "(arrivée et départ). Le rapport les compte d'après leur date.")
    montrer(rap["act"], CFG_DATE, vide="Aucun courrier enregistré sur la période.")
    try:
        st.page_link("pages/3_Courrier.py", label="Ouvrir la page Courrier", icon="📨")
    except Exception:
        st.info("Pour saisir un courrier, utilisez la page **Courrier** du menu.")

st.markdown("---")
st.caption(f"🎨 Charte bleu ciel · module rapport v{rp.VERSION_RAPPORT}")
st.caption("NAP SARL — Rapports · sources : bons de commande, journal des paiements, déchargements, page Courrier "
           "(`courrier.xlsx`)")
