import streamlit as st
import pandas as pd
import importlib

import dechargement_data as dd
import fournisseurs_data as fd
importlib.reload(dd)   # relit les fichiers à chaque affichage (évite l'ancienne version en mémoire)
importlib.reload(fd)

st.set_page_config(page_title="Fournisseurs", page_icon="👥", layout="wide")

if getattr(dd, "VERSION_DATA", 0) < 9 or getattr(fd, "VERSION_FRS", 0) < 2:
    st.error("⚠️ Les fichiers **dechargement_data.py** et **fournisseurs_data.py** doivent être à jour, à côté de "
             "`app.py`. Remplacez-les par la dernière version puis arrêtez et relancez Streamlit.")
    st.stop()

# ---------- Authentification (comme les autres pages) ----------
if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"


def flash(type_, message):
    st.session_state["_flash"] = (type_, message)


# Première utilisation : création de la liste à partir des bons et déchargements existants
_init = fd.initialiser_si_absent(name)
if _init and _init["ajoutes"]:
    st.session_state["_flash"] = ("success", f"👥 Liste créée automatiquement : {_init['ajoutes']} fournisseur(s) trouvé(s) "
                                  f"dans vos bons de commande et déchargements ({_init['codes_bons']} bon(s) rattaché(s)).")
    st.session_state["_a_verifier"] = _init["a_verifier"]

if st.session_state.pop("_reset_add", False):
    for _k in ("add_nom", "add_tel", "add_commune", "add_confirm"):
        st.session_state.pop(_k, None)

st.title("👥 Fournisseurs")
if "_flash" in st.session_state:
    type_, message = st.session_state.pop("_flash")
    getattr(st, type_)(message)

liste = fd.charger_fournisseurs()
tab_liste, tab_add, tab_maj = st.tabs(["📋 Liste des fournisseurs", "➕ Ajouter un fournisseur",
                                       "✏️ Mettre à jour les informations"])

# ---------------------------------------------------------------- LISTE
with tab_liste:
    m1, m2, m3 = st.columns(3)
    m1.metric("Fournisseurs", dd.fmt(len(liste)))
    m2.metric("Actifs", dd.fmt((liste["Actif"] == "OUI").sum()))
    m3.metric("Inactifs", dd.fmt((liste["Actif"] != "OUI").sum()))

    f1, f2, f3 = st.columns([2, 1, 1])
    q = f1.text_input("Rechercher (nom, code, téléphone, commune)", key="q_liste")
    statut = f2.selectbox("Afficher", ["Actifs", "Inactifs", "Tous"], key="statut_liste")
    avec_soldes = f3.checkbox("Avec camions et reste à payer", value=True, key="soldes_liste")

    vue = liste.copy()
    if statut == "Actifs":
        vue = vue[vue["Actif"] == "OUI"]
    elif statut == "Inactifs":
        vue = vue[vue["Actif"] != "OUI"]
    if q.strip() and not vue.empty:
        vue = vue[vue.apply(lambda r: q.strip().lower() in " ".join(r.astype(str)).lower(), axis=1)]
    cols = ["Code", "Nom", "Telephone", "Commune", "Actif"]
    if avec_soldes:
        vue = vue.merge(fd.soldes_par_fournisseur(), on="Code", how="left")
        for c in ("Camions", "Montant_Du", "Montant_Paye", "Reste_A_Payer"):
            vue[c] = vue[c].fillna(0)
        cols += ["Camions", "Montant_Du", "Montant_Paye", "Reste_A_Payer"]
    vue = vue[cols]
    if vue.empty:
        st.info("Aucun fournisseur ne correspond.")
    else:
        st.dataframe(vue, hide_index=True, width="stretch", height=480, column_config={
            "Telephone": "Téléphone", "Camions": st.column_config.NumberColumn("Camions", format="%d"),
            "Montant_Du": st.column_config.NumberColumn("Total dû (FCFA)", format="%d"),
            "Montant_Paye": st.column_config.NumberColumn("Total payé (FCFA)", format="%d"),
            "Reste_A_Payer": st.column_config.NumberColumn("Reste à payer (FCFA)", format="%d")})
    st.caption(f"{len(vue)} fournisseur(s) affiché(s)." + (" Camions et montants : tous produits confondus."
                                                           if avec_soldes else ""))
    st.download_button("⬇️ Télécharger la liste (Excel)", data=dd.to_excel_bytes({"Fournisseurs": vue}),
                       file_name="liste_fournisseurs.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.info("💡 Un fournisseur est aussi ajouté automatiquement ici quand vous créez un bon de commande "
            "pour un nom qui n'est pas encore dans la liste.")

    with st.expander("🔄 Compléter la liste avec les fournisseurs déjà présents dans les bons et déchargements"):
        st.caption("Ajoute les fournisseurs trouvés dans vos bons de commande et déchargements qui ne sont pas encore "
                   "dans la liste (sans doublon), et renseigne leur code sur les bons qui n'en ont pas.")
        if st.button("🔄 Compléter la liste maintenant", key="btn_completer"):
            res = fd.completer_depuis_donnees(name)
            st.session_state["_a_verifier"] = res["a_verifier"]
            flash("success" if res["ajoutes"] or res["codes_bons"] else "info",
                  f"{res['ajoutes']} fournisseur(s) ajouté(s) · {res['codes_bons']} bon(s) de commande rattaché(s).")
            st.rerun()
    a_verifier = st.session_state.get("_a_verifier") or []
    if a_verifier:
        with st.expander(f"⚠️ {len(a_verifier)} nom(s) proche(s) à vérifier (doublons possibles)", expanded=True):
            st.caption("Ces fournisseurs ont des noms voisins. Si c'est la même personne, gardez-en un seul "
                       "(désactivez l'autre dans « Mettre à jour les informations »). Sinon, ne faites rien.")
            st.dataframe(pd.DataFrame(a_verifier, columns=["Code", "Nom", "Code proche", "Nom proche"]),
                         hide_index=True, width="stretch")

# ---------------------------------------------------------------- AJOUTER
with tab_add:
    st.subheader("Ajouter un fournisseur")
    with st.form("form_ajout_frs"):
        nom_a = st.text_input("Nom du fournisseur *", key="add_nom")
        a1, a2 = st.columns(2)
        tel_a = a1.text_input("Téléphone", key="add_tel")
        com_a = a2.text_input("Commune", key="add_commune")
        confirme = st.checkbox("C'est bien un NOUVEAU fournisseur (même si des noms proches existent déjà)",
                               key="add_confirm")
        envoye = st.form_submit_button("💾 Ajouter à la liste", type="primary")
    if envoye:
        exact = fd.chercher_exact(nom_a)
        proches = fd.fournisseurs_proches(nom_a)
        if not nom_a.strip():
            st.error("Le nom du fournisseur est obligatoire.")
        elif exact:
            st.error(f"Ce fournisseur existe déjà : {exact[0]} — {exact[1]}.")
        elif proches and not confirme:
            st.warning("Des noms proches existent déjà : " + " ; ".join(f"{c} — {n}" for c, n in proches) +
                       ". Si c'est le même fournisseur, ne l'ajoutez pas ; sinon cochez la case de confirmation "
                       "puis réessayez.")
        else:
            ok, msg, _ = fd.creer_fournisseur(nom_a, tel_a, com_a, name)
            flash("success" if ok else "error", msg)
            st.session_state["_reset_add"] = True
            st.rerun()

# ---------------------------------------------------------------- METTRE À JOUR
with tab_maj:
    st.subheader("Mettre à jour les informations d'un fournisseur")
    if liste.empty:
        st.info("Aucun fournisseur à modifier.")
    else:
        opts = {f"{r.Code} — {r.Nom}" + ("  (inactif)" if r.Actif != "OUI" else ""): r.Code
                for r in liste.itertuples()}
        choix = st.selectbox("Fournisseur à mettre à jour", list(opts.keys()), key="maj_choix")
        code = opts[choix]
        r = liste[liste["Code"] == code].iloc[0]
        histo = []
        if r["Cree_Par"] or r["Date_Creation"]:
            histo.append(f"Créé par {r['Cree_Par'] or '—'} le {r['Date_Creation'] or '—'}")
        if r["Modifie_Par"] or r["Date_Modification"]:
            histo.append(f"dernière mise à jour par {r['Modifie_Par'] or '—'} le {r['Date_Modification'] or '—'}")
        if histo:
            st.caption(" · ".join(histo))

        with st.form(f"form_maj_{code}"):
            nom_m = st.text_input("Nom du fournisseur *", value=r["Nom"], key=f"maj_nom_{code}")
            b1, b2 = st.columns(2)
            tel_m = b1.text_input("Téléphone", value=r["Telephone"], key=f"maj_tel_{code}")
            com_m = b2.text_input("Commune", value=r["Commune"], key=f"maj_com_{code}")
            actif_m = st.checkbox("Fournisseur actif (proposé dans la création des bons de commande)",
                                  value=(r["Actif"] == "OUI"), key=f"maj_actif_{code}")
            propager = st.checkbox("Si le nom change : le répercuter sur les anciens bons de commande et "
                                   "déchargements de ce fournisseur", value=True, key=f"maj_prop_{code}")
            maj = st.form_submit_button("💾 Mettre à jour", type="primary")
        if maj:
            ok, msg = fd.modifier_fournisseur(code, nom=nom_m, telephone=tel_m, commune=com_m, actif=actif_m,
                                              utilisateur=name, propager_nom=propager)
            if ok:
                flash("success", msg)
                st.rerun()
            else:
                st.warning(msg)

st.markdown("---")
st.caption("NAP SARL — Fournisseurs · fichier `fournisseurs.xlsx` (colonnes existantes conservées)")
