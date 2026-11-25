import streamlit as st
import pandas as pd
from datetime import datetime
import os
import re
import subprocess
import sys
import glob
import importlib

import dechargement_data as dd
import fournisseurs_data as fd
import telechargement_bons as tb
import suivi_bons as sb
importlib.reload(dd)   # relit les fichiers à chaque affichage (évite l'ancienne version en mémoire)
importlib.reload(fd)
importlib.reload(tb)
importlib.reload(sb)

st.set_page_config(page_title="Bons de Commande", page_icon="📄", layout="wide")

if (getattr(dd, "VERSION_DATA", 0) < 9 or getattr(fd, "VERSION_FRS", 0) < 2
        or getattr(sb, "VERSION_SUIVI", 0) < 1):
    st.error("⚠️ Les fichiers **dechargement_data.py**, **fournisseurs_data.py** et **suivi_bons.py** doivent être à "
             "jour, à côté de `app.py`. Remplacez-les par la dernière version puis arrêtez et relancez Streamlit.")
    st.stop()

# Vérifier authentification
if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name")
username = st.session_state.get("username")

# Première utilisation : la liste des fournisseurs est créée à partir des bons et déchargements existants
_init = fd.initialiser_si_absent(name)
if _init and _init["ajoutes"]:
    st.session_state["_flash_bc"] = (f"👥 Liste des fournisseurs créée automatiquement : {_init['ajoutes']} fournisseur(s) "
                                     f"trouvé(s) dans vos bons et déchargements.")

# ==================== CONFIGURATION ====================
FICHIER_EXCEL = "donnees_bons.xlsx"
DOSSIER_SORTIE = "bons_generes"

PRODUITS_SUFFIXES = {
    "Shea Nut (Noir de Karité)": "-BC-SHEA/NAP-2026",
    "Soja (Graines de Soja)": "-BC-SOYABEAN/NAP-2026",
    "Cotton Seed (Graines de Coton)": "-BC-COTTON/NAP-2026",
}
# =======================================================

# ---------- Fonctions utilitaires ----------
def extraire_numero(reference):
    if not reference:
        return 0
    base = os.path.basename(str(reference)).split("-")[0]
    match = re.search(r'(\d+)', base)
    if match:
        return int(match.group(1))
    match_fallback = re.search(r'(\d+)', str(reference))
    return int(match_fallback.group(1)) if match_fallback else 0

def charger_donnees():
    colonnes_ordonnees = [
        "Reference", "Date_Emission", "Heure_Creation", "Utilisateur", "Produit",
        "Code_Fournisseur", "Nom_Fournisseur", "Telephone_Fournisseur", "Commune_Origine",
        "Quantite", "Numero_Tracteur", "Chauffeur",
        "Numero_Permis", "Telephone_Chauffeur"
    ]
    if os.path.exists(FICHIER_EXCEL):
        df = pd.read_excel(FICHIER_EXCEL)
        if "Heure_Creation" not in df.columns:
            df.insert(2, "Heure_Creation", "")
        if "Utilisateur" not in df.columns:
            df.insert(3, "Utilisateur", "")
        if "Produit" not in df.columns:
            df.insert(4, "Produit", "")
        if "Code_Fournisseur" not in df.columns:
            df["Code_Fournisseur"] = ""
        cols = [c for c in colonnes_ordonnees if c in df.columns] + [c for c in df.columns if c not in colonnes_ordonnees]
        return df[cols]
    else:
        df = pd.DataFrame(columns=colonnes_ordonnees)
        df.to_excel(FICHIER_EXCEL, index=False)
        return df

def obtenir_prochain_numero(df):
    dernier = 0
    if not df.empty and "Reference" in df.columns:
        for ref in df["Reference"].dropna():
            num = extraire_numero(ref)
            if num > dernier:
                dernier = num
    return dernier + 1 if dernier > 0 else 1

def lister_fichiers_bons():
    """Liste tous les bons générés (Word ET PDF : SHEA / SOYABEAN / COTTON) dans bons_generes/ et ses sous-dossiers."""
    os.makedirs(DOSSIER_SORTIE, exist_ok=True)
    fichiers_valides = []
    for ext in ("docx", "pdf"):
        for f in glob.glob(os.path.join(DOSSIER_SORTIE, "**", f"*.{ext}"), recursive=True):
            nom = os.path.basename(f)
            if any(x in nom.lower() for x in ["template", "prototype", "modele", "model"]):
                continue
            if re.match(r'^\d{3,}-BC-(SHEA|SOYABEAN|COTTON)-NAP-2026', nom):
                fichiers_valides.append(f)
    return sorted(set(fichiers_valides), key=lambda p: (os.path.basename(p), p.count(os.sep)), reverse=True)


def _bouton_telechargement(col, chemin, label, mime, cle):
    """Bouton de téléchargement direct d'un fichier (Word ou PDF)."""
    with col:
        if not chemin:
            st.caption("—")
            return
        try:
            with open(chemin, "rb") as f:
                donnees = f.read()
            st.download_button(label=label, data=donnees, file_name=os.path.basename(chemin),
                               mime=mime, key=cle, width="stretch")
        except Exception as e:
            st.error(f"Erreur : {e}")

def supprimer_reference(reference_a_supprimer):
    df = charger_donnees()
    if reference_a_supprimer not in df["Reference"].astype(str).values:
        return False, "Cette référence n'existe pas."
    df = df[df["Reference"].astype(str) != str(reference_a_supprimer)]
    df.to_excel(FICHIER_EXCEL, index=False)
    return True, f"La référence {reference_a_supprimer} a été supprimée avec succès."


def modifier_bon(reference, champs):
    """Met à jour les champs d'un bon existant (référence inchangée)."""
    df = charger_donnees()
    mask = df["Reference"].astype(str) == str(reference)
    if not mask.any():
        return False, "Cette référence n'existe pas."
    idx = df.index[mask][0]
    autorises = [
        "Date_Emission", "Produit", "Code_Fournisseur", "Nom_Fournisseur",
        "Telephone_Fournisseur", "Commune_Origine", "Quantite",
        "Numero_Tracteur", "Chauffeur", "Numero_Permis", "Telephone_Chauffeur",
    ]
    for k, v in (champs or {}).items():
        if k in autorises and k in df.columns:
            df.at[idx, k] = v if v is not None else ""
    df.to_excel(FICHIER_EXCEL, index=False)
    return True, f"Bon {reference} modifié avec succès."

# ---------- Sidebar ----------
st.sidebar.header("➕ Nouveau Bon de Commande")

df_actuel = charger_donnees()
prochain_num = obtenir_prochain_numero(df_actuel)
prochain_num_str = f"{prochain_num:04d}"

# Produit HORS formulaire → suffixe mis à jour immédiatement
produit_choisi = st.sidebar.radio(
    "Produit *",
    options=list(PRODUITS_SUFFIXES.keys()),
    index=1,
    help="SHEA → -BC-SHEA/NAP-2026 | SOJA → -BC-SOYABEAN/NAP-2026 | COTON → -BC-COTTON/NAP-2026"
)
suffixe_auto = PRODUITS_SUFFIXES[produit_choisi]
st.sidebar.info(f"Référence → `XXXX{suffixe_auto}`")

# ---------- Fournisseur : choisi dans la liste, ou saisi s'il est nouveau ----------
if st.session_state.pop("_reset_frs", False):
    for _k in ("bc_frs_choix", "bc_frs_nom"):
        st.session_state.pop(_k, None)

if st.sidebar.button("👥 Voir / masquer la liste des fournisseurs", width="stretch"):
    st.session_state["voir_frs"] = not st.session_state.get("voir_frs", False)

NOUVEAU = "➕ Nouveau fournisseur…"
options_frs = fd.options_selection()
choix_frs = st.sidebar.selectbox("Fournisseur *", [""] + list(options_frs.keys()) + [NOUVEAU], key="bc_frs_choix",
                                 format_func=lambda x: "— Choisir dans la liste —" if x == "" else x)
nom_nouveau, tel_defaut, commune_defaut, cle_frs = "", "", "", "aucun"
if choix_frs == NOUVEAU:
    cle_frs = "nouveau"
    nom_nouveau = st.sidebar.text_input("Nom du nouveau fournisseur *", key="bc_frs_nom")
    st.sidebar.caption("➕ Il sera ajouté automatiquement à la liste des fournisseurs (avec un code FRS) "
                       "quand vous enregistrerez le bon.")
    if nom_nouveau.strip():
        exact = fd.chercher_exact(nom_nouveau)
        if exact:
            st.sidebar.info(f"Ce fournisseur existe déjà ({exact[0]}) : il sera utilisé, sans doublon.")
        else:
            proches = fd.fournisseurs_proches(nom_nouveau)
            if proches:
                st.sidebar.warning("Noms proches déjà dans la liste : " + " ; ".join(f"{c} — {n}" for c, n in proches)
                                   + ". Si c'est le même, choisissez-le dans la liste.")
elif choix_frs:
    cle_frs, _nom_sel = options_frs[choix_frs]
    _liste = fd.charger_fournisseurs()
    _fiche = _liste[_liste["Code"] == cle_frs]
    if not _fiche.empty:
        tel_defaut, commune_defaut = _fiche.iloc[0]["Telephone"], _fiche.iloc[0]["Commune"]

with st.sidebar.form("form_nouveau_bon", clear_on_submit=True):
    st.markdown("**Référence** *(suffixe selon le produit)*")
    col_num, col_suf = st.columns([1, 2])
    with col_num:
        numero_evolution = st.text_input("N° Évolution *", value=prochain_num_str, placeholder="0297")
    with col_suf:
        st.text_input("Suffixe (auto)", value=suffixe_auto, disabled=True)

    utilisateur = st.text_input("Utilisateur (Créé par) *", value=name)
    date_emission = st.date_input("Date d'émission", value=datetime.now())
    tel_fourn = st.text_input("Téléphone Fournisseur", value=tel_defaut, key=f"tel_{cle_frs}")
    commune = st.text_input("Commune d'origine", value=commune_defaut, key=f"com_{cle_frs}")
    quantite = st.text_input("Quantité (ex: 50 MT)")
    tracteur = st.text_input("N° Tracteur / Remorque")
    chauffeur = st.text_input("Nom du Chauffeur")
    permis = st.text_input("N° Permis")
    tel_chauffeur = st.text_input("Téléphone Chauffeur")

    submitted = st.form_submit_button("💾 Enregistrer dans Excel", width="stretch")

    if submitted:
        num_clean = extraire_numero(numero_evolution)
        if num_clean == 0:
            st.error("Le numéro d'évolution est obligatoire et doit être numérique !")
        elif not utilisateur.strip():
            st.error("Le nom de l'utilisateur est obligatoire !")
        elif not choix_frs:
            st.error("Choisissez un fournisseur dans la liste (ou « Nouveau fournisseur… ») !")
        elif choix_frs == NOUVEAU and not nom_nouveau.strip():
            st.error("Saisissez le nom du nouveau fournisseur !")
        else:
            reference = f"{num_clean:04d}{suffixe_auto}"
            df = charger_donnees()
            if reference in df["Reference"].astype(str).values:
                st.error(f"La référence {reference} existe déjà !")
            else:
                # Fournisseur : créé automatiquement dans la liste s'il n'y est pas encore
                if choix_frs == NOUVEAU:
                    code_frs, nom_frs, cree_frs = fd.assurer_fournisseur(nom_nouveau, tel_fourn, commune, name)
                else:
                    (code_frs, nom_frs), cree_frs = options_frs[choix_frs], False
                heure_creation = datetime.now().strftime("%H:%M:%S")
                nouvelle_ligne = {
                    "Reference": reference,
                    "Date_Emission": date_emission.strftime("%d/%m/%Y"),
                    "Heure_Creation": heure_creation,
                    "Utilisateur": utilisateur.strip(),
                    "Produit": produit_choisi,
                    "Code_Fournisseur": code_frs,
                    "Nom_Fournisseur": nom_frs,
                    "Telephone_Fournisseur": tel_fourn.strip(),
                    "Commune_Origine": commune.strip(),
                    "Quantite": quantite.strip(),
                    "Numero_Tracteur": tracteur.strip(),
                    "Chauffeur": chauffeur.strip(),
                    "Numero_Permis": permis.strip(),
                    "Telephone_Chauffeur": tel_chauffeur.strip()
                }
                df = pd.concat([df, pd.DataFrame([nouvelle_ligne])], ignore_index=True)
                df.to_excel(FICHIER_EXCEL, index=False)
                msg = f"✅ Bon {reference} enregistré — {code_frs} {nom_frs}"
                if cree_frs:
                    msg += f"  ·  ➕ nouveau fournisseur ajouté à la liste ({code_frs})"
                st.session_state["_flash_bc"] = msg
                st.session_state["_reset_frs"] = True
                st.rerun()

# ---------- Zone principale ----------
st.title("📄 Gestion des Bons de Commande – NAP SARL")
if "_flash_bc" in st.session_state:
    st.success(st.session_state.pop("_flash_bc"))

if st.session_state.get("voir_frs"):
    with st.expander("👥 Liste des fournisseurs", expanded=True):
        _lst = fd.charger_fournisseurs()
        _q = st.text_input("Rechercher un fournisseur", key="q_frs_bc")
        if _q.strip() and not _lst.empty:
            _lst = _lst[_lst.apply(lambda r: _q.strip().lower() in " ".join(r.astype(str)).lower(), axis=1)]
        st.dataframe(_lst[["Code", "Nom", "Telephone", "Commune", "Actif"]], hide_index=True, width="stretch",
                     height=300)
        st.caption(f"{len(_lst)} fournisseur(s) · pour modifier un fournisseur : page « Fournisseurs ».")

# ==================== MINI TABLEAU DE BORD + RECHERCHE D'UN CAMION ====================
st.markdown("---")
st.subheader("📈 Suivi : bons émis, déchargés et reste à décharger")
_prod_dash = st.radio("Produit", [None] + dd.PRODUITS, horizontal=True, key="dash_produit",
                      format_func=lambda p: "Tous les produits" if p is None else dd.libelle_produit(p))
_bons_all = sb.charger_bons_tous()
_dech_all = sb.charger_dechargements_tous()
_suivi_tous = sb.suivi_par_bon(_bons_all, _dech_all)                 # tous produits (pour la recherche)
_suivi = _suivi_tous if _prod_dash is None else _suivi_tous[_suivi_tous["Produit"] == _prod_dash]
_ind = sb.indicateurs(_suivi)
_ICONES_STATUT = {"LIVRE": "🟢 LIVRÉ", "EN COURS": "🟠 EN COURS", "EN ATTENTE": "🔴 EN ATTENTE"}


def _t1(x):
    return f"{x:,.1f}".replace(",", " ")


k1, k2, k3, k4 = st.columns(4)
k1.metric("Bons de commande émis", _ind["emis"],
          help=f"{_ind['emis_aujourdhui']} émis aujourd'hui")
k2.metric("Bons déchargés", _ind["decharges"],
          delta=f"{_ind['camions']} camion(s) · {_ind['livres']} bon(s) complet(s)", delta_color="off",
          help="Bons pour lesquels au moins un camion a été déchargé")
k3.metric("Reste à décharger", _ind["a_decharger"], delta=f"{_t1(_ind['reste_T'])} T en attente",
          delta_color="inverse" if _ind["a_decharger"] else "off",
          help="Bons dont la livraison n'est pas terminée (EN ATTENTE ou EN COURS)")
k4.metric("Déchargé / commandé", f"{_t1(_ind['livre_T'])} / {_t1(_ind['qte_T'])} T")
if _ind["qte_T"] > 0:
    st.progress(min(_ind["livre_T"] / _ind["qte_T"], 1.0),
                text=f"{min(_ind['livre_T'] / _ind['qte_T'] * 100, 100):.0f} % de la quantité commandée est déchargée")

if _ind["emis"] == 0:
    st.info("Aucun bon de commande pour ce produit.")
elif _ind["a_decharger"] == 0:
    st.success("✅ Tous les bons de commande sont déchargés.")
else:
    _message = (f"⚠️ **{_ind['a_decharger']} bon(s) de commande restent à décharger** — "
                f"{_t1(_ind['reste_T'])} T en attente.")
    if _ind["en_retard"]:
        st.error(_message + f"  🔴 **{_ind['en_retard']} en retard** (émis depuis plus de {sb.JOURS_ALERTE} jours).")
    else:
        st.warning(_message)
    _liste = sb.bons_a_decharger(_suivi)
    with st.expander(f"📋 Bons restant à décharger ({len(_liste)})", expanded=bool(_ind["en_retard"])):
        _vue = _liste.assign(Statut=_liste["Statut"].map(_ICONES_STATUT))[
            ["Alerte", "Reference", "Produit", "Fournisseur", "Date", "Age_Jours", "Camion_Bon", "Qte_T", "Livre_T",
             "Reste_T", "Statut"]]
        st.dataframe(_vue, hide_index=True, width="stretch", column_config={
            "Reference": "N° Bon de commande", "Date": st.column_config.DateColumn("Émis le", format="DD/MM/YYYY"),
            "Age_Jours": st.column_config.NumberColumn("Âge (jours)", format="%d"), "Camion_Bon": "Camion prévu",
            "Qte_T": st.column_config.NumberColumn("Commandé (T)", format="%.1f"),
            "Livre_T": st.column_config.NumberColumn("Déchargé (T)", format="%.1f"),
            "Reste_T": st.column_config.NumberColumn("Reste (T)", format="%.1f")})

st.markdown("#### 🔎 Vérifier un camion ou un bon")
_q = st.text_input("Numéro de camion (ou numéro de bon)", key="q_camion_bc",
                   placeholder="Ex. BE5458RB   ou   0013   (espaces et tirets sans importance)")
if _q.strip():
    _res = sb.rechercher(_q, _bons_all, _dech_all)
    if _res is None:
        st.info("Saisissez au moins 2 caractères.")
    else:
        _rb, _rd = _res["bons"], _res["dech"]
        if not _rb.empty:
            _refs = set(_rb["Reference"])
            _deja = _suivi_tous[_suivi_tous["Reference"].isin(_refs) & (_suivi_tous["Nb_Camions"] > 0)]
            _txt = (f"✅ **Provient de nos bons de commande** : {len(_rb)} bon(s) correspondent à « {_q.strip()} » "
                    f"({', '.join(_rb['Reference'].str.split('-').str[0])}).")
            if not _deja.empty:
                _txt += f"  Déjà déchargé pour {len(_deja)} bon(s)."
            st.success(_txt)
        elif not _rd.empty:
            _refs_dech = sorted({r for r in _rd["Ref_BC"] if r})
            if _refs_dech:
                st.warning(f"⚠️ « {_q.strip()} » a été déchargé avec le n° de bon **{', '.join(_refs_dech)}**, mais ce bon "
                           "n'existe pas dans la liste actuelle des bons de commande (ancien numéro ?).")
            else:
                st.warning(f"⚠️ « {_q.strip()} » a été **déchargé sans aucun numéro de bon de commande**.")
        else:
            st.error(f"❌ **Aucun bon de commande ni déchargement** ne correspond à « {_q.strip()} » : "
                     "ce camion / ce bon ne provient pas de nos bons de commande.")
        if not _rb.empty:
            _s = _suivi_tous[_suivi_tous["Reference"].isin(set(_rb["Reference"]))]
            st.markdown("**Bon(s) de commande trouvé(s)**")
            st.dataframe(_s.assign(Statut=_s["Statut"].map(_ICONES_STATUT))[
                ["Reference", "Produit", "Fournisseur", "Date", "Camion_Bon", "Chauffeur", "Quantite", "Statut",
                 "Livre_T", "Reste_T"]], hide_index=True, width="stretch", column_config={
                "Reference": "N° Bon de commande", "Date": st.column_config.DateColumn("Émis le", format="DD/MM/YYYY"),
                "Camion_Bon": "Camion prévu", "Quantite": "Quantité",
                "Livre_T": st.column_config.NumberColumn("Déchargé (T)", format="%.1f"),
                "Reste_T": st.column_config.NumberColumn("Reste (T)", format="%.1f")})
        if not _rd.empty:
            st.markdown("**Déchargement(s) correspondant(s)**")
            _d = _rd.assign(Alerte=["⚠️ Camion différent du bon" if dd.camion_different(a, b) else ""
                                    for a, b in zip(_rd["Camion_BC"], _rd["N_Camion"])])
            st.dataframe(_d[["Date_Dechargement", "Produit", "Transaction", "Ref_BC", "Fournisseur", "N_Camion",
                             "Camion_BC", "Poids_Net_KG", "Alerte"]], hide_index=True, width="stretch",
                         column_config={"Date_Dechargement": st.column_config.DateColumn("Déchargé le", format="DD/MM/YYYY"),
                                        "Ref_BC": "N° Bon de commande", "N_Camion": "Camion livré",
                                        "Camion_BC": "Camion prévu sur le bon",
                                        "Poids_Net_KG": st.column_config.NumberColumn("Poids net (kg)", format="%d")})

st.markdown("---")

st.subheader("📊 Tableau des Bons de Commande")
df = charger_donnees()

if df.empty:
    st.info("Aucun bon de commande enregistré pour le moment.")
else:
    st.dataframe(df, width="stretch", height=400)
    st.markdown("---")

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("🔄 Actualiser le tableau", width="stretch"):
            st.rerun()
    with col2:
        regenerer_tout = st.checkbox(
            "Régénérer aussi les anciens bons (nouveau format)", value=False, key="regen_bc",
            help="Recrée tous les bons en Word et PDF, y compris ceux déjà générés.")
        signature_digitale = st.checkbox(
            "🖊️ Signature numérique Direction (cachet)", value=False, key="sig_bc",
            help="Cochez pour coller la signature + cachet STE NAP SARL dans la zone « POUR L'ACHETEUR ». "
                 "Décochez pour laisser l'espace vide (signature physique).")
        if st.button("📄 Générer les nouveaux Bons (Word + PDF)", width="stretch", type="primary"):
            with st.spinner("Génération en cours..."):
                try:
                    commande = [sys.executable, "generer_bons.py"]
                    if regenerer_tout:
                        commande.append("--tout")
                    if signature_digitale:
                        commande.append("--signature")
                    result = subprocess.run(
                        commande, capture_output=True, text=True, encoding="utf-8", errors="replace",
                        cwd=os.getcwd()
                    )
                    if result.returncode == 0:
                        st.success("Bons générés en Word et PDF avec succès ! Ils sont dans la liste de téléchargement ci-dessous.")
                        st.code(result.stdout)
                    else:
                        st.error("Erreur lors de la génération")
                        st.code((result.stdout or "") + "\n" + (result.stderr or ""))
                except Exception as e:
                    st.error(f"Erreur : {e}")
    with col3:
        with open(FICHIER_EXCEL, "rb") as f:
            st.download_button(
                label="⬇️ Télécharger Excel",
                data=f,
                file_name="donnees_bons.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )

# ---------- MODIFICATION D'UN BON ----------
st.markdown("---")
st.subheader("✏️ Modifier un Bon de Commande")
df_mod = charger_donnees()
if df_mod.empty:
    st.info("Aucun bon à modifier.")
else:
    refs_mod = df_mod["Reference"].astype(str).tolist()[::-1]  # plus récents en premier
    ref_edit = st.selectbox(
        "Choisissez le bon à modifier",
        options=refs_mod,
        key="select_mod_bc",
        help="La référence (N°) ne change pas. Modifiez les autres champs puis enregistrez.",
    )
    row = df_mod[df_mod["Reference"].astype(str) == str(ref_edit)].iloc[0]

    def _s(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        return str(v).strip()

    # Date
    date_val = row.get("Date_Emission")
    try:
        if hasattr(date_val, "date"):
            date_def = date_val.date() if hasattr(date_val, "date") else date_val
        else:
            date_def = pd.to_datetime(date_val, dayfirst=True, errors="coerce")
            date_def = date_def.date() if pd.notna(date_def) else datetime.now().date()
    except Exception:
        date_def = datetime.now().date()

    # Produit : retrouver la clé radio si possible
    prod_opts = list(PRODUITS_SUFFIXES.keys())
    prod_cur = _s(row.get("Produit"))
    idx_prod = 0
    for i, p in enumerate(prod_opts):
        if prod_cur and (prod_cur == p or prod_cur.lower() in p.lower() or p.lower() in prod_cur.lower()):
            idx_prod = i
            break
        if "SHEA" in str(ref_edit).upper() and "Shea" in p:
            idx_prod = i
            break
        if "SOYABEAN" in str(ref_edit).upper() and "Soja" in p:
            idx_prod = i
            break
        if "COTTON" in str(ref_edit).upper() and "Cotton" in p:
            idx_prod = i
            break

    with st.form("form_modifier_bon"):
        st.caption(f"Référence **{ref_edit}** (non modifiable) · Créé par : {_s(row.get('Utilisateur'))} · {_s(row.get('Heure_Creation'))}")
        c1, c2, c3 = st.columns(3)
        with c1:
            m_date = st.date_input("Date d'émission", value=date_def)
            m_produit = st.selectbox("Produit", prod_opts, index=idx_prod)
            m_nom = st.text_input("Nom fournisseur *", value=_s(row.get("Nom_Fournisseur")))
            m_code = st.text_input("Code fournisseur", value=_s(row.get("Code_Fournisseur")))
        with c2:
            m_tel = st.text_input("Téléphone fournisseur", value=_s(row.get("Telephone_Fournisseur")))
            m_commune = st.text_input("Commune d'origine", value=_s(row.get("Commune_Origine")))
            m_qte = st.text_input("Quantité", value=_s(row.get("Quantite")))
            m_tracteur = st.text_input("N° Tracteur / Remorque", value=_s(row.get("Numero_Tracteur")))
        with c3:
            m_chauffeur = st.text_input("Chauffeur", value=_s(row.get("Chauffeur")))
            m_permis = st.text_input("N° Permis", value=_s(row.get("Numero_Permis")))
            m_tel_ch = st.text_input("Téléphone chauffeur", value=_s(row.get("Telephone_Chauffeur")))
        st.info(
            "Après modification, régénérez le bon si besoin : cochez « Régénérer aussi les anciens bons » "
            "puis cliquez sur « Générer les nouveaux Bons (Word + PDF) »."
        )
        btn_mod = st.form_submit_button("💾 Enregistrer les modifications", type="primary", width="stretch")
        if btn_mod:
            if not m_nom.strip():
                st.error("Le nom du fournisseur est obligatoire.")
            else:
                ok, msg = modifier_bon(ref_edit, {
                    "Date_Emission": m_date.strftime("%d/%m/%Y"),
                    "Produit": m_produit,
                    "Code_Fournisseur": m_code.strip(),
                    "Nom_Fournisseur": m_nom.strip(),
                    "Telephone_Fournisseur": m_tel.strip(),
                    "Commune_Origine": m_commune.strip(),
                    "Quantite": m_qte.strip(),
                    "Numero_Tracteur": m_tracteur.strip(),
                    "Chauffeur": m_chauffeur.strip(),
                    "Numero_Permis": m_permis.strip(),
                    "Telephone_Chauffeur": m_tel_ch.strip(),
                })
                if ok:
                    st.session_state["_flash_bc"] = "✅ " + msg
                    st.rerun()
                else:
                    st.error(msg)

# ---------- SUPPRESSION (admin) ----------
if username == "admin":
    st.markdown("---")
    st.subheader("🗑️ Supprimer un Bon de Commande (Admin uniquement)")
    df = charger_donnees()
    if not df.empty:
        liste_references = df["Reference"].astype(str).tolist()
        reference_a_supprimer = st.selectbox("Choisissez la référence à supprimer :", options=liste_references, key="select_suppression")
        col_del1, col_del2 = st.columns([1, 3])
        with col_del1:
            if st.button("🗑️ Supprimer cette référence", type="primary", width="stretch"):
                success, message = supprimer_reference(reference_a_supprimer)
                if success:
                    st.success(message)
                    st.rerun()
                else:
                    st.error(message)
        with col_del2:
            st.warning("⚠️ Cette action est irréversible.")
    else:
        st.info("Aucun bon à supprimer.")

# ---------- Téléchargement Word et PDF ----------
st.markdown("---")
st.subheader("📥 Télécharger les Bons de Commande (Word et PDF)")

if st.button("🔄 Actualiser la liste des fichiers"):
    st.rerun()

tb.afficher_moteur()
tb.afficher_erreur_pdf()
tous_fichiers = lister_fichiers_bons()

if not tous_fichiers:
    st.warning("Aucun bon trouvé dans le dossier `bons_generes/`.")
    st.info("1. Cliquez d'abord sur « Générer les nouveaux Bons (Word + PDF) »\n2. Puis actualisez la liste")
else:
    MIME_WORD = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    groupes = {}
    for fichier in sorted(tous_fichiers, key=lambda p: p.count(os.sep)):
        base, ext = os.path.splitext(os.path.basename(fichier))
        groupes.setdefault(base, {}).setdefault(ext.lower(), fichier)   # garde le fichier le plus proche du dossier
    # Anciens bons (Word seul) : on garde l'ancien système de conversion PDF
    anciens_word = [g[".docx"] for g in groupes.values() if ".docx" in g and ".pdf" not in g]
    st.success(f"**{len(groupes)} bon(s) disponible(s) :**")
    if anciens_word:
        tb.afficher_tout_creer(anciens_word, "bc")
    for base in sorted(groupes, reverse=True):
        mots = groupes[base]
        col1, col2, col3 = st.columns([5, 1, 1])
        with col1:
            st.write(f"📄 **{base}**")
        if ".docx" in mots and ".pdf" not in mots:
            try:
                tb.afficher_boutons(mots[".docx"], f"bc_{mots['.docx']}", col2, col3)
            except Exception as e:
                st.error(f"Erreur : {e}")
        else:
            _bouton_telechargement(col2, mots.get(".docx"), "⬇️ Word", MIME_WORD, f"dlw_{base}")
            _bouton_telechargement(col3, mots.get(".pdf"), "⬇️ PDF", "application/pdf", f"dlp_{base}")

st.markdown("---")
st.caption("NAP SARL – Bons multi-produits SHEA / SOYABEAN / COTTON")