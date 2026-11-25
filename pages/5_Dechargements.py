import streamlit as st
import pandas as pd
from datetime import date

import dechargement_data as dd
try:
    import clients_data as cd
except ImportError:
    cd = None
try:
    import proforma_data as pf
except ImportError:
    pf = None

st.set_page_config(page_title="Déchargements", page_icon="🚛", layout="wide")

# ---------- Authentification (comme les autres pages) ----------
if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"
username = st.session_state.get("username") or ""
peut_approuver = username in dd.ROLES_APPROBATION
est_admin = username == "admin"

# ---------- Produit actif (un fichier de déchargements par produit) ----------
if "produit_actif" not in st.session_state:
    st.session_state["produit_actif"] = dd.PRODUIT_DEFAUT
produit = st.sidebar.selectbox("📦 Produit", dd.PRODUITS, index=dd.PRODUITS.index(st.session_state["produit_actif"]),
                               format_func=dd.libelle_produit)
st.session_state["produit_actif"] = produit
dd.set_produit(produit)
st.sidebar.divider()

ICONES = {"PAYE": "🟢 PAYE", "PARTIEL": "🟡 PARTIEL", "NON PAYE": "🔴 NON PAYE", "SURPLUS": "🔵 SURPLUS", "—": "—"}


def flash(type_, message):
    """Message affiché après le rechargement de la page."""
    st.session_state["_flash"] = (type_, message)


def tableau_dechargements(d):
    """Colonnes affichées + icône de statut."""
    v = d.copy()
    v["Statut_Paiement"] = v["Statut_Paiement"].map(ICONES)
    v["Alerte_Camion"] = ["⚠️ Différent" if x else "" for x in v["Camion_Diff"]]
    # Colonnes optionnelles (affectation BC ↔ fournisseur ticket)
    for col, default in [("Type_Mouvement", "ACHAT"), ("Fournisseur_BC", ""), ("Affectation_BC", ""), ("Alerte_Affectation", "")]:
        if col not in v.columns:
            v[col] = default
    if "Alerte_Affectation" not in v.columns or v["Alerte_Affectation"].eq("").all():
        # recalcul affichage si colonne absente
        def _alerte_aff(r):
            fb = str(r.get("Fournisseur_BC") or "").strip()
            ft = str(r.get("Fournisseur") or "").strip()
            if fb and ft and fb.casefold() != ft.casefold():
                return "⚠️ BC d'un autre frs"
            return ""
        v["Alerte_Affectation"] = v.apply(_alerte_aff, axis=1)
    cols = ["Date_Dechargement", "Transaction", "Type_Mouvement", "Ref_BC", "Fournisseur", "Fournisseur_BC", "Alerte_Affectation",
            "N_Camion", "Camion_BC", "Alerte_Camion",
            "Weighment_Ticket", "Poids_Net_KG", "Nb_Sacs", "PU_FCFA_KG", "Montant_Du", "Montant_Paye", "Solde",
            "Statut_Paiement", "Bons_Paiement", "Lieu_Stockage", "Remarque"]
    cols = [c for c in cols if c in v.columns]
    return v[cols]


CONFIG_COLONNES = {
    "Date_Dechargement": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
    "Type_Mouvement": "Type",
    "Ref_BC": "N° BC / Proforma",
    "Fournisseur": "Fournisseur / Client",
    "Fournisseur_BC": "Fournisseur du BC",
    "Alerte_Affectation": "Échange BC",
    "N_Camion": "Camion livré",
    "Camion_BC": "Camion sur le bon",
    "Alerte_Camion": "Alerte camion",
    "Remarque": "Remarque",
    "Weighment_Ticket": "Ticket pesée",
    "Poids_Net_KG": st.column_config.NumberColumn("Poids net (kg)", format="%d"),
    "Nb_Sacs": st.column_config.NumberColumn("Sacs", format="%d"),
    "PU_FCFA_KG": st.column_config.NumberColumn("PU", format="%d"),
    "Montant_Du": st.column_config.NumberColumn("Montant dû", format="%d"),
    "Montant_Paye": st.column_config.NumberColumn("Payé", format="%d"),
    "Solde": st.column_config.NumberColumn("Solde", format="%d"),
    "Statut_Paiement": "Statut",
    "Bons_Paiement": "Bon(s) de paiement",
    "Lieu_Stockage": "Stockage",
}

# ---------- Données ----------
df, aff = dd.charger_complet()
bons = dd.charger_bons_commande()
frs = dict(dd.fournisseurs_connus())   # tous produits confondus

# ==================== SIDEBAR : nouveau déchargement ====================
st.sidebar.header("➕ Nouveau déchargement")

type_mvt = st.sidebar.radio(
    "Type de mouvement *",
    ["ACHAT", "VENTE"],
    horizontal=True,
    help="ACHAT = réception matière auprès d'un fournisseur. "
         "VENTE = sortie / déchargement lié à une vente export (client).",
    key="type_mvt_new",
)
st.sidebar.caption(
    "🛒 **ACHAT** → compte fournisseur · 🚢 **VENTE** → flux lié à l'export / client"
)

# Options clients (pour VENTE)
opts_cli = {}
if cd is not None:
    try:
        opts_cli = cd.options_selection()
    except Exception:
        opts_cli = {}
labels_cli = ["— Choisir un client —"] + list(opts_cli.keys())


# --- Lien document : BC (ACHAT) ou Proforma (VENTE) ---
sans_bc = False
confirmer_creation_bc = False
bc = None
ref_bc = ""
ref_proforma = ""

if type_mvt == "ACHAT":
    mode_bc = st.sidebar.radio(
        "Lien bon de commande *",
        [
            "Avec bon de commande existant",
            "Sans bon — créer automatiquement à l'enregistrement",
        ],
        help="Si le camion arrive sans BC déjà émis, choisissez « Sans bon ».",
    )
    sans_bc = mode_bc.startswith("Sans bon")
    SANS_BC = "— Sans bon de commande —"
    deja_decharges = set(df["Ref_BC"])
    options_bc = {SANS_BC: None}
    for _, b in bons.sort_values("Reference", ascending=False).iterrows():
        marque = "  ✔ déjà déchargé" if b["Reference"] in deja_decharges else ""
        options_bc[f"{b['Reference']} · {b['Nom_Fournisseur']} · {b['Quantite']}{marque}"] = b
    if sans_bc:
        st.sidebar.warning(
            "**Mode sans BC** : création auto d'un bon "
            f"(ex. `XXXX-BC-{dd.get_produit()}/NAP-{date.today().year}`)."
        )
        confirmer_creation_bc = st.sidebar.checkbox(
            "Je confirme : créer un nouveau bon de commande",
            value=False,
            key="confirm_creer_bc",
        )
    else:
        choix_bc = st.sidebar.selectbox(
            "Bon de commande *",
            list(options_bc.keys()),
            help="Choisir le bon : fournisseur et camion préremplis.",
        )
        bc = options_bc[choix_bc]
        ref_bc = bc["Reference"] if bc is not None else ""
        if bc is None:
            st.sidebar.info("Aucun BC — passez en mode « Sans bon » ou choisissez un BC.")
else:
    # VENTE : pas de BC — uniquement Proformas NAP (proformas.xlsx / page Proformas)
    st.sidebar.info(
        "🚢 **VENTE** : pas de bon de commande.\n\n"
        "Document lié = **Proforma NAP** uniquement (pas les factures MITC / autres)."
    )
    options_pf = {"— Sans proforma NAP —": ""}
    df_pf = None
    if pf is None:
        st.sidebar.warning(
            "Module `proforma_data` introuvable. "
            "Placez `proforma_data.py` à la racine du projet."
        )
    else:
        try:
            df_pf = pf.charger()
        except Exception as e:
            st.sidebar.error(f"Lecture proformas.xlsx impossible : {e}")
            df_pf = None
        if df_pf is not None and not df_pf.empty:
            for _, r in df_pf.sort_values("ID", ascending=False).iterrows():
                lab = (
                    f"NAP {r['Numero']} · {r['Client']} · {r['Produit']} · "
                    f"{float(r['Montant_TTC']):,.0f} FCFA".replace(",", " ")
                )
                options_pf[lab] = str(r["Numero"])
        else:
            st.sidebar.caption("Aucune proforma NAP — créez-en dans la page **Proformas**.")
    choix_pf = st.sidebar.selectbox(
        "Proforma NAP (optionnel)",
        list(options_pf.keys()),
        help="Uniquement les proformas NAP (fichier proformas.xlsx). "
             "Les autres factures du système ne sont pas proposées.",
    )
    ref_proforma = options_pf[choix_pf]
    ref_bc = ref_proforma  # traçabilité dans colonne Ref_BC

# Fournisseur du bon : retrouvé automatiquement (même s'il n'est pas encore dans fournisseurs.xlsx)
cle_bon, nom_bon, bon_connu = (None, None, True)
if type_mvt == "ACHAT" and bc is not None:
    cle_bon, nom_bon, bon_connu = dd.resoudre_fournisseur_bon(frs, bc["Code_Fournisseur"], bc["Nom_Fournisseur"])
    if cle_bon is not None and cle_bon not in frs:
        frs[cle_bon] = nom_bon

if type_mvt == "ACHAT" and cle_bon is not None and cle_bon in frs:
    st.sidebar.caption(f"Fournisseur du bon : **{frs[cle_bon]}**")
    if not bon_connu:
        st.sidebar.warning("Ce fournisseur n'est pas dans fournisseurs.xlsx.")

dernier = df.sort_values("Date_Dechargement").iloc[-1] if not df.empty else None
pu_defaut = int(dernier["PU_FCFA_KG"]) if dernier is not None else 0
lieu_defaut = dernier["Lieu_Stockage"] if dernier is not None else ""
cle_form = ref_bc or "libre"

# Aide métier : échange de bons entre fournisseurs
if bc is not None and nom_bon:
    st.sidebar.info(
        f"**BC** : `{ref_bc}`\n\n"
        f"**Fournisseur du bon** : {nom_bon}\n\n"
        "Si le **ticket de déchargement** porte un autre nom, choisissez ce nom ci-dessous : "
        "le déchargement sera affecté au fournisseur du ticket, avec conservation du N° BC."
    )

# Préparer listes selon type
opts_frs = {dd.label_fournisseur(k, n): (k, n) for k, n in sorted(frs.items(), key=lambda x: x[1])}
labels_frs = ["— Choisir —"] + list(opts_frs.keys())
idx_frs = 0
if type_mvt == "ACHAT" and cle_bon is not None and cle_bon in frs:
    try:
        idx_frs = 1 + list(opts_frs.keys()).index(dd.label_fournisseur(cle_bon, frs[cle_bon]))
    except Exception:
        idx_frs = 0

with st.sidebar.form("form_dechargement", clear_on_submit=True):
    if type_mvt == "VENTE":
        # Pré-sélection client si proforma NAP choisie
        idx_cli = 0
        if ref_proforma and pf is not None:
            try:
                _dfp = pf.charger()
                _row = _dfp[_dfp["Numero"].astype(str) == str(ref_proforma)]
                if not _row.empty:
                    _cli_pf = str(_row.iloc[0].get("Client") or "").strip()
                    for i, lab in enumerate(labels_cli):
                        if i and _cli_pf and _cli_pf.casefold() in lab.casefold():
                            idx_cli = i
                            break
            except Exception:
                pass
        frs_sel = st.selectbox(
            "Client *",
            labels_cli,
            index=min(idx_cli, max(len(labels_cli) - 1, 0)),
            key=f"cli_{cle_form}",
            help="Client local ou international (page Clients). "
                 "Si une proforma NAP est choisie, le client de la proforma est proposé.",
        )
        if ref_proforma:
            st.caption(f"Proforma NAP liée : **{ref_proforma}**")
        if not opts_cli:
            st.caption("Aucun client : créez-en dans la page **Clients**.")
    else:
        frs_sel = st.selectbox(
            "Fournisseur du ticket (livreur réel) *",
            labels_frs,
            index=idx_frs,
            key=f"frs_{cle_form}",
            help="Nom figurant sur le ticket de déchargement / pesée.",
        )
    date_d = st.date_input("Date de déchargement *", value=date.today())
    transaction = st.text_input("N° Transaction (pesée) *", value=dd.prochain_transaction(df))
    ticket = st.text_input(
        "Weighment ticket / Ticket de pesée",
        help="⚠️ Chaque N° de ticket de pesée doit être unique. "
             "Doublon = même ticket + même type (ACHAT ou VENTE). Un ticket ACHAT peut coexister avec le même N° en VENTE.",
    )
    if bc is not None:
        st.caption(f"Camion prévu sur le bon : **{bc['Numero_Tracteur'] or '—'}**")
    camion = st.text_input("N° Camion (réellement livré)", value=bc["Numero_Tracteur"] if bc is not None else "",
                           key=f"cam_{cle_form}", help="Modifiez-le si le camion arrivé n'est pas celui du bon.")
    motif = st.text_input(
        "Remarque / motif (camion différent ou échange de bon)",
        key=f"mot_{cle_form}",
        placeholder="ex. Bon échangé avec TCHANIGA · camion de remplacement…",
    )
    poids = st.number_input("Poids net (KG) *", min_value=0, step=10, value=0)
    sacs = st.number_input("Nombre de sacs", min_value=0, step=1, value=0)
    pu = st.number_input("PU (FCFA/KG) *", min_value=0, step=1, value=pu_defaut)
    lieu = st.text_input("Lieu de stockage", value=lieu_defaut)
    st.caption("Montant dû = Poids net × PU (calculé automatiquement)")
    envoye = st.form_submit_button("💾 Enregistrer le déchargement", width="stretch")

if envoye:
    err = None
    if type_mvt == "VENTE":
        if frs_sel == "— Choisir un client —" or frs_sel not in opts_cli:
            err = "Choisissez un client (page Clients)."
        else:
            cle, nom_frs = opts_cli[frs_sel]
    else:
        if frs_sel == "— Choisir —" or frs_sel not in opts_frs:
            err = "Choisissez un fournisseur."
        else:
            cle, nom_frs = opts_frs[frs_sel]
    if err:
        st.sidebar.error(err)
    elif not transaction.strip():
        st.sidebar.error("Le N° de transaction est obligatoire.")
    elif poids <= 0:
        st.sidebar.error("Le poids net doit être supérieur à 0.")
    elif pu <= 0:
        st.sidebar.error("Le prix unitaire doit être supérieur à 0.")
    else:
        # Fournisseur du BC (émetteur du bon) vs contrepartie ticket
        frs_bc_nom = (nom_bon or "") if (bc is not None and type_mvt == "ACHAT") else ""
        frs_bc_code = (cle_bon or "") if (bc is not None and type_mvt == "ACHAT") else ""
        affectation = bool(
            type_mvt == "ACHAT"
            and frs_bc_nom and nom_frs
            and frs_bc_nom.casefold().strip() != nom_frs.casefold().strip()
        )
        payload = {
            "Transaction": transaction.strip(),
            "Type_Mouvement": type_mvt,
            "Date_Dechargement": pd.Timestamp(date_d),
            "Ref_BC": ref_bc if type_mvt == "ACHAT" else ref_bc,
            "Code_Frs": cle if cle != nom_frs else "",
            "Fournisseur": nom_frs,  # fournisseur (ACHAT) ou client (VENTE)
            "Fournisseur_BC": frs_bc_nom,
            "Code_Frs_BC": frs_bc_code if isinstance(frs_bc_code, str) else "",
            "Affectation_BC": "OUI" if affectation else "NON",
            "N_Camion": camion.strip(),
            "Camion_BC": bc["Numero_Tracteur"] if bc is not None else "",
            "Remarque": motif.strip(),
            "Weighment_Ticket": ticket.strip(),
            "Poids_Net_KG": poids,
            "Nb_Sacs": sacs,
            "PU_FCFA_KG": pu,
            "Lieu_Stockage": lieu.strip(),
            "Creer_BC_Auto": bool(type_mvt == "ACHAT" and sans_bc and confirmer_creation_bc),
        }
        if type_mvt == "ACHAT" and sans_bc and not confirmer_creation_bc:
            st.sidebar.error("Cochez la confirmation pour créer automatiquement le bon de commande.")
        elif type_mvt == "ACHAT" and (not sans_bc) and not ref_bc:
            st.sidebar.error("Sélectionnez un bon de commande, ou passez en mode « Sans bon ».")
        else:
            ok, msg = dd.ajouter_dechargement(payload, name)
            if ok:
                extra = ""
                if affectation:
                    extra += (
                        f" 🔄 **Affectation BC** : bon émis pour **{frs_bc_nom}**, "
                        f"déchargement compté pour **{nom_frs}** (ticket). "
                        f"N° BC conservé : `{ref_bc}`."
                    )
                if bc is not None and bc.get("Qte_KG", 0) > 0:
                    cumul = df[df["Ref_BC"] == ref_bc]["Poids_Net_KG"].sum() + poids if not df.empty else poids
                    ecart = f"{cumul - bc['Qte_KG']:+,.0f}".replace(",", " ")
                    extra += (
                        f" Poids cumulé sur ce bon : {dd.fmt(cumul, 'kg')} pour {dd.fmt(bc['Qte_KG'], 'kg')} "
                        f"commandés (écart {ecart} kg)."
                    )
                if bc is not None and dd.camion_different(bc["Numero_Tracteur"], camion):
                    extra += f" ⚠️ Camion différent du bon (bon : {bc['Numero_Tracteur']} · livré : {camion.strip()})."
                flash("success", msg + extra)
                st.rerun()
            else:
                st.sidebar.error(msg)

# ==================== PAGE PRINCIPALE ====================
st.title(f"🚛 Déchargements & traçabilité — {dd.libelle_produit(produit)}")

if "_flash" in st.session_state:
    type_, message = st.session_state.pop("_flash")
    getattr(st, type_)(message)

if df.empty:
    st.info(f"Aucun déchargement pour {dd.libelle_produit(produit)} pour le moment. Ajoutez-en un (menu de gauche) ou importez votre fichier Excel "
            "dans l'onglet « 📥 Importer mon Excel ».")
else:
    if "Type_Mouvement" not in df.columns:
        df["Type_Mouvement"] = "ACHAT"
    n_achat = int((df["Type_Mouvement"] == "ACHAT").sum())
    n_vente = int((df["Type_Mouvement"] == "VENTE").sum())
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Camions", dd.fmt(len(df)))
    k2.metric("🛒 Achat", dd.fmt(n_achat))
    k3.metric("🚢 Vente", dd.fmt(n_vente))
    k4.metric("Poids net", f"{df['Poids_Net_KG'].sum() / 1000:,.1f} T".replace(",", " "))
    k5.metric("Dû (FCFA)", dd.fmt(df["Montant_Du"].sum()))
    k6.metric("Reste à payer", dd.fmt(df["Solde"].sum()))

# ==================== VÉRIFICATION CAMION / TICKET ====================
st.markdown("---")
st.subheader("🔎 Vérifier un déchargement (camion ou ticket pont)")
st.caption(
    "Saisissez un **N° de camion** ou un **N° de ticket de pesée (Weighment / Ticket pont)**. "
    "La recherche ignore espaces et tirets. Filtre optionnel ACHAT / VENTE."
)
vc1, vc2, vc3 = st.columns([3, 1, 1])
with vc1:
    q_verif = st.text_input(
        "N° camion ou N° ticket pont",
        key="verif_camion_ticket",
        placeholder="Ex. BE5458RB  ou  123456  ou  WT-9087",
    )
with vc2:
    f_type_v = st.selectbox("Type", ["Tous", "ACHAT", "VENTE"], key="verif_type")
with vc3:
    st.write("")  # align
    st.write("")
    btn_verif = st.button("🔎 Vérifier", type="primary", key="btn_verif_dech")

if btn_verif or (q_verif and len(q_verif.strip()) >= 2 and st.session_state.get("_auto_verif") == q_verif):
    st.session_state["_auto_verif"] = q_verif
    if not q_verif or len(q_verif.strip()) < 2:
        st.warning("Saisissez au moins 2 caractères.")
    else:
        tm = None if f_type_v == "Tous" else f_type_v
        res = dd.rechercher_dechargement(q_verif.strip(), type_mouvement=tm)
        if res is None or res.empty:
            st.error(
                f"❌ **Aucun déchargement** pour « {q_verif.strip()} » "
                f"sur le produit **{dd.libelle_produit(produit)}**"
                + (f" (type {f_type_v})" if f_type_v != "Tous" else "")
                + "."
            )
            st.info("Vérifiez le produit dans le menu de gauche, ou le type ACHAT / VENTE.")
        else:
            # Enrichir avec calculs si besoin
            res_c, _ = dd.charger_complet()
            trs = set(res["Transaction"].map(dd._s))
            res_c = res_c[res_c["Transaction"].map(dd._s).isin(trs)] if not res_c.empty else res
            st.success(
                f"✅ **{len(res_c)} déchargement(s)** trouvé(s) pour « {q_verif.strip()} » "
                f"— produit {dd.libelle_produit(produit)}."
            )
            aff_v = res_c.copy()
            if "Statut_Paiement" in aff_v.columns:
                aff_v["Statut_Paiement"] = aff_v["Statut_Paiement"].map(ICONES)
            if "Camion_Diff" in aff_v.columns:
                aff_v["Alerte_Camion"] = ["⚠️ Différent" if x else "" for x in aff_v["Camion_Diff"]]
            cols_v = [c for c in [
                "Date_Dechargement", "Transaction", "Type_Mouvement", "Ref_BC",
                "Fournisseur", "N_Camion", "Camion_BC", "Alerte_Camion",
                "Weighment_Ticket", "Poids_Net_KG", "PU_FCFA_KG", "Montant_Du",
                "Montant_Paye", "Solde", "Statut_Paiement", "Lieu_Stockage", "Remarque",
            ] if c in aff_v.columns]
            st.dataframe(
                aff_v[cols_v],
                hide_index=True,
                width="stretch",
                column_config=CONFIG_COLONNES,
            )
            # Résumé rapide
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Lignes", len(res_c))
            m2.metric("Poids total", f"{res_c['Poids_Net_KG'].sum()/1000:,.2f} T".replace(",", " ") if "Poids_Net_KG" in res_c.columns else "—")
            m3.metric("Montant dû", dd.fmt(res_c["Montant_Du"].sum()) if "Montant_Du" in res_c.columns else "—")
            if "Statut_Paiement" in res_c.columns:
                n_paye = int((res_c["Statut_Paiement"] == "PAYE").sum())
                m4.metric("Payés", f"{n_paye}/{len(res_c)}")

st.markdown("---")


# ==================== AJUSTER QUANTITÉ BC (éviter EN COURS artificiel) ====================
with st.expander("✏️ Ajuster la quantité d'un bon de commande (lié aux déchargements)"):
    st.caption(
        "Si le poids réellement déchargé diffère de la quantité commandée, vous pouvez "
        "**mettre à jour la quantité du BC** pour que le statut ne reste pas « EN COURS » à tort. "
        "Uniquement pour les bons **liés** à au moins un déchargement du produit actuel."
    )
    if df.empty:
        st.info("Aucun déchargement — pas de BC lié à ajuster.")
    else:
        refs_liees = sorted({
            dd._s(r) for r in df["Ref_BC"].tolist()
            if dd._s(r) and "-BC-" in dd._s(r).upper()
        }, reverse=True)
        if not refs_liees:
            st.info("Aucun déchargement lié à un N° de bon de commande.")
        else:
            ref_adj = st.selectbox("Bon de commande lié", refs_liees, key="adj_bc_ref")
            # Quantité actuelle dans donnees_bons
            qte_act = "—"
            qte_kg_act = 0.0
            try:
                _bons_all = dd.charger_bons_commande()
                _rowb = _bons_all[_bons_all["Reference"].map(dd._s) == ref_adj] if not _bons_all.empty else _bons_all
                if not _rowb.empty:
                    qte_act = dd._s(_rowb.iloc[0].get("Quantite"))
                    qte_kg_act = float(_rowb.iloc[0].get("Qte_KG") or 0)
            except Exception:
                pass
            poids_livre = dd.poids_livre_sur_bc(ref_adj, df)
            c1, c2, c3 = st.columns(3)
            c1.metric("Quantité actuelle sur le BC", qte_act or "—")
            c2.metric("Poids déjà déchargé", f"{poids_livre:,.0f} kg".replace(",", " "))
            c3.metric("Écart (livré − commandé)", f"{poids_livre - qte_kg_act:+,.0f} kg".replace(",", " ") if qte_kg_act else "—")
            mode_adj = st.radio(
                "Nouvelle quantité",
                [
                    "Aligner sur le poids total déchargé",
                    "Saisir manuellement (ex. 48.5 MT)",
                ],
                key="adj_bc_mode",
            )
            qte_manuelle = ""
            if mode_adj.startswith("Saisir"):
                qte_manuelle = st.text_input("Quantité (ex. 48.5 MT ou 48500 KG)", key="adj_bc_qte")
            st.warning(
                "Cette action modifie **donnees_bons.xlsx**. "
                "Le Word du bon n'est pas régénéré automatiquement."
            )
            if st.button("💾 Mettre à jour la quantité du bon", type="primary", key="btn_adj_bc"):
                if mode_adj.startswith("Aligner"):
                    if poids_livre <= 0:
                        st.error("Aucun poids déchargé sur ce bon.")
                    else:
                        ok, msg = dd.modifier_quantite_bon(ref_adj, poids_kg=poids_livre, utilisateur=name)
                        flash("success" if ok else "error", msg)
                        st.rerun()
                else:
                    if not qte_manuelle.strip():
                        st.error("Saisissez la nouvelle quantité.")
                    else:
                        ok, msg = dd.modifier_quantite_bon(ref_adj, quantite_texte=qte_manuelle.strip(), utilisateur=name)
                        flash("success" if ok else "error", msg)
                        st.rerun()


tab1, tab2, tab3, tab4 = st.tabs(["📋 Déchargements", "✅ Approuver un paiement", "⚖️ Rapport des poids",
                                  "📥 Importer mon Excel"])

# ---------------------------------------------------------------- TAB 1
with tab1:
    # --- Alerte doublons tickets de pesée par type ACHAT / VENTE ---
    if not df.empty and "Weighment_Ticket" in df.columns:
        tmp = df.copy()
        tmp["_ticket"] = tmp["Weighment_Ticket"].map(dd._s)
        if "Type_Mouvement" in tmp.columns:
            tmp["_type"] = tmp["Type_Mouvement"].map(lambda x: dd._norm_type_mvt(x) if hasattr(dd, "_norm_type_mvt") else ("VENTE" if str(x).upper().startswith("V") else "ACHAT"))
        else:
            tmp["_type"] = "ACHAT"
        non_vide = tmp[tmp["_ticket"] != ""].copy()
        if not non_vide.empty:
            non_vide["_cle"] = non_vide["_ticket"].str.casefold() + "|" + non_vide["_type"]
            counts = non_vide["_cle"].value_counts()
            dups_keys = set(counts[counts > 1].index)
            if dups_keys:
                st.error(
                    f"⚠️ **Ticket(s) de pesée en doublon** (même N° + même type ACHAT ou VENTE). "
                    "Un ticket peut servir une fois en ACHAT et une fois en VENTE, pas deux fois le même type."
                )
                lignes_dup = non_vide[non_vide["_cle"].isin(dups_keys)]
                cols = [c for c in [
                    "Date_Dechargement", "Transaction", "Type_Mouvement", "Weighment_Ticket", "Fournisseur",
                    "N_Camion", "Poids_Net_KG", "Ref_BC", "Montant_Du",
                ] if c in lignes_dup.columns]
                st.dataframe(
                    lignes_dup[cols],
                    hide_index=True,
                    width="stretch",
                    height=min(260, 40 + 28 * len(lignes_dup)),
                )
    if df.empty:
        st.info("Rien à afficher.")
    else:
        d_min, d_max = df["Date_Dechargement"].min().date(), df["Date_Dechargement"].max().date()
        f0, f1, f2, f3, f4 = st.columns(5)
        f_type = f0.selectbox("Type", ["Tous", "ACHAT", "VENTE"], key="l_type")
        f_frs = f1.selectbox("Fournisseur", ["Tous"] + sorted(df["Fournisseur"].unique().tolist()), key="l_frs")
        f_stat = f2.selectbox("Statut paiement", ["Tous", "NON PAYE", "PARTIEL", "PAYE", "SURPLUS"], key="l_stat")
        f_q = f3.text_input("Recherche", placeholder="N° BC, camion, ticket…", key="l_q")
        f_per = f4.date_input("Période", value=(d_min, d_max), key="l_per")
        f_diff = st.checkbox("⚠️ Camion différent du bon", key="l_diff")
        f_aff = st.checkbox("🔄 Seulement les échanges de bons (fournisseur ticket ≠ fournisseur BC)", key="l_aff")

        v = df.copy()
        if f_type != "Tous" and "Type_Mouvement" in v.columns:
            v = v[v["Type_Mouvement"] == f_type]
        if f_frs != "Tous":
            v = v[v["Fournisseur"] == f_frs]
        if f_stat != "Tous":
            v = v[v["Statut_Paiement"] == f_stat]
        if f_diff:
            v = v[v["Camion_Diff"]]
        if f_aff:
            if "Affectation_BC" in v.columns:
                v = v[v["Affectation_BC"].astype(str).str.upper() == "OUI"]
            elif "Fournisseur_BC" in v.columns:
                v = v[v.apply(
                    lambda r: bool(str(r.get("Fournisseur_BC") or "").strip())
                    and str(r.get("Fournisseur_BC") or "").casefold() != str(r.get("Fournisseur") or "").casefold(),
                    axis=1,
                )]
        if f_q.strip():
            q = f_q.strip().lower()
            if not v.empty:
                v = v[v.apply(lambda r: q in " ".join(r.astype(str)).lower(), axis=1)]
        if isinstance(f_per, (tuple, list)) and len(f_per) == 2:
            v = v[(v["Date_Dechargement"].dt.date >= f_per[0]) & (v["Date_Dechargement"].dt.date <= f_per[1])]

        v = v.sort_values(["Date_Dechargement", "Transaction"], ascending=False)
        tonnes = f"{v['Poids_Net_KG'].sum() / 1000:,.1f}".replace(",", " ")
        st.caption(f"{len(v)} livraison(s) · {tonnes} T · dû {dd.fmt(v['Montant_Du'].sum())} · "
                   f"payé {dd.fmt(v['Montant_Paye'].sum())} · reste {dd.fmt(v['Solde'].sum())} FCFA")
        st.dataframe(tableau_dechargements(v), hide_index=True, width="stretch", height=450,
                     column_config=CONFIG_COLONNES)
        st.download_button("⬇️ Télécharger cette liste (Excel)",
                           data=dd.to_excel_bytes({"Dechargements": tableau_dechargements(v)}),
                           file_name="dechargements.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # ---------- MODIFIER UN DÉCHARGEMENT ----------
        st.markdown("---")
        st.subheader("✏️ Modifier un déchargement (PU, poids, camion…)")
        tr_edit = st.selectbox(
            "Transaction à modifier",
            [""] + df["Transaction"].tolist(),
            key="edit_tr",
            format_func=lambda x: "— Choisir —" if x == "" else x
        )
        if tr_edit:
            L_edit = df[df["Transaction"] == tr_edit].iloc[0]
            with st.form("form_mod_dech"):
                frs_bc_e = L_edit["Fournisseur_BC"] if "Fournisseur_BC" in L_edit.index else ""
                aff_e = L_edit["Affectation_BC"] if "Affectation_BC" in L_edit.index else ""
                st.caption(
                    f"Transaction **{tr_edit}** · Fournisseur ticket : **{L_edit['Fournisseur']}** · "
                    f"Fournisseur BC : **{frs_bc_e or '—'}** · "
                    f"Affectation : **{aff_e or 'NON'}** · "
                    f"Bon : {L_edit['Ref_BC'] or '—'} · Montant dû : {dd.fmt(L_edit['Montant_Du'])}"
                )
                e1, e2, e3 = st.columns(3)
                with e1:
                    edit_date = st.date_input(
                        "Date de déchargement",
                        value=L_edit["Date_Dechargement"].date() if pd.notna(L_edit["Date_Dechargement"]) else date.today()
                    )
                    edit_poids = st.number_input("Poids net (KG) *", min_value=0, step=10,
                                                 value=int(L_edit["Poids_Net_KG"]))
                    edit_sacs = st.number_input("Nombre de sacs", min_value=0, step=1,
                                                value=int(L_edit["Nb_Sacs"]))
                with e2:
                    edit_pu = st.number_input("PU (FCFA/KG) *", min_value=0, step=1,
                                              value=int(L_edit["PU_FCFA_KG"]))
                    edit_camion = st.text_input("N° Camion livré", value=str(L_edit["N_Camion"] or ""))
                    edit_ticket = st.text_input(
                        "Weighment ticket / Ticket de pesée",
                        value=str(L_edit["Weighment_Ticket"] or ""),
                        help="Doit rester unique : un ticket déjà utilisé sur une autre transaction sera refusé.",
                    )
                with e3:
                    _tm = str(L_edit.get("Type_Mouvement") or "ACHAT")
                    edit_type = st.selectbox("Type Achat/Vente", ["ACHAT", "VENTE"],
                                            index=0 if _tm != "VENTE" else 1)
                    edit_lieu = st.text_input("Lieu de stockage", value=str(L_edit["Lieu_Stockage"] or ""))
                    edit_remarque = st.text_input("Remarque", value=str(L_edit["Remarque"] or ""))
                    edit_ref_bc = st.text_input("N° Bon de commande", value=str(L_edit["Ref_BC"] or ""))
                st.info(f"Montant dû recalculé = {edit_poids} kg × {edit_pu} = **{edit_poids * edit_pu:,}** FCFA"
                        .replace(",", " "))
                btn_mod = st.form_submit_button("💾 Enregistrer les modifications", type="primary", width="stretch")
                if btn_mod:
                    if edit_poids <= 0 or edit_pu <= 0:
                        st.error("Poids et PU doivent être > 0.")
                    else:
                        ok, msg = dd.modifier_dechargement(tr_edit, {
                            "Date_Dechargement": edit_date,
                            "Type_Mouvement": edit_type,
                            "Poids_Net_KG": edit_poids,
                            "Nb_Sacs": edit_sacs,
                            "PU_FCFA_KG": edit_pu,
                            "N_Camion": edit_camion,
                            "Weighment_Ticket": edit_ticket,
                            "Lieu_Stockage": edit_lieu,
                            "Remarque": edit_remarque,
                            "Ref_BC": edit_ref_bc,
                        }, name)
                        flash("success" if ok else "error", msg)
                        st.rerun()

        # ---------- SUPPRIMER UN DÉCHARGEMENT ----------
        if est_admin or peut_approuver:
            with st.expander("🗑️ Supprimer un déchargement"):
                tr_sup = st.selectbox("Transaction à supprimer", df["Transaction"].tolist(), key="sup_tr")
                st.warning("La ligne ira en **corbeille** (restorable). Les paiements affectés suivent en corbeille.")
                if st.checkbox("Je confirme la suppression", key="sup_conf") and st.button("🗑️ Supprimer", type="primary"):
                    ok, msg = dd.supprimer_dechargement(tr_sup, utilisateur=name, motif="Suppression depuis la liste")
                    flash("success" if ok else "error", msg)
                    st.rerun()

        # ---------- SUPPRIMER UN PAIEMENT AFFECTÉ ----------
        if (est_admin or peut_approuver) and not aff.empty:
            with st.expander("🗑️ Supprimer un paiement affecté (détacher un bon de paiement d'une livraison)"):
                opts_aff = {}
                for r in aff.sort_values("ID", ascending=False).itertuples():
                    lab = f"ID {r.ID} · Trans {r.Transaction} · {dd.fmt(r.Montant)} FCFA · {r.Bons_Paiement}"
                    opts_aff[lab] = r.ID
                choix_aff = st.selectbox("Paiement à retirer", list(opts_aff.keys()), key="sup_aff")
                if st.checkbox("Je confirme le retrait de ce paiement", key="sup_aff_conf") and \
                        st.button("🗑️ Retirer ce paiement", type="primary", key="btn_sup_aff"):
                    ok, msg = dd.supprimer_affectation(opts_aff[choix_aff])
                    flash("success" if ok else "error", msg)
                    st.rerun()

        # ---------- CORBEILLE / HISTORIQUE SUPPRESSIONS ----------
        with st.expander("🗑️ Corbeille — historique des suppressions (restaurer)"):
            try:
                corb_d, corb_a = dd.charger_corbeille()
            except Exception as e:
                corb_d, corb_a = pd.DataFrame(), pd.DataFrame()
                st.caption(f"Corbeille indisponible : {e}")
            if corb_d is None or corb_d.empty:
                st.info("Corbeille vide — aucune suppression à annuler.")
            else:
                st.caption(
                    f"**{len(corb_d)}** déchargement(s) en corbeille · "
                    f"**{len(corb_a) if corb_a is not None and not corb_a.empty else 0}** paiement(s) affecté(s) associés."
                )
                cols_c = [c for c in [
                    "Supprime_Le", "Supprime_Par", "Motif_Suppression",
                    "Date_Dechargement", "Transaction", "Type_Mouvement", "Ref_BC",
                    "Fournisseur", "Weighment_Ticket", "N_Camion", "Poids_Net_KG",
                    "Montant_Du", "Remarque",
                ] if c in corb_d.columns]
                st.dataframe(corb_d[cols_c], hide_index=True, width="stretch", height=280,
                             column_config={
                                 "Date_Dechargement": st.column_config.DateColumn("Date déch.", format="DD/MM/YYYY"),
                                 "Poids_Net_KG": st.column_config.NumberColumn("Poids kg", format="%d"),
                                 "Montant_Du": st.column_config.NumberColumn("Dû", format="%d"),
                             })
                if est_admin or peut_approuver:
                    opts_corb = corb_d["Transaction"].map(dd._s).tolist()
                    tr_rest = st.selectbox("Transaction à restaurer", opts_corb, key="corb_rest")
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        if st.button("♻️ Restaurer cette transaction", type="primary", key="btn_rest"):
                            ok, msg = dd.restaurer_dechargement(tr_rest)
                            flash("success" if ok else "error", msg)
                            st.rerun()
                    with c2:
                        if st.checkbox("Confirmer suppression définitive", key="corb_purge_one"):
                            if st.button("⛔ Effacer définitivement", key="btn_purge_one"):
                                ok, msg = dd.purger_corbeille(tr_rest)
                                flash("success" if ok else "error", msg)
                                st.rerun()
                    with c3:
                        if st.checkbox("Confirmer vider toute la corbeille", key="corb_purge_all"):
                            if st.button("🗑️ Vider la corbeille", key="btn_purge_all"):
                                ok, msg = dd.purger_corbeille(None)
                                flash("success" if ok else "error", msg)
                                st.rerun()
                else:
                    st.info("Restauration / purge réservées aux comptes : " + ", ".join(dd.ROLES_APPROBATION))

# ---------------------------------------------------------------- TAB 2
with tab2:
    st.subheader("✅ Approuver le paiement d'une livraison")
    st.caption("1) Choisissez la livraison · 2) Choisissez le(s) bon(s) de paiement · 3) Cliquez sur le bouton.")
    if df.empty:
        st.info("Aucune livraison à traiter.")
    else:
        c1, c2 = st.columns(2)
        a_frs = c1.selectbox("Fournisseur", ["Tous"] + sorted(df["Fournisseur"].unique().tolist()), key="a_frs")
        a_stat = c2.selectbox("Afficher", ["À payer (non payé + partiel)", "Tous", "NON PAYE", "PARTIEL", "PAYE",
                                           "SURPLUS"], key="a_stat")
        cand = df.copy()
        if a_frs != "Tous":
            cand = cand[cand["Fournisseur"] == a_frs]
        if a_stat.startswith("À payer"):
            cand = cand[cand["Statut_Paiement"].isin(["NON PAYE", "PARTIEL"])]
        elif a_stat != "Tous":
            cand = cand[cand["Statut_Paiement"] == a_stat]
        cand = cand.sort_values("Date_Dechargement", ascending=False)

        if cand.empty:
            st.success("Aucune livraison dans ce filtre. 🎉")
        else:
            opts = {}
            for r in cand.itertuples():
                lab = (f"{r.Transaction} · {r.Ref_BC or 'sans BC'} · {r.Fournisseur} · "
                       f"{dd.fmt(r.Poids_Net_KG)} kg · solde {dd.fmt(r.Solde)} · {r.Statut_Paiement}")
                opts[lab] = r.Transaction
            choix = st.selectbox("Livraison à traiter", list(opts.keys()), key="a_choix")
            tr = opts[choix]
            L = df[df["Transaction"] == tr].iloc[0]
            sfx = f"{tr}_{len(aff)}"  # change après chaque approbation → champs remis à zéro

            st.markdown(f"**{L['Fournisseur']}** · Bon de commande : **{L['Ref_BC'] or 'aucun'}** · "
                        f"Camion : {L['N_Camion'] or '—'} · Ticket : {L['Weighment_Ticket'] or '—'}")
            if L["Camion_Diff"]:
                st.warning(f"⚠️ Camion différent : bon = {L['Camion_BC']} · livré = {L['N_Camion']}"
                           + (f" · motif : {L['Remarque']}" if L["Remarque"] else ""))
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Poids net", dd.fmt(L["Poids_Net_KG"], "kg"))
            m2.metric("PU", dd.fmt(L["PU_FCFA_KG"]))
            m3.metric("Montant dû", dd.fmt(L["Montant_Du"]))
            m4.metric("Déjà payé", dd.fmt(L["Montant_Paye"]))
            m5.metric("Solde", dd.fmt(L["Solde"]), delta=ICONES[L["Statut_Paiement"]], delta_color="off")

            deja = aff[aff["Transaction"] == tr]
            if not deja.empty:
                st.markdown("**Paiements déjà approuvés pour cette livraison**")
                st.dataframe(deja[["Date_Approbation", "Bons_Paiement", "Montant", "Approuve_Par"]], hide_index=True,
                             width="stretch", column_config={
                                 "Date_Approbation": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
                                 "Bons_Paiement": "Bon(s) de paiement",
                                 "Montant": st.column_config.NumberColumn("Montant", format="%d"),
                                 "Approuve_Par": "Approuvé par"})

            st.markdown("---")
            if not peut_approuver:
                st.info("🔒 Seuls les comptes " + ", ".join(dd.ROLES_APPROBATION) + " peuvent approuver un paiement. "
                        "Vous pouvez consulter la situation.")
            else:
                journal = dd.bons_paiement_achats(dd.charger_journal())
                tous = st.checkbox("Afficher les bons de paiement de tous les fournisseurs", value=False,
                                   key=f"tous_{sfx}")
                jr = journal if tous else journal[journal["Prestataire"].map(lambda p: dd.meme_nom(p, L["Fournisseur"]))]
                utilises = {b for t in aff["Bons_Paiement"] for b in dd.liste_bons_texte(t)}
                labels = {}
                for r in jr.itertuples():
                    d_txt = r.Date.strftime("%d/%m/%Y") if pd.notna(r.Date) else ""
                    suite = " (déjà utilisé)" if r.No_Transaction in utilises else ""
                    labels[f"{r.No_Transaction} · {r.Prestataire} · {dd.fmt(r.Montant)} FCFA · {d_txt}{suite}"] = r.No_Transaction
                sel = st.multiselect("Bon(s) de paiement (journal des paiements)", list(labels.keys()), key=f"bons_{sfx}",
                                     help="Plusieurs bons possibles pour une même livraison.")
                autre = st.text_input("Autre bon non listé (ex. CHQ-0390-2026, séparés par &)", key=f"autre_{sfx}")
                bons_choisis = [labels[s] for s in sel] + dd.liste_bons_texte(autre)
                montant = st.number_input("Montant à enregistrer (FCFA) — utile seulement pour un paiement partiel",
                                          min_value=0, step=1000, value=int(max(L["Solde"], 0)), key=f"mt_{sfx}")

                b1, b2, b3 = st.columns(3)
                if b1.button("✅ Approuver PAYÉ (solde complet)", type="primary", width="stretch", key=f"ok_{sfx}"):
                    if L["Solde"] <= 0:
                        st.warning("Cette livraison est déjà soldée.")
                    elif not bons_choisis:
                        st.error("Choisissez au moins un bon de paiement.")
                    else:
                        ok, msg = dd.ajouter_affectation(tr, " & ".join(bons_choisis), L["Solde"], name)
                        flash("success" if ok else "error",
                              f"✅ Livraison {tr} approuvée PAYÉE ({dd.fmt(L['Solde'])} FCFA) — {' & '.join(bons_choisis)}")
                        st.rerun()
                if b2.button("💾 Enregistrer ce montant (partiel)", width="stretch", key=f"part_{sfx}"):
                    if montant <= 0:
                        st.error("Saisissez un montant supérieur à 0.")
                    elif not bons_choisis:
                        st.error("Choisissez au moins un bon de paiement.")
                    else:
                        ok, msg = dd.ajouter_affectation(tr, " & ".join(bons_choisis), montant, name)
                        nouveau = dd.statut_paiement(L["Montant_Du"], L["Montant_Paye"] + montant)
                        flash("success" if ok else "error",
                              f"💾 {dd.fmt(montant)} FCFA enregistrés sur la livraison {tr} → {ICONES[nouveau]}")
                        st.rerun()
                with b3:
                    conf = st.checkbox("Confirmer l'annulation", key=f"cf_{sfx}")
                    if st.button("❌ Remettre NON PAYÉ", width="stretch", key=f"np_{sfx}",
                                 disabled=L["Montant_Paye"] <= 0):
                        if not conf:
                            st.warning("Cochez « Confirmer l'annulation ».")
                        else:
                            ok, msg = dd.annuler_paiements(tr)
                            flash("success", "❌ " + msg)
                            st.rerun()

        st.markdown("---")
        st.markdown("**Derniers paiements approuvés**")
        if aff.empty:
            st.caption("Aucun paiement approuvé pour le moment.")
        else:
            dern = aff.merge(df[["Transaction", "Ref_BC", "Fournisseur"]], on="Transaction", how="left")
            dern = dern.sort_values("ID", ascending=False).head(15)
            st.dataframe(dern[["Date_Approbation", "Transaction", "Ref_BC", "Fournisseur", "Bons_Paiement", "Montant",
                               "Approuve_Par"]], hide_index=True, width="stretch", column_config={
                "Date_Approbation": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
                "Ref_BC": "N° BC / Proforma", "Bons_Paiement": "Bon(s) de paiement",
                "Montant": st.column_config.NumberColumn("Montant", format="%d"), "Approuve_Par": "Approuvé par"})

# ---------------------------------------------------------------- TAB 3
with tab3:
    st.subheader("⚖️ Rapport des poids")
    if df.empty and bons.empty:
        st.info("Aucune donnée.")
    else:
        r1, r2 = st.columns(2)
        noms_dispo = sorted((set(df["Fournisseur"]) | set(bons["Nom_Fournisseur"])) - {""})
        r_frs = r1.selectbox("Fournisseur", ["Tous"] + noms_dispo, key="r_frs")
        if df.empty:
            r_per = ()
        else:
            r_per = r2.date_input("Période (indicateurs, jours, fournisseurs)",
                                  value=(df["Date_Dechargement"].min().date(), df["Date_Dechargement"].max().date()),
                                  key="r_per")
        sub = df.copy()
        if r_frs != "Tous":
            sub = sub[sub["Fournisseur"] == r_frs]
        if isinstance(r_per, (tuple, list)) and len(r_per) == 2 and not sub.empty:
            sub = sub[(sub["Date_Dechargement"].dt.date >= r_per[0]) & (sub["Date_Dechargement"].dt.date <= r_per[1])]

        if not sub.empty:
            p1, p2, p3, p4, p5 = st.columns(5)
            p1.metric("Camions", dd.fmt(len(sub)))
            p2.metric("Poids net", f"{sub['Poids_Net_KG'].sum() / 1000:,.1f} T".replace(",", " "))
            p3.metric("Sacs", dd.fmt(sub["Nb_Sacs"].sum()))
            p4.metric("Poids moyen / camion", dd.fmt(sub["Poids_Net_KG"].mean(), "kg"))
            pu_moy = sub["Montant_Du"].sum() / sub["Poids_Net_KG"].sum() if sub["Poids_Net_KG"].sum() else 0
            p5.metric("PU moyen pondéré", dd.fmt(pu_moy))

        # --- Par bon de commande ---
        st.markdown("#### 📑 Suivi par bon de commande (commandé ↔ livré)")
        st.caption("Ce tableau tient compte du fournisseur choisi, pas de la période. "
                   "EN ATTENTE = bon émis, aucun camion encore déchargé.")
        suivi = dd.suivi_par_bc(df, bons)
        if r_frs != "Tous":
            suivi = suivi[suivi["Fournisseur"] == r_frs]
        statuts = st.multiselect("Statut de livraison", ["EN ATTENTE", "EN COURS", "LIVRE", "DECHARGE"],
                                 default=["EN ATTENTE", "EN COURS", "LIVRE", "DECHARGE"], key="r_stat")
        suivi = suivi[suivi["Statut_Livraison"].isin(statuts)]
        aff_suivi = suivi.copy()
        aff_suivi["Statut_Paiement"] = aff_suivi["Statut_Paiement"].map(ICONES)
        st.dataframe(aff_suivi, hide_index=True, width="stretch", height=420, column_config={
            "Reference": "N° Bon de commande", "Code_Frs": "Code",
            "Qte_Commandee_KG": st.column_config.NumberColumn("Commandé (kg)", format="%d"),
            "Poids_Livre_KG": st.column_config.NumberColumn("Livré (kg)", format="%d"),
            "Ecart_KG": st.column_config.NumberColumn("Écart (kg)", format="%d"),
            "Pct_Livre": st.column_config.NumberColumn("% livré", format="%.1f"),
            "Nb_Camions": st.column_config.NumberColumn("Camions", format="%d"),
            "Nb_Sacs": st.column_config.NumberColumn("Sacs", format="%d"),
            "Montant_Du": st.column_config.NumberColumn("Dû", format="%d"),
            "Montant_Paye": st.column_config.NumberColumn("Payé", format="%d"),
            "Solde": st.column_config.NumberColumn("Solde", format="%d"),
            "Statut_Livraison": "Livraison", "Statut_Paiement": "Paiement",
            "Derniere_Date": st.column_config.DateColumn("Dernier déchargement", format="DD/MM/YYYY")})

        # --- Par jour / par fournisseur ---
        par_jour = pd.DataFrame()
        par_frs = pd.DataFrame()
        if not sub.empty:
            par_jour = (sub.assign(Jour=sub["Date_Dechargement"].dt.date)
                        .groupby("Jour").agg(Camions=("Transaction", "count"), Poids_KG=("Poids_Net_KG", "sum"),
                                             Sacs=("Nb_Sacs", "sum"), Montant_Du=("Montant_Du", "sum"))
                        .reset_index().sort_values("Jour", ascending=False))
            par_frs = (sub.groupby("Fournisseur").agg(Camions=("Transaction", "count"), Poids_KG=("Poids_Net_KG", "sum"),
                                                      Montant_Du=("Montant_Du", "sum"), Montant_Paye=("Montant_Paye", "sum"),
                                                      Solde=("Solde", "sum")).reset_index()
                       .sort_values("Poids_KG", ascending=False))
            g1, g2 = st.columns(2)
            with g1:
                st.markdown("#### 📅 Poids par jour")
                st.bar_chart(par_jour.sort_values("Jour").set_index("Jour")["Poids_KG"] / 1000)
                st.caption("Tonnes déchargées par jour")
                st.dataframe(par_jour, hide_index=True, width="stretch", height=260)
            with g2:
                st.markdown("#### 👥 Poids par fournisseur")
                st.dataframe(par_frs, hide_index=True, width="stretch", height=380)

        st.download_button("⬇️ Télécharger le rapport complet (Excel)",
                           data=dd.to_excel_bytes({"Par bon de commande": suivi, "Par jour": par_jour,
                                                   "Par fournisseur": par_frs}),
                           file_name="rapport_poids_dechargements.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------------- TAB 4
with tab4:
    st.subheader("📥 Importer le fichier Excel « COMMANDES DECHARGEES & PAIEMENTS »")
    st.caption(f"⚠️ Les lignes seront importées dans le produit **{dd.libelle_produit(produit)}** (changez le produit dans le menu de gauche si besoin). "
               "À faire une seule fois (ou pour compléter). Les transactions déjà présentes sont ignorées : "
               "aucun doublon. Chaque « Montant Paye » devient un paiement approuvé avec ses bons de paiement.")
    if not peut_approuver:
        st.info("🔒 Import réservé aux comptes : " + ", ".join(dd.ROLES_APPROBATION))
    else:
        fichier = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"], key="up_dech")
        if fichier is not None and st.button("📥 Importer maintenant", type="primary"):
            try:
                ajoutes, ignores, nb_pay = dd.importer_excel_dechargement(fichier, name)
                flash("success", f"Import terminé : {ajoutes} déchargement(s) ajouté(s), {nb_pay} paiement(s) repris, "
                                 f"{ignores} déjà présent(s) ignoré(s).")
                st.rerun()
            except Exception as e:
                st.error(f"Erreur d'import : {e}")

st.markdown("---")
st.caption(f"NAP SARL — Déchargements {dd.libelle_produit(produit)} · fichier `{dd.fichier_actif()}` (créé automatiquement à côté de l'app)")
