# -*- coding: utf-8 -*-
"""NAP SARL — Factures Proforma (AIB/TVA optionnels, FR/EN)."""
import streamlit as st
import pandas as pd
from datetime import date
import os

import proforma_data as pf
from generer_proforma import generer_proforma_word, generer_proforma_pdf

st.set_page_config(page_title="Factures Proforma", page_icon="🧾", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

st.title("🧾 Factures Proforma — NAP SARL")
st.caption("Modèle moderne type bon de paiement · AIB et TVA en option · Impression FR / EN")

df = pf.charger()
k1, k2, k3 = st.columns(3)
k1.metric("Proformas", len(df))
k2.metric("Total TTC enregistré", f"{df['Montant_TTC'].sum():,.0f}".replace(",", " ") if not df.empty else "0")
k3.metric("Prochain n°", pf.prochain_numero())

tab_new, tab_list, tab_print = st.tabs(["➕ Nouvelle proforma", "📋 Registre", "🖨️ Générer Word FR / EN"])

# ----- NOUVELLE -----
with tab_new:
    st.subheader("Créer une facture proforma")
    with st.form("form_proforma"):
        st.markdown("#### Identification")
        c1, c2, c3 = st.columns(3)
        with c1:
            numero = st.text_input("N° Proforma", value=pf.prochain_numero())
            d_fact = st.date_input("Date", value=date.today())
        with c2:
            client = st.text_input("Client *", placeholder="ETS HAMED MOHAMED")
            adresse = st.text_input("Adresse", placeholder="EL MINA, ILOT…")
        with c3:
            ifu = st.text_input("IFU / RCCM client", placeholder="RNC 01581677")
            pays = st.text_input("Pays", placeholder="MAURITANIE")

        st.markdown("#### Produit & montants")
        p1, p2 = st.columns(2)
        with p1:
            produit = st.text_input("Produit *", placeholder="UD CAKE")
            description = st.text_area("Description", placeholder="UD CAKE", height=60)
            incoterm = st.text_input("Incoterm", value="EX-USINE")
            conditionnement = st.text_input("Conditionnement", value="SACS PP DE 50KG")
        with p2:
            qte = st.number_input("Quantité (Tonne) *", min_value=0.0, step=1.0, value=0.0)
            pu = st.number_input("Prix unitaire (FCFA) *", min_value=0.0, step=1000.0, value=0.0)

        st.markdown("#### Taxes (optionnel)")
        t1, t2, t3, t4 = st.columns(4)
        with t1:
            app_aib = st.checkbox("Appliquer AIB", value=False)
        with t2:
            taux_aib = st.number_input("Taux AIB %", min_value=0.0, max_value=100.0, value=1.0, step=0.1,
                                       disabled=not app_aib)
        with t3:
            app_tva = st.checkbox("Appliquer TVA", value=False)
        with t4:
            taux_tva = st.number_input("Taux TVA %", min_value=0.0, max_value=100.0, value=18.0, step=0.1,
                                       disabled=not app_tva)

        ht, aib, tva, ttc = pf.calculer(qte, pu, app_aib, taux_aib, app_tva, taux_tva)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("HT", f"{ht:,.0f}".replace(",", " "))
        m2.metric(f"AIB ({taux_aib}%)" if app_aib else "AIB", f"{aib:,.0f}".replace(",", " "))
        m3.metric(f"TVA ({taux_tva}%)" if app_tva else "TVA", f"{tva:,.0f}".replace(",", " "))
        m4.metric("TTC", f"{ttc:,.0f}".replace(",", " "))

        # Montant en lettres généré automatiquement à l'enregistrement / impression
        lettres_auto = pf.montant_en_lettres(ttc, lang="FR")
        st.info(f"**Montant en lettres (auto) :** {lettres_auto}")
        lettres = ""  # auto à l'enregistrement

        st.markdown("#### Conditions")
        cond_pay = st.selectbox("Condition de paiement", [""] + pf.CONDITIONS_PAIEMENT)
        if not cond_pay:
            cond_pay = st.text_input("Condition (saisie libre)", value="Paiement de 100 % d'avance en 7 jours.")
        validite = st.text_input("Validité de l'offre", value="Période de chargement")
        livraison = st.text_input("Livraison", placeholder="DEC 2026 – FEV 2027")
        mode_pay = st.selectbox("Mode de paiement", pf.MODES_PAIEMENT)
        compte = st.text_area("Compte bancaire", value=pf.COMPTE_DEFAUT, height=70)
        autres = st.text_area("Autres conditions", value=pf.AUTRES_DEFAUT, height=100)

        ok = st.form_submit_button("💾 Enregistrer la proforma", type="primary")

    if ok:
        if not client.strip() or not produit.strip():
            st.error("Client et produit sont obligatoires.")
        elif qte <= 0 or pu <= 0:
            st.error("Quantité et prix doivent être > 0.")
        else:
            success, msg, eid = pf.ajouter({
                "Numero": numero.strip(),
                "Date_Facture": d_fact,
                "Client": client.strip(),
                "Adresse": adresse.strip(),
                "IFU_Client": ifu.strip(),
                "Pays": pays.strip(),
                "Produit": produit.strip(),
                "Description": description.strip(),
                "Incoterm": incoterm.strip(),
                "Conditionnement": conditionnement.strip(),
                "Quantite_T": qte,
                "Prix_Unitaire": pu,
                "Appliquer_AIB": app_aib,
                "Taux_AIB": taux_aib,
                "Appliquer_TVA": app_tva,
                "Taux_TVA": taux_tva,
                "Montant_Lettres": lettres.strip(),
                "Condition_Paiement": cond_pay,
                "Validite": validite,
                "Livraison": livraison,
                "Mode_Paiement": mode_pay,
                "Compte_Bancaire": compte,
                "Autres_Conditions": autres,
            }, user=name)
            if success:
                st.success(msg)
                st.session_state["proforma_sel"] = eid
                st.rerun()
            else:
                st.error(msg)

# ----- LISTE -----
with tab_list:
    st.subheader("Registre des proformas")
    if df.empty:
        st.info("Aucune proforma.")
    else:
        aff = df.copy()
        aff["Date_Facture"] = pd.to_datetime(aff["Date_Facture"], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(
            aff[[
                "ID", "Numero", "Date_Facture", "Client", "Produit", "Quantite_T",
                "Prix_Unitaire", "Montant_HT", "Appliquer_AIB", "Montant_AIB",
                "Appliquer_TVA", "Montant_TVA", "Montant_TTC", "Statut",
            ]],
            hide_index=True, width="stretch", height=400,
        )
        st.download_button(
            "⬇️ Export Excel",
            data=pf.to_excel_bytes(df),
            file_name=f"Proformas_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ----- PRINT -----
with tab_print:
    st.subheader("Générer Word — Français / Anglais")
    if df.empty:
        st.info("Enregistrez d'abord une proforma.")
    else:
        opts = {
            f"#{int(r.ID)} · {r.Numero} · {r.Client} · {r.Produit} · "
            f"{float(r.Montant_TTC):,.0f} FCFA".replace(",", " "): int(r.ID)
            for _, r in df.iterrows()
        }
        choix = st.selectbox("Proforma", list(opts.keys()))
        eid = opts[choix]
        row = df[df["ID"] == eid].iloc[0]
        data = row.to_dict()

        st.write(
            f"**{data.get('Numero')}** — {data.get('Client')} — "
            f"TTC **{float(data.get('Montant_TTC') or 0):,.0f} FCFA**".replace(",", " ")
        )
        aib_on = str(data.get("Appliquer_AIB")).upper() in ("OUI", "YES", "1")
        tva_on = str(data.get("Appliquer_TVA")).upper() in ("OUI", "YES", "1")
        st.caption(
            f"AIB: {'Oui ' + str(data.get('Taux_AIB')) + '%' if aib_on else 'Non'} · "
            f"TVA: {'Oui ' + str(data.get('Taux_TVA')) + '%' if tva_on else 'Non'}"
        )

        lang = st.radio("Langue du document", ["FR", "EN"], horizontal=True,
                        format_func=lambda x: "Français" if x == "FR" else "English")

        b1, b2 = st.columns(2)
        with b1:
            if st.button("📄 Générer Word", type="primary", use_container_width=True):
                try:
                    chemin = generer_proforma_word(data, lang=lang)
                    st.success(f"Word généré : `{chemin}`")
                    with open(chemin, "rb") as f:
                        st.download_button(
                            f"⬇️ Télécharger Word ({lang})",
                            data=f.read(),
                            file_name=os.path.basename(chemin),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary",
                        )
                except Exception as e:
                    st.error(f"Erreur Word : {e}")
                    import traceback
                    st.code(traceback.format_exc())
        with b2:
            if st.button("📕 Générer PDF", type="primary", use_container_width=True):
                try:
                    pdf = generer_proforma_pdf(data, lang=lang)
                    safe = str(data.get("Numero") or eid).replace("/", "-")
                    st.success("PDF prêt.")
                    st.download_button(
                        f"⬇️ Télécharger PDF ({lang})",
                        data=pdf,
                        file_name=f"Proforma_{safe}_{lang}.pdf",
                        mime="application/pdf",
                        type="primary",
                    )
                except Exception as e:
                    st.error(f"Erreur PDF : {e}")
                    import traceback
                    st.code(traceback.format_exc())

st.caption("STE NAP SARL — Proforma · AIB optionnel (taux modifiable) · TVA optionnelle · FR/EN")
