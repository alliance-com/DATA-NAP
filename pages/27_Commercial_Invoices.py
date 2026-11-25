# -*- coding: utf-8 -*-
"""
Export — Commercial Invoices NAP SARL
Sous-dossier export · génération Word type modèle commercial.
"""
import streamlit as st
import pandas as pd
from datetime import date, datetime
import sys
import os

_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from export import commercial_invoice_data as ci
from export.generer_commercial_invoice import generer_commercial_invoice
from export.ci_amount_words import amount_in_words

st.set_page_config(page_title="Commercial Invoices", page_icon="🧾", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "User"

st.title("🧾 Commercial Invoices — NAP SARL")
st.caption(
    "Registre + génération Word au format commercial invoice · "
    "montants et libellés en **gras automatiques** · n° NAP/xxxx/année"
)

try:
    df = ci.charger()
except Exception as e:
    st.error(str(e))
    st.stop()

k1, k2, k3, k4 = st.columns(4)
k1.metric("Invoices", len(df))
k2.metric("Buyers", df["Buyer"].nunique() if not df.empty else 0)
if not df.empty:
    eur = df[df["Currency"] == "EUR"]["Net_Amount"].sum()
    usd = df[df["Currency"] == "USD"]["Net_Amount"].sum()
else:
    eur = usd = 0.0
k3.metric("Total net EUR", f"{eur:,.2f}")
k4.metric("Total net USD", f"{usd:,.2f}")

tab_new, tab_list, tab_print = st.tabs([
    "➕ Nouvelle invoice",
    "📋 Registre",
    "🖨️ Générer Word",
])

# ---------- NOUVELLE ----------
with tab_new:
    st.subheader("Créer une Commercial Invoice")
    with st.form("form_ci"):
        st.markdown("#### Identification")
        a1, a2, a3 = st.columns(3)
        with a1:
            inv_no = st.text_input("Invoice No", value=ci.prochain_numero())
            d_inv = st.date_input("Date", value=date.today())
        with a2:
            buyers = ci.liste_buyers()
            buyer_sel = st.selectbox("Buyer (liste)", ["— Saisie libre —"] + buyers)
            buyer = st.text_input(
                "Buyer *",
                value="" if buyer_sel == "— Saisie libre —" else buyer_sel,
                placeholder="BUNGE LODERS CROKLAAN B.V",
            )
        with a3:
            buyer_addr = st.text_area(
                "Adresse buyer",
                height=90,
                placeholder="HOGEWEG 1, 1521 AZ WORMERVEER\nTHE NETHERLANDS",
            )

        st.markdown("#### Produit & logistique")
        b1, b2, b3 = st.columns(3)
        with b1:
            product = st.selectbox(
                "Produit",
                ["CRUDE SHEA BUTTER", "SHEA NUT MEAL", "Autre"],
            )
            if product == "Autre":
                product = st.text_input("Libellé produit", value="")
            nb_ctr = st.number_input("Nb conteneurs", min_value=0, step=1, value=5)
            ctr_type = st.text_input("Type conteneur", value="20'")
        with b2:
            contract = st.text_input("Contract No", placeholder="4600072071")
            bl = st.text_input("Bill of Lading", placeholder="264169298")
            liner = st.text_input("Shipping Line", value="MAERSK LINE")
        with b3:
            incoterm = st.text_input("Incoterm / Prices are", value="FOB COTONOU PORT")
            origin = st.text_input("Origin", value="COTONOU BENIN")
            dest = st.text_input("Final destination", placeholder="TANJUNG PELEPAS")

        st.markdown("#### Montants")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            qty = st.number_input("Quantity (MT) *", min_value=0.0, step=0.001, format="%.3f")
            price = st.number_input("Price per MT *", min_value=0.0, step=1.0, format="%.2f")
        with c2:
            currency = st.selectbox("Currency", ci.DEVISES)
            advance = st.number_input("Less Advance", min_value=0.0, step=100.0, format="%.2f")
        with c3:
            pi_pct = st.number_input("Less PI %", min_value=0.0, max_value=100.0, step=1.0, value=0.0)
        with c4:
            gross, adv, pi_amt, net = ci.calculer_montants(qty, price, advance, pi_pct)
            st.metric("Gross", f"{gross:,.2f}")
            st.metric("Net due", f"{net:,.2f} {currency}")

        words_auto = amount_in_words(net, currency)
        st.info(f"**IN WORDS (auto) :** {words_auto}")
        words_manual = st.text_input("IN WORDS (laisser vide = automatique)", value="")

        desc_auto = ci.construire_description({
            "Product": product,
            "Nb_Containers": nb_ctr,
            "Contract_No": contract,
            "Incoterm": incoterm,
            "BL_No": bl,
            "Shipping_Line": liner,
            "Origin": origin,
            "Final_Destination": dest,
        })
        description = st.text_area("Description (auto modifiable)", value=desc_auto, height=140)
        bank = st.text_area("Bank details", value=ci.BANK_DEFAULT, height=60)
        notes = st.text_input("Notes internes")
        statut = st.selectbox("Statut", ci.STATUTS, index=1)

        ok = st.form_submit_button("💾 Enregistrer l'invoice", type="primary")

    if ok:
        if not buyer.strip():
            st.error("Buyer obligatoire.")
        elif qty <= 0 or price <= 0:
            st.error("Quantité et prix > 0.")
        else:
            ok2, msg, eid = ci.ajouter({
                "Invoice_No": inv_no.strip(),
                "Date_Invoice": d_inv,
                "Buyer": buyer.strip(),
                "Buyer_Address": buyer_addr.strip(),
                "Description": description.strip(),
                "Product": product,
                "Nb_Containers": nb_ctr,
                "Container_Type": ctr_type,
                "Contract_No": contract,
                "Incoterm": incoterm,
                "BL_No": bl,
                "Shipping_Line": liner,
                "Origin": origin,
                "Final_Destination": dest,
                "Qty_MT": qty,
                "Price_Per_MT": price,
                "Currency": currency,
                "Advance": advance,
                "Less_PI_Pct": pi_pct,
                "Amount_Words": words_manual.strip(),
                "Bank_Details": bank,
                "Notes": notes,
                "Statut": statut,
            }, user=name)
            if ok2:
                st.success(msg)
                st.session_state["ci_sel"] = eid
                st.rerun()
            else:
                st.error(msg)

# ---------- REGISTRE ----------
with tab_list:
    st.subheader("Registre des Commercial Invoices")
    if df.empty:
        st.info("Aucune invoice.")
    else:
        f1, f2, f3 = st.columns(3)
        with f1:
            buyers_f = ["Tous"] + ci.liste_buyers()
            bf = st.selectbox("Buyer", buyers_f)
        with f2:
            cf = st.selectbox("Currency", ["Toutes"] + ci.DEVISES)
        with f3:
            sf = st.selectbox("Statut", ["Tous"] + ci.STATUTS)
        vue = df.copy()
        if bf != "Tous":
            vue = vue[vue["Buyer"] == bf]
        if cf != "Toutes":
            vue = vue[vue["Currency"] == cf]
        if sf != "Tous":
            vue = vue[vue["Statut"] == sf]
        aff = vue.copy()
        aff["Date_Invoice"] = pd.to_datetime(aff["Date_Invoice"], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(
            aff[[
                "ID", "Invoice_No", "Date_Invoice", "Buyer", "Product", "BL_No",
                "Qty_MT", "Price_Per_MT", "Currency", "Gross_Amount", "Advance",
                "Net_Amount", "Statut",
            ]],
            hide_index=True, width="stretch", height=420,
        )
        st.download_button(
            "⬇️ Export Excel",
            data=ci.to_excel_bytes(vue),
            file_name=f"Commercial_Invoices_{date.today():%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ---------- PRINT ----------
with tab_print:
    st.subheader("Générer le Word Commercial Invoice")
    st.caption(
        "Le document suit le format de votre modèle (en-tête NAP, tableau DESCRIPTION / QTY / PRICE / AMOUNT, "
        "IN WORDS, bank details, clauses). Les valeurs variables sont en **gras**."
    )
    if df.empty:
        st.info("Enregistrez d'abord une invoice.")
    else:
        opts = {
            f"{r.Invoice_No} · {r.Buyer} · {r.BL_No} · {float(r.Net_Amount):,.2f} {r.Currency}": int(r.ID)
            for _, r in df.iterrows()
        }
        choix = st.selectbox("Invoice", list(opts.keys()))
        eid = opts[choix]
        row = df[df["ID"] == eid].iloc[0]
        data = row.to_dict()

        c1, c2, c3 = st.columns(3)
        c1.write(f"**{data.get('Invoice_No')}**")
        c2.write(f"{data.get('Buyer')}")
        c3.write(f"Net **{float(data.get('Net_Amount') or 0):,.2f} {data.get('Currency')}**")

        if st.button("📄 Générer Word", type="primary", use_container_width=True):
            try:
                chemin = generer_commercial_invoice(data)
                st.success(f"Généré : `{chemin}`")
                with open(chemin, "rb") as f:
                    st.download_button(
                        "⬇️ Télécharger Word",
                        data=f.read(),
                        file_name=os.path.basename(chemin),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                    )
            except Exception as e:
                st.error(str(e))
                import traceback
                st.code(traceback.format_exc())

st.caption("STE NAP SARL — Commercial Invoice · export/")
