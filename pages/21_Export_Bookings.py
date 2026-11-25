# -*- coding: utf-8 -*-
"""Export — Contrats & Bookings / BL (+ modification ETD, ETA, statuts…)"""
import streamlit as st
import pandas as pd
from datetime import date, datetime
import sys
import os

_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
from export import export_data as ex

st.set_page_config(page_title="Export Bookings", page_icon="📄", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "User"

st.title("📄 Export — Contrats & Bookings")

try:
    _ = ex.charger_bookings()
except FileNotFoundError as e:
    st.error(str(e))
    st.stop()


def _to_date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, date) and not isinstance(v, datetime):
        return v
    ts = pd.to_datetime(v, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.date()


tab_c, tab_b, tab_edit = st.tabs(["📑 Contrats", "📦 Bookings / BL", "✏️ Modifier un booking"])

# ===== CONTRATS =====
with tab_c:
    st.subheader("Nouveau contrat")
    with st.form("form_contract"):
        c1, c2, c3 = st.columns(3)
        with c1:
            contract_no = st.text_input("N° Contrat *")
            d_contract = st.date_input("Date contrat", value=date.today())
            commodity = st.selectbox("Commodity *", [""] + ex.liste_commodities())
        with c2:
            buyers = ex.liste_buyers()
            buyer = st.selectbox("Buyer *", [""] + buyers)
            if buyer == "" and st.checkbox("Nouveau buyer (saisie libre)"):
                buyer = st.text_input("Nom buyer")
            qty = st.number_input("Qty contrat (MT)", min_value=0.0, step=1.0)
            price = st.number_input("Prix unitaire", min_value=0.0, step=1.0)
        with c3:
            currency = st.selectbox("Devise", ["EUR", "USD", "XOF"])
            ports = ex.liste_ports()
            pod = st.selectbox("POD", [""] + ports)
            if pod == "":
                pod = st.text_input("POD (saisie)")
            country = st.selectbox("Country", [""] + ex.liste_countries())
            if country == "":
                country = st.text_input("Country (saisie)")
            notes = st.text_input("Notes")
        ok_c = st.form_submit_button("💾 Enregistrer le contrat", type="primary")
    if ok_c:
        if not contract_no.strip() or not commodity or not buyer:
            st.error("Contrat, commodity et buyer sont obligatoires.")
        else:
            ok, msg = ex.ajouter_contract({
                "Contract_No": contract_no.strip(),
                "Date_Contract": d_contract,
                "Buyer": buyer.strip() if isinstance(buyer, str) else buyer,
                "Commodity": commodity,
                "Qty_Contract_MT": qty,
                "Price_Unit": price,
                "Currency": currency,
                "POD": pod,
                "Country": country,
                "Notes": notes,
                "Actif": "OUI",
            })
            st.success(msg) if ok else st.error(msg)
            if ok:
                st.rerun()

    st.markdown("---")
    df_c = ex.charger_contracts()
    if df_c.empty:
        st.info("Aucun contrat.")
    else:
        aff = df_c.copy()
        aff["Date_Contract"] = pd.to_datetime(aff["Date_Contract"], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(aff, hide_index=True, width="stretch")

# ===== BOOKINGS =====
with tab_b:
    st.sidebar.header("➕ Nouveau Booking / BL")
    contracts = ex.charger_contracts()
    contract_opts = [""] + (contracts["Contract_No"].astype(str).tolist() if not contracts.empty else [])

    with st.sidebar.form("form_booking"):
        booking_no = st.text_input("Booking No *")
        bl_no = st.text_input("BL No (si vide = Booking)")
        contract_no = st.selectbox("Contrat", contract_opts)
        commodity = st.selectbox("Commodity *", [""] + ex.liste_commodities(), key="nb_comm")
        buyer = st.selectbox("Buyer *", [""] + ex.liste_buyers(), key="nb_buyer")
        if not buyer:
            buyer = st.text_input("Buyer (saisie libre)")
        liner = st.selectbox("Liner", [""] + ex.liste_liners())
        forwarder = st.selectbox("Forwarder *", [""] + ex.liste_forwarders())
        pod = st.selectbox("POD", [""] + ex.liste_ports(), key="nb_pod")
        country = st.selectbox("Country", [""] + ex.liste_countries(), key="nb_cty")
        ctr_booked = st.number_input("Conteneurs booked", min_value=0, step=1, value=0)
        ctr_loaded = st.number_input("Conteneurs loaded", min_value=0, step=1, value=0)
        not_alloc = st.number_input("Not yet allocated", min_value=0, step=1, value=0)
        net_qty = st.number_input("Net Qty (KG)", min_value=0.0, step=100.0)
        gross_qty = st.number_input("Gross Qty (KG)", min_value=0.0, step=100.0)
        nb_bags = st.number_input("Nb bags", min_value=0, step=1)
        tank = st.selectbox("Storage tank", [""] + ex.liste_tanks())
        rotation = st.number_input("Rotation", min_value=0, step=1, value=1)
        stuffing = st.selectbox("Stuffing status", ex.STUFFING_STATUSES)
        vessel = st.text_input("Vessel")
        voy = st.text_input("VOY")
        etd = st.date_input("ETD", value=None, key="nb_etd")
        eta = st.date_input("ETA", value=None, key="nb_eta")
        shipping = st.selectbox("Shipping status", ex.SHIPPING_STATUSES)
        notes = st.text_input("Notes")
        ok_b = st.form_submit_button("💾 Enregistrer Booking", type="primary")

    if ok_b:
        if not booking_no.strip() or not commodity or not forwarder:
            st.sidebar.error("Booking, commodity et forwarder obligatoires.")
        else:
            pod_f, country_f, buyer_f = pod, country, buyer
            if contract_no and not contracts.empty:
                row = contracts[contracts["Contract_No"].astype(str) == str(contract_no)]
                if not row.empty:
                    r0 = row.iloc[0]
                    pod_f = pod_f or r0.get("POD") or ""
                    country_f = country_f or r0.get("Country") or ""
                    buyer_f = buyer_f or r0.get("Buyer") or ""
                    if not commodity:
                        commodity = r0.get("Commodity") or commodity
            ok, msg = ex.ajouter_booking({
                "Booking_No": booking_no.strip(),
                "BL_No": bl_no.strip(),
                "Contract_No": contract_no,
                "Commodity": commodity,
                "Buyer": buyer_f,
                "POD": pod_f,
                "Country": country_f,
                "Liner": liner,
                "Forwarder": forwarder,
                "Ctr_Booked": ctr_booked,
                "Ctr_Loaded": ctr_loaded,
                "Not_Allocated": not_alloc,
                "Net_Qty_KG": net_qty,
                "Gross_Qty_KG": gross_qty,
                "Nb_Bags": nb_bags,
                "Storage_Tank": tank,
                "Rotation": rotation,
                "Stuffing_Status": stuffing,
                "Vessel": vessel,
                "VOY": voy,
                "ETD": etd,
                "ETA": eta,
                "Shipping_Status": shipping,
                "Notes": notes,
            }, user=name)
            if ok:
                st.sidebar.success(msg)
                st.rerun()
            else:
                st.sidebar.error(msg)

    st.subheader("Liste des bookings")
    df = ex.charger_bookings()
    if df.empty:
        st.info("Aucun booking enregistré.")
    else:
        aff = df.copy()
        for c in ["ETD", "ETA", "Date_Creation"]:
            aff[c] = pd.to_datetime(aff[c], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(aff, hide_index=True, width="stretch", height=420)

# ===== MODIFICATION =====
with tab_edit:
    st.subheader("✏️ Modifier un booking existant")
    st.caption(
        "Idéal pour mettre à jour **ETD**, **ETA**, vessel, statuts, quantités, conteneurs loaded, etc. "
        "sans ressaisir tout le dossier."
    )
    df = ex.charger_bookings()
    if df.empty:
        st.info("Aucun booking à modifier.")
    else:
        labels = {
            f"{r.Booking_No}  |  BL {r.BL_No}  |  {r.Commodity}  |  {r.Forwarder}  |  "
            f"ETD {_to_date(r.ETD) or '—'}  |  {r.Stuffing_Status}/{r.Shipping_Status}": r.Booking_No
            for _, r in df.iterrows()
        }
        choix = st.selectbox("Sélectionner le booking", list(labels.keys()))
        bno = labels[choix]
        row = df[df["Booking_No"].astype(str) == str(bno)].iloc[0]

        st.markdown(f"### Booking `{bno}`")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Ctr booked", int(row.get("Ctr_Booked") or 0))
        m2.metric("Ctr loaded", int(row.get("Ctr_Loaded") or 0))
        m3.metric("Balance", int(row.get("Balance") or 0))
        m4.metric("Stuffing", row.get("Stuffing_Status") or "—")

        with st.form("form_edit_booking"):
            st.markdown("#### Identité & acteurs")
            e1, e2, e3 = st.columns(3)
            with e1:
                bl_no = st.text_input("BL No", value=str(row.get("BL_No") or ""))
                contract_no = st.text_input("Contrat", value=str(row.get("Contract_No") or ""))
                commodity = st.selectbox(
                    "Commodity",
                    options=ex.liste_commodities() or [str(row.get("Commodity") or "")],
                    index=max(0, (ex.liste_commodities() or [str(row.get("Commodity") or "")]).index(str(row.get("Commodity") or "")) if str(row.get("Commodity") or "") in (ex.liste_commodities() or []) else 0),
                )
            with e2:
                buyers = ex.liste_buyers()
                b0 = str(row.get("Buyer") or "")
                buyer = st.selectbox("Buyer", options=([b0] + [b for b in buyers if b != b0]) if b0 else buyers)
                fwds = ex.liste_forwarders()
                f0 = str(row.get("Forwarder") or "")
                forwarder = st.selectbox("Forwarder", options=([f0] + [f for f in fwds if f != f0]) if f0 else fwds)
                liners = ex.liste_liners()
                l0 = str(row.get("Liner") or "")
                liner = st.selectbox("Liner", options=([l0] + [l for l in liners if l != l0]) if l0 else ([""] + liners))
            with e3:
                pod = st.text_input("POD", value=str(row.get("POD") or ""))
                country = st.text_input("Country", value=str(row.get("Country") or ""))
                tank = st.selectbox(
                    "Storage tank",
                    options=[""] + ex.liste_tanks(),
                    index=([""] + ex.liste_tanks()).index(str(row.get("Storage_Tank") or "")) if str(row.get("Storage_Tank") or "") in ([""] + ex.liste_tanks()) else 0,
                )

            st.markdown("#### Conteneurs & quantités")
            q1, q2, q3, q4 = st.columns(4)
            with q1:
                ctr_booked = st.number_input("Ctr booked", min_value=0, step=1, value=int(row.get("Ctr_Booked") or 0))
            with q2:
                ctr_loaded = st.number_input("Ctr loaded", min_value=0, step=1, value=int(row.get("Ctr_Loaded") or 0))
            with q3:
                not_alloc = st.number_input("Not allocated", min_value=0, step=1, value=int(row.get("Not_Allocated") or 0))
            with q4:
                nb_bags = st.number_input("Nb bags", min_value=0, step=1, value=int(row.get("Nb_Bags") or 0))
            q5, q6, q7 = st.columns(3)
            with q5:
                net_qty = st.number_input("Net Qty (KG)", min_value=0.0, step=100.0, value=float(row.get("Net_Qty_KG") or 0))
            with q6:
                gross_qty = st.number_input("Gross Qty (KG)", min_value=0.0, step=100.0, value=float(row.get("Gross_Qty_KG") or 0))
            with q7:
                rotation = st.number_input("Rotation", min_value=0, step=1, value=int(row.get("Rotation") or 0) if str(row.get("Rotation") or "").isdigit() or isinstance(row.get("Rotation"), (int, float)) else 0)

            st.markdown("#### Navire & dates (souvent à modifier)")
            d1, d2, d3, d4 = st.columns(4)
            with d1:
                vessel = st.text_input("Vessel", value=str(row.get("Vessel") or ""))
            with d2:
                voy = st.text_input("VOY", value=str(row.get("VOY") or ""))
            with d3:
                etd_val = _to_date(row.get("ETD"))
                etd = st.date_input("ETD", value=etd_val)
            with d4:
                eta_val = _to_date(row.get("ETA"))
                eta = st.date_input("ETA", value=eta_val)

            st.markdown("#### Statuts")
            s1, s2 = st.columns(2)
            with s1:
                st0 = str(row.get("Stuffing_Status") or "PENDING")
                stuffing = st.selectbox(
                    "Stuffing status",
                    ex.STUFFING_STATUSES,
                    index=ex.STUFFING_STATUSES.index(st0) if st0 in ex.STUFFING_STATUSES else 0,
                )
            with s2:
                sh0 = str(row.get("Shipping_Status") or "PENDING")
                shipping = st.selectbox(
                    "Shipping status",
                    ex.SHIPPING_STATUSES,
                    index=ex.SHIPPING_STATUSES.index(sh0) if sh0 in ex.SHIPPING_STATUSES else 0,
                )
            notes = st.text_area("Notes", value=str(row.get("Notes") or ""), height=70)

            save = st.form_submit_button("💾 Enregistrer les modifications", type="primary", use_container_width=True)

        if save:
            updates = {
                "BL_No": bl_no.strip(),
                "Contract_No": contract_no.strip(),
                "Commodity": commodity,
                "Buyer": buyer,
                "Forwarder": forwarder,
                "Liner": liner,
                "POD": pod.strip(),
                "Country": country.strip(),
                "Storage_Tank": tank,
                "Ctr_Booked": ctr_booked,
                "Ctr_Loaded": ctr_loaded,
                "Not_Allocated": not_alloc,
                "Nb_Bags": nb_bags,
                "Net_Qty_KG": net_qty,
                "Gross_Qty_KG": gross_qty,
                "Rotation": rotation,
                "Vessel": vessel.strip(),
                "VOY": voy.strip(),
                "ETD": etd,
                "ETA": eta,
                "Stuffing_Status": stuffing,
                "Shipping_Status": shipping,
                "Notes": notes,
            }
            ok, msg = ex.maj_booking(bno, updates)
            if ok:
                st.success(msg + " — ETD / ETA / statuts mis à jour.")
                st.rerun()
            else:
                st.error(msg)

        # Raccourcis rapides hors formulaire
        st.markdown("---")
        st.markdown("#### Raccourcis rapides")
        r1, r2, r3, r4 = st.columns(4)
        with r1:
            if st.button("🟠 Stuffing → IN PROGRESS", use_container_width=True):
                ok, msg = ex.maj_booking(bno, {"Stuffing_Status": "IN PROGRESS"})
                st.success(msg) if ok else st.error(msg)
                if ok:
                    st.rerun()
        with r2:
            if st.button("🟢 Stuffing → COMPLETED", use_container_width=True):
                ok, msg = ex.maj_booking(bno, {"Stuffing_Status": "COMPLETED"})
                st.success(msg) if ok else st.error(msg)
                if ok:
                    st.rerun()
        with r3:
            if st.button("🚢 Shipping → ON VESSEL", use_container_width=True):
                ok, msg = ex.maj_booking(bno, {"Shipping_Status": "BOOKED ON VESSEL"})
                st.success(msg) if ok else st.error(msg)
                if ok:
                    st.rerun()
        with r4:
            if st.button("✅ Shipping → SHIPPED", use_container_width=True):
                ok, msg = ex.maj_booking(bno, {"Shipping_Status": "SHIPPED"})
                st.success(msg) if ok else st.error(msg)
                if ok:
                    st.rerun()

st.caption("NAP SARL — Export Bookings · modification ETD / ETA / vessel / statuts")
