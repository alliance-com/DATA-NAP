# -*- coding: utf-8 -*-
"""
Page Streamlit — Packing List (export/)
Placez ce fichier dans : DATA/pages/40_Packing_List.py
Modules : DATA/export/packing_list_data.py · generer_packing_list.py · export_data.py
"""
from __future__ import annotations

import os
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

# Racine projet (parent de pages/)
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

# Package export
try:
    from export import packing_list_data as pld
    from export import generer_packing_list as gpl
except ImportError as e:
    pld = None
    gpl = None
    _import_err = str(e)

try:
    from export import export_data as ed
except ImportError:
    try:
        import export_data as ed
    except ImportError:
        ed = None

st.set_page_config(page_title="Packing List", page_icon="📦", layout="wide")

if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"

if pld is None:
    st.error(
        "Module **export.packing_list_data** introuvable.\n\n"
        "Vérifiez la structure :\n"
        "```\nDATA/\n  export/\n    __init__.py\n    packing_list_data.py\n    generer_packing_list.py\n    export_data.py\n  pages/\n    40_Packing_List.py\n```\n"
        f"Détail : {_import_err if '_import_err' in dir() else ''}"
    )
    st.stop()

st.title("📦 Packing List — Export")
st.caption("Données depuis **export/export_data** (Bookings + Containers) · Word modèle NAP")

tab_new, tab_reg = st.tabs(["➕ Nouvelle Packing List", "📚 Registre"])


def _ctr_editor(key_prefix, initial):
    """Éditeur conteneurs : prérempli depuis le booking (tous les ctr liés)."""
    initial = list(initial or [])
    # Si on vient de charger un booking, forcer le nombre = len(initial)
    default_n = len(initial) if initial else 1
    # Clé stable liée au contenu pour forcer refresh après préremplissage
    sig = "|".join(
        f"{c.get('Container_No','')}:{c.get('Seal_No','')}" for c in initial
    )[:80]
    n = st.number_input(
        "Nombre de conteneurs",
        min_value=0, max_value=50,
        value=default_n,
        key=f"{key_prefix}_n_{sig}" if sig else f"{key_prefix}_n",
        help="Chargé automatiquement depuis les conteneurs du booking sélectionné.",
    )
    if initial:
        # Aperçu tableau
        st.dataframe(
            pd.DataFrame(initial)[
                [c for c in ["S_No", "Container_No", "Seal_No", "Tan_No", "Rotation", "Bags", "Gross_KG", "Net_KG"]
                 if c in (initial[0].keys() if initial else [])]
            ],
            hide_index=True,
            width="stretch",
            height=min(280, 40 + 28 * len(initial)),
        )
        st.caption(
            f"**{len(initial)}** conteneur(s) liés au booking — poids repris de T_Containers. "
            "**Gross WT ≠ Net WT** : vérifiez chaque ligne avant génération."
        )
        # Totaux poids pour contrôle
        tg = sum(float(c.get("Gross_KG") or 0) for c in initial)
        tn = sum(float(c.get("Net_KG") or 0) for c in initial)
        m1, m2, m3 = st.columns(3)
        m1.metric("Conteneurs", len(initial))
        m2.metric("Σ Gross WT (kg)", f"{tg:,.0f}".replace(",", " "))
        m3.metric("Σ Net WT (kg)", f"{tn:,.0f}".replace(",", " "))
        if tg > 0 and tn > 0 and abs(tg - tn) < 0.5:
            st.warning(
                "⚠️ Gross et Net sont identiques sur toutes les lignes. "
                "Si votre base n’a qu’une colonne Weight_KG, renseignez le **brut** (Gross) manuellement "
                "ou ajoutez les colonnes Gross_KG / Net_KG dans T_Containers."
            )
        elif tn > 0 and tg <= 0:
            st.info("Net renseigné — Gross à compléter (ou ratio booking appliqué si Net_Qty/Gross_Qty présents).")

    rows = []
    for i in range(int(n)):
        base = initial[i] if i < len(initial) else {}
        with st.expander(
            f"Conteneur {i+1} — {base.get('Container_No') or 'à saisir'}",
            expanded=(i < 3) or (not base.get("Container_No")),
        ):
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                cno = st.text_input("Container No", value=str(base.get("Container_No") or ""), key=f"{key_prefix}_cno_{sig}_{i}")
                seal = st.text_input("Seal No", value=str(base.get("Seal_No") or ""), key=f"{key_prefix}_seal_{sig}_{i}")
            with c2:
                tan = st.text_input("Tan No / Tank", value=str(base.get("Tan_No") or ""), key=f"{key_prefix}_tan_{sig}_{i}")
                rot = st.text_input("Rotation", value=str(base.get("Rotation") or ""), key=f"{key_prefix}_rot_{sig}_{i}")
            with c3:
                bags = st.number_input("Bags", min_value=0, value=int(base.get("Bags") or 0), key=f"{key_prefix}_bags_{sig}_{i}")
                gross = st.number_input("Gross kg", min_value=0.0, value=float(base.get("Gross_KG") or 0), step=10.0, key=f"{key_prefix}_g_{sig}_{i}")
            with c4:
                net = st.number_input("Net kg", min_value=0.0, value=float(base.get("Net_KG") or 0), step=10.0, key=f"{key_prefix}_n_{sig}_{i}")
        rows.append({
            "S_No": i + 1,
            "Container_No": cno.strip(),
            "Seal_No": seal.strip(),
            "Tan_No": tan.strip(),
            "Rotation": rot.strip(),
            "Bags": int(bags),
            "Gross_KG": float(gross),
            "Net_KG": float(net),
        })
    return rows


with tab_new:
    st.subheader("1. Charger depuis un Booking")
    bookings, booking_map = [], {}
    if ed is not None:
        try:
            bdf = ed.charger_bookings()
            if not bdf.empty:
                for r in bdf.itertuples():
                    lab = f"{r.Booking_No} · BL {r.BL_No or '—'} · {r.Buyer} · {r.Commodity}"
                    bookings.append(lab)
                    booking_map[lab] = r.Booking_No
        except Exception as e:
            st.warning(f"Lecture bookings : {e}")
    else:
        st.info("export_data non disponible — saisie manuelle uniquement.")

    col_b1, col_b2 = st.columns([3, 1])
    with col_b1:
        choix_b = st.selectbox("Booking source", ["— Saisie manuelle —"] + bookings, key="pl_booking_sel")
    with col_b2:
        st.write("")
        st.write("")
        load_btn = st.button("📥 Préremplir depuis Booking", type="primary")

    if load_btn and choix_b != "— Saisie manuelle —":
        bno = booking_map.get(choix_b)
        data0, err = pld.prefill_from_booking(bno)
        if data0 is None:
            st.error(err or "Échec chargement.")
        else:
            ctrs_loaded = data0.pop("containers", []) or []
            n_booked = data0.pop("_ctr_booked", 0)
            n_loaded = data0.pop("_ctr_loaded", 0)
            data0.pop("_nb_containers", None)
            data0.pop("_net_booking_kg", None)
            data0.pop("_gross_booking_kg", None)
            st.session_state["pl_form"] = data0
            st.session_state["pl_ctrs"] = ctrs_loaded
            st.success(
                f"Booking **{bno}** · **{len(ctrs_loaded)} conteneur(s)** sur la Packing List "
                f"(Ctr booked={n_booked} · loaded={n_loaded})."
            )
            if err:
                st.warning(err)
            st.rerun()

    form = st.session_state.get("pl_form") or {}
    ctrs0 = st.session_state.get("pl_ctrs") or []

    st.markdown("---")
    st.subheader("2. Formulaire Packing List (modifiable)")

    c1, c2, c3 = st.columns(3)
    with c1:
        packing_no = st.text_input("N° Packing List", value=form.get("Packing_No") or pld.prochain_numero())
        date_pl = st.date_input("Date Packing List", value=date.today())
        bl_no = st.text_input("BL No *", value=str(form.get("BL_No") or ""))
        dated_bl = form.get("Dated_BL")
        try:
            if hasattr(dated_bl, "date"):
                dated_def = dated_bl.date() if not isinstance(dated_bl, date) else dated_bl
            else:
                dated_def = pd.to_datetime(dated_bl, errors="coerce", dayfirst=True)
                dated_def = dated_def.date() if pd.notna(dated_def) else date.today()
        except Exception:
            dated_def = date.today()
        dated_bl_v = st.date_input("Dated (BL)", value=dated_def)
        booking_no = st.text_input("Booking No", value=str(form.get("Booking_No") or ""))
    with c2:
        vessel = st.text_input("Vessel Name", value=str(form.get("Vessel_Name") or ""))
        voyage = st.text_input("Voyage", value=str(form.get("Voyage") or ""))
        pol = st.text_input("Port of Loading", value=str(form.get("Port_Loading") or pld.DEFAULT_POL))
        pod = st.text_input("Port of Discharge", value=str(form.get("Port_Discharge") or ""))
        origin = st.text_input("Country of Origin", value=str(form.get("Country_Origin") or pld.DEFAULT_ORIGIN))
        final_dest = st.text_input("Final Destination", value=str(form.get("Final_Destination") or pod))
    with c3:
        shipper = st.text_input("Shipper", value=str(form.get("Shipper") or pld.DEFAULT_SHIPPER))
        shipper_addr = st.text_area("Shipper address", value=str(form.get("Shipper_Address") or pld.DEFAULT_SHIPPER_ADDR), height=70)
        consignee = st.text_input("Consignee *", value=str(form.get("Consignee") or ""))
        consignee_addr = st.text_area("Consignee address", value=str(form.get("Consignee_Address") or ""), height=70)
        notify = st.text_input("Notify Party", value=str(form.get("Notify") or consignee))
        notify_addr = st.text_area("Notify address", value=str(form.get("Notify_Address") or ""), height=70)

    commodity = st.text_input("Commodity", value=str(form.get("Commodity") or ""))
    product_lines = st.text_area(
        "Description produit (comme sur le SI / Packing List)",
        value=str(form.get("Product_Lines") or ""),
        height=140,
        help=(
            "Exemple type SI :\n"
            "5 CONTAINER EACH CONTAINER WITH HEATING PAD\n"
            "RAW CRUDE SHEA BUTTER IN FLEXITANK OF REP OF BENIN\n"
            "ORIGIN WEST AFRICA MATERIAL DESCRIPTION : FA010C\n"
            "CRUDE SHB\n"
            "HS CODE : 151590990"
        ),
    )
    hc1, hc2, hc3 = st.columns(3)
    with hc1:
        hs_code = st.text_input(
            "HS Code",
            value=str(form.get("Hs_Code") or "151590990"),
            help="Ex. Shea Butter : 151590990",
        )
    with hc2:
        heating = st.selectbox(
            "Heating pad",
            ["OUI", "NON"],
            index=0 if str(form.get("Heating_Pad") or "").upper() in ("OUI", "YES") else 1,
        )
    with hc3:
        st.write("")
        st.write("")
        if st.button("🔄 Reconstruire description SI", help="Recalcule le bloc produit (nb conteneurs × modèle SI)"):
            n_c = len(st.session_state.get("pl_ctrs") or []) or 1
            # nb from number_input if already rendered — use session form ctrs length
            txt = pld.construire_description_pl(
                commodity, n_c, hs_code=hs_code, heating=(heating == "OUI")
            )
            st.session_state["pl_form"] = {
                **(st.session_state.get("pl_form") or {}),
                "Product_Lines": txt,
                "Hs_Code": hs_code,
                "Heating_Pad": heating,
                "Commodity": commodity,
            }
            st.rerun()
    notes = st.text_input("Notes", value=str(form.get("Notes") or ""))

    st.markdown("#### Conteneurs")
    ctrs = _ctr_editor("plnew", ctrs0)

    if st.button("💾 Enregistrer & générer Word", type="primary"):
        if not bl_no.strip():
            st.error("BL No obligatoire.")
        elif not consignee.strip():
            st.error("Consignee obligatoire.")
        elif not any(c.get("Container_No") for c in ctrs):
            st.error("Au moins un conteneur avec N°.")
        elif gpl is None:
            st.error("Module generer_packing_list introuvable.")
        else:
            payload = {
                "Packing_No": packing_no.strip(),
                "Date_PL": date_pl,
                "Booking_No": booking_no.strip(),
                "BL_No": bl_no.strip(),
                "Dated_BL": dated_bl_v,
                "Shipper": shipper.strip(),
                "Shipper_Address": shipper_addr.strip(),
                "Consignee": consignee.strip(),
                "Consignee_Address": consignee_addr.strip(),
                "Notify": notify.strip(),
                "Notify_Address": notify_addr.strip(),
                "Port_Loading": pol.strip(),
                "Port_Discharge": pod.strip(),
                "Country_Origin": origin.strip(),
                "Final_Destination": final_dest.strip(),
                "Vessel_Name": vessel.strip(),
                "Voyage": voyage.strip(),
                "Commodity": commodity.strip(),
                "Product_Lines": product_lines.strip(),
                "Hs_Code": hs_code.strip(),
                "Heating_Pad": heating,
                "Notes": notes.strip(),
            }
            ok, msg, pid = pld.ajouter(payload, ctrs, user=name)
            if not ok:
                st.error(msg)
            else:
                try:
                    chemin = gpl.generer_packing_list(payload, ctrs)
                    st.session_state["pl_last_file"] = chemin
                    st.session_state["pl_form"] = {}
                    st.session_state["pl_ctrs"] = []
                    st.success(msg + f" · Word : `{os.path.basename(chemin)}`")
                except Exception as e:
                    st.success(msg)
                    st.warning(f"Word non généré : {e}")
                st.rerun()

    if st.session_state.get("pl_last_file"):
        fp = st.session_state["pl_last_file"]
        if os.path.exists(fp):
            with open(fp, "rb") as f:
                st.download_button(
                    "⬇️ Télécharger Packing List (Word)",
                    data=f.read(),
                    file_name=os.path.basename(fp),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                )

with tab_reg:
    df = pld.charger()
    if df.empty:
        st.info("Aucune Packing List enregistrée.")
    else:
        st.dataframe(
            df[[c for c in ["Packing_No", "Date_PL", "BL_No", "Booking_No", "Consignee", "Vessel_Name", "Commodity", "Statut", "Cree_Par"] if c in df.columns]],
            hide_index=True,
            width="stretch",
        )
        ids = df["ID"].tolist()
        labels = {i: f"{r.Packing_No} · BL {r.BL_No}" for i, r in zip(df["ID"], df.itertuples())}
        sel = st.selectbox("Régénérer", ids, format_func=lambda i: labels.get(i, str(i)))
        if st.button("🔁 Régénérer Word") and gpl:
            row, ctr = pld.get_full(sel)
            if row:
                try:
                    chemin = gpl.generer_packing_list(row, ctr.to_dict("records") if ctr is not None and not ctr.empty else [])
                    with open(chemin, "rb") as f:
                        st.download_button(
                            "⬇️ Télécharger",
                            data=f.read(),
                            file_name=os.path.basename(chemin),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                    st.success(f"Généré : {chemin}")
                except Exception as e:
                    st.error(str(e))

st.caption("NAP SARL · Packing List · package export/")
