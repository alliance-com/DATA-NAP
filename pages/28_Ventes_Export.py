# -*- coding: utf-8 -*-
"""
Ventes liées à l'exportation — NAP SARL
Produits : Tourteau Karité/Soja/Cotton, Graine Soja, Huiles, Beurre Karité
"""
import streamlit as st
import pandas as pd
from datetime import date
import sys
import os

_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

try:
    from export import vente_data as vd
except ImportError:
    try:
        import vente_data as vd
    except ImportError:
        import sys, os
        _root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        if _root not in sys.path:
            sys.path.insert(0, _root)
        _exp = os.path.join(_root, "export")
        if _exp not in sys.path:
            sys.path.insert(0, _exp)
        try:
            from export import vente_data as vd
        except ImportError:
            import vente_data as vd

st.set_page_config(page_title="Ventes Export", page_icon="🚢", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "User"

st.title("🚢 Ventes — Export NAP SARL")
st.caption(
    "Enregistrement des ventes export · lien Contrat / BL / Commercial Invoice · "
    "produits : Tourteaux, Graines, Huiles, Beurre Karité"
)

try:
    df = vd.charger()
except Exception as e:
    st.error(str(e))
    st.stop()

# KPI
c1, c2, c3, c4 = st.columns(4)
c1.metric("Ventes", len(df))
c2.metric("Clients", df["Client"].nunique() if not df.empty else 0)
if not df.empty:
    vol = df["Quantite_MT"].sum()
    # totaux par devise
    tot_eur = df[df["Devise"] == "EUR"]["Montant_Net"].sum()
    tot_usd = df[df["Devise"] == "USD"]["Montant_Net"].sum()
else:
    vol = tot_eur = tot_usd = 0.0
c3.metric("Volume (MT)", f"{vol:,.2f}")
c4.metric("Net EUR / USD", f"{tot_eur:,.0f} / {tot_usd:,.0f}")

tab_new, tab_list, tab_edit, tab_dash = st.tabs([
    "➕ Nouvelle vente",
    "📋 Registre",
    "✏️ Modifier / lier export",
    "📊 Par produit",
])

# ---------- NOUVELLE ----------
with tab_new:
    st.subheader("Saisir une vente export")
    with st.form("form_vente"):
        st.markdown("#### Client & référence")
        a1, a2, a3 = st.columns(3)
        with a1:
            ref = st.text_input("Réf. vente", value=vd.prochain_ref())
            d_vente = st.date_input("Date vente", value=date.today())
            statut = st.selectbox("Statut", vd.STATUTS, index=1)
        with a2:
            clients = vd.liste_clients()
            cl_sel = st.selectbox("Client (liste)", ["— Saisie libre —"] + clients)
            client = st.text_input(
                "Client *",
                value="" if cl_sel == "— Saisie libre —" else cl_sel,
                placeholder="BUNGE LODERS CROKLAAN B.V",
            )
            pays = st.text_input("Pays client", placeholder="Netherlands / UAE / …")
        with a3:
            adr = st.text_area("Adresse client", height=90)

        st.markdown("#### Produit & montants")
        b1, b2, b3, b4 = st.columns(4)
        with b1:
            produit = st.selectbox("Produit *", vd.PRODUITS_VENTE)
            qte = st.number_input("Quantité (MT) *", min_value=0.0, step=0.01, format="%.3f")
        with b2:
            pu = st.number_input("Prix unitaire *", min_value=0.0, step=1.0, format="%.2f")
            devise = st.selectbox("Devise", vd.DEVISES)
        with b3:
            avance = st.number_input("Avance / Less Advance", min_value=0.0, step=100.0, format="%.2f")
            incoterm = st.selectbox("Incoterm", vd.INCOTERMS)
        with b4:
            brut, av, net = vd.calculer(qte, pu, avance)
            st.metric("Brut", f"{brut:,.2f}")
            st.metric("Net", f"{net:,.2f} {devise}")

        st.markdown("#### Lien export (optionnel — à compléter plus tard)")
        c1, c2, c3 = st.columns(3)
        with c1:
            contrat = st.text_input("N° Contrat", placeholder="4600064934")
            bl = st.text_input("N° BL / Connaissement")
        with c2:
            invoice = st.text_input("N° Commercial Invoice", placeholder="NAP/1134/2026")
            dest = st.text_input("Destination finale", placeholder="TANJUNG PELEPAS / MERSIN")
        with c3:
            liner = st.text_input("Shipping Line", value="MAERSK LINE")
            n_ctr = st.number_input("Nb conteneurs", min_value=0, step=1, value=0)

        notes = st.text_input("Notes")
        ok = st.form_submit_button("💾 Enregistrer la vente", type="primary")

    if ok:
        if not client.strip():
            st.error("Client obligatoire.")
        elif qte <= 0 or pu <= 0:
            st.error("Quantité et prix > 0.")
        else:
            ok2, msg, eid = vd.ajouter({
                "Ref_Vente": ref.strip(),
                "Date_Vente": d_vente,
                "Client": client.strip(),
                "Client_Adresse": adr.strip(),
                "Pays_Client": pays.strip(),
                "Produit": produit,
                "Quantite_MT": qte,
                "PU": pu,
                "Devise": devise,
                "Avance": avance,
                "Incoterm": incoterm,
                "Contrat_No": contrat,
                "BL_No": bl,
                "Invoice_No": invoice,
                "Destination": dest,
                "Shipping_Line": liner,
                "Nb_Containers": n_ctr,
                "Statut": statut,
                "Notes": notes,
            }, user=name)
            if ok2:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

# ---------- REGISTRE ----------
with tab_list:
    st.subheader("Registre des ventes")
    if df.empty:
        st.info("Aucune vente.")
    else:
        f1, f2, f3 = st.columns(3)
        with f1:
            pf = st.selectbox("Produit", ["Tous"] + vd.PRODUITS_VENTE)
        with f2:
            cf = st.selectbox("Client", ["Tous"] + vd.liste_clients())
        with f3:
            sf = st.selectbox("Statut", ["Tous"] + vd.STATUTS)
        vue = df.copy()
        if pf != "Tous":
            vue = vue[vue["Produit"] == pf]
        if cf != "Tous":
            vue = vue[vue["Client"] == cf]
        if sf != "Tous":
            vue = vue[vue["Statut"] == sf]
        aff = vue.copy()
        aff["Date_Vente"] = pd.to_datetime(aff["Date_Vente"], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(
            aff[[
                "ID", "Ref_Vente", "Date_Vente", "Client", "Produit", "Quantite_MT",
                "PU", "Devise", "Montant_Net", "Contrat_No", "BL_No", "Invoice_No", "Statut",
            ]],
            hide_index=True, width="stretch", height=420,
        )
        st.download_button(
            "⬇️ Export Excel",
            data=vd.to_excel_bytes(vue),
            file_name=f"Ventes_Export_{date.today():%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ---------- MODIFIER ----------
with tab_edit:
    st.subheader("Modifier une vente · rattacher Contrat / BL / Invoice")
    if df.empty:
        st.info("Aucune vente.")
    else:
        opts = {
            f"{r.Ref_Vente} · {r.Client} · {r.Produit} · {float(r.Montant_Net):,.0f} {r.Devise}": int(r.ID)
            for _, r in df.iterrows()
        }
        choix = st.selectbox("Vente", list(opts.keys()))
        eid = opts[choix]
        row = df[df["ID"] == eid].iloc[0]
        with st.form("form_edit_vente"):
            e1, e2 = st.columns(2)
            with e1:
                d2 = st.date_input(
                    "Date",
                    value=row["Date_Vente"].date() if pd.notna(row["Date_Vente"]) else date.today(),
                )
                client2 = st.text_input("Client", value=str(row["Client"] or ""))
                adr2 = st.text_area("Adresse", value=str(row["Client_Adresse"] or ""), height=70)
                pays2 = st.text_input("Pays", value=str(row["Pays_Client"] or ""))
                prod2 = st.selectbox(
                    "Produit",
                    vd.PRODUITS_VENTE,
                    index=vd.PRODUITS_VENTE.index(row["Produit"])
                    if row["Produit"] in vd.PRODUITS_VENTE else 0,
                )
            with e2:
                qte2 = st.number_input("Qté MT", value=float(row["Quantite_MT"] or 0), step=0.01, format="%.3f")
                pu2 = st.number_input("PU", value=float(row["PU"] or 0), step=1.0, format="%.2f")
                devise2 = st.selectbox(
                    "Devise",
                    vd.DEVISES,
                    index=vd.DEVISES.index(row["Devise"]) if row["Devise"] in vd.DEVISES else 0,
                )
                av2 = st.number_input("Avance", value=float(row["Avance"] or 0), step=100.0, format="%.2f")
                statut2 = st.selectbox(
                    "Statut",
                    vd.STATUTS,
                    index=vd.STATUTS.index(row["Statut"]) if row["Statut"] in vd.STATUTS else 0,
                )
            st.markdown("**Lien export**")
            l1, l2, l3 = st.columns(3)
            with l1:
                contrat2 = st.text_input("Contrat", value=str(row["Contrat_No"] or ""))
                bl2 = st.text_input("BL", value=str(row["BL_No"] or ""))
            with l2:
                inv2 = st.text_input("Commercial Invoice", value=str(row["Invoice_No"] or ""))
                dest2 = st.text_input("Destination", value=str(row["Destination"] or ""))
            with l3:
                liner2 = st.text_input("Shipping Line", value=str(row["Shipping_Line"] or ""))
                nctr2 = st.number_input("Nb ctr", value=int(row["Nb_Containers"] or 0), step=1)
            notes2 = st.text_input("Notes", value=str(row["Notes"] or ""))
            if st.form_submit_button("💾 Enregistrer", type="primary"):
                ok, msg = vd.maj(eid, {
                    "Date_Vente": d2,
                    "Client": client2,
                    "Client_Adresse": adr2,
                    "Pays_Client": pays2,
                    "Produit": prod2,
                    "Quantite_MT": qte2,
                    "PU": pu2,
                    "Devise": devise2,
                    "Avance": av2,
                    "Contrat_No": contrat2,
                    "BL_No": bl2,
                    "Invoice_No": inv2,
                    "Destination": dest2,
                    "Shipping_Line": liner2,
                    "Nb_Containers": nctr2,
                    "Statut": statut2,
                    "Notes": notes2,
                    "Incoterm": row.get("Incoterm") or "FOB COTONOU PORT",
                })
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

# ---------- DASH PRODUIT ----------
with tab_dash:
    st.subheader("Volumes & montants par produit")
    if df.empty:
        st.info("Aucune donnée.")
    else:
        g = (
            df.groupby("Produit")
            .agg(
                Nb=("ID", "count"),
                Volume_MT=("Quantite_MT", "sum"),
                Net_EUR=("Montant_Net", lambda s: s[df.loc[s.index, "Devise"] == "EUR"].sum() if len(s) else 0),
            )
            .reset_index()
        )
        # recalcul simple par devise
        rows = []
        for prod in vd.PRODUITS_VENTE:
            sub = df[df["Produit"] == prod]
            if sub.empty:
                continue
            rows.append({
                "Produit": prod,
                "Nb ventes": len(sub),
                "Volume MT": round(sub["Quantite_MT"].sum(), 3),
                "Net EUR": round(sub[sub["Devise"] == "EUR"]["Montant_Net"].sum(), 2),
                "Net USD": round(sub[sub["Devise"] == "USD"]["Montant_Net"].sum(), 2),
                "Net FCFA": round(sub[sub["Devise"] == "FCFA"]["Montant_Net"].sum(), 0),
            })
        if rows:
            st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
            st.bar_chart(pd.DataFrame(rows).set_index("Produit")["Volume MT"])

st.caption("NAP SARL — Ventes export · fichier `ventes_export.xlsx`")
