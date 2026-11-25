# NAP SARL — Module Export v2 (aligné Export General Report)

## Processus

1. **Transitaires** — comptes forwarders (HAMA INTER, RAJIBER, TRANS SOHA…)
2. **Bookings** — 1 ligne = 1 Booking/BL (feuille GENERAL)
3. **Conteneurs** — 1 ligne = 1 conteneur, BL/Booking répétés (Stuffing Report)
4. **Dashboard** — synthèse (SUMMARY)
5. **Soldes Forwarders** — facturation / solde

## Fichiers

| Fichier | Emplacement |
|---------|-------------|
| `export_data.py` | Racine projet |
| `transitaires_data.py` | Racine projet |
| `pages/7_Transitaires.py` | pages/ |
| `pages/8_Export_Dashboard.py` | pages/ |
| `pages/8_Export_Bookings.py` | pages/ |
| `pages/8_Export_Conteneurs.py` | pages/ |
| `pages/9_Export_Soldes_Forwarders.py` | pages/ |

Excel généré : **`export_master.xlsx`** (feuilles Bookings + Conteneurs)

## Ordre d'utilisation

1. Créer les transitaires
2. Créer les bookings (contrat, buyer, forwarder, n° booking/BL, ctr booked)
3. Enregistrer chaque conteneur lié au booking (camion, n°, plomb, poids…)
4. Le booking met à jour auto : Container_Loaded, Balance, Stuffing_Status, Net_Qty
