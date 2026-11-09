# 🏉 RCA Unified Monitoring

Plateforme unifiée Streamlit de suivi de la performance, du bien-être (*wellness*) et des charges GPS pour le **RCA** (Rugby Club).

---

## 🚀 Fonctionnalités Principales

1. **📝 Saisie Quotidienne & Questionnaire de Hooper**
   - Évaluation de l'état de forme sur 4 composantes : Sommeil, Fatigue, Courbatures, Stress (échelle 1 à 7).
   - Calcul automatique de l'indice total de Hooper avec alertes visuelles de fatigue.
   - Saisie de la charge interne perçue via le **sRPE** (Échelle de Borg CR-10 × Durée de séance en minutes).
   - Déclaration des gênes et douleurs localisées pour le staff médical.

2. **📊 Suivi Collectif & Tableau de Bord Effectif**
   - Vue d'ensemble de la disponibilité du groupe (Avants vs Arrières).
   - Surveillance en temps réel des joueurs en alerte (Hooper ≥ 15, douleurs actives).
   - Analyse temporelle des charges collectives quotidiennes et hebdomadaires.
   - Matrice interactive État de forme vs Charge d'entraînement.

3. **🛰️ Analyse des Données GPS & Charges Externes**
   - Importation glisser-déposer de fichiers GPS (CSV ou Excel) avec standardisation automatique des colonnes (Catapult, Titan, Statsports, etc.).
   - Métriques clés rugby : Distance totale, Course haute intensité (HSR > 20 km/h), Sprint (> 25 km/h), Vitesse Max, PlayerLoad, Accélérations/Décélérations, Intensité relative (m/min).
   - Calcul du **Ratio Charge Aiguë:Chronique (ACWR)** sur base glissante 7j / 28j avec identification de la *Sweet Spot* (0.8 - 1.3) et de la *Danger Zone* (> 1.5).

4. **📄 Génération de Rapports PDF Individuels**
   - Fiches de synthèse athlète éditées au format standard A4 via ReportLab.
   - Intégration de la charte graphique officielle du club, des moyennes wellness, des pics GPS et des recommandations du staff.
   - Téléchargement en 1 clic pour les entraîneurs et préparateurs physiques.

5. **🔌 Synchronisation Google Sheets Optimisée**
   - Authentification sécurisée via les secrets Streamlit (`st.secrets`).
   - Mise en cache intelligente (`@st.cache_data(ttl=600)`) pour prévenir les erreurs de dépassement de quota (HTTP 429 Rate Limit).
   - Mode démo automatique avec jeu de données simulé réaliste si aucune connexion distante n'est configurée.

---

## 📁 Architecture du Projet

```text
rca-unified-monitoring/
│
├── .gitignore
├── README.md
├── requirements.txt
├── app.py                         # Hub central & diagnostic de connexion
│
├── .streamlit/
│   ├── config.toml                # Thème officiel RCA (Bleu marine, typographie)
│   └── secrets.toml (ignoré Git)  # Clés de service Google Sheets
│
├── modules/
│   ├── __init__.py
│   ├── alerts.py                  # Seuils et alertes ACWR, monotonie et contrainte
│   ├── calculations.py            # Calculs de charge Foster, moyennes glissantes, ACWR
│   ├── database.py                # Base de données locale de secours
│   ├── gps_processing.py          # Traitement Catapult, splits, PlayerLoad, injection Excel
│   ├── gsheets.py                 # Multi-onglets Google Sheets (Saisies, RPE INDV, Lignes)
│   ├── gsheet_sync.py             # Authentification gspread polyvalente & cache
│   ├── pdf_reports.py             # Moteur PDF officiel Coachs (FPDF paysage) & Fiche A4 (ReportLab)
│   ├── photo_manager.py           # Extraction trombinoscope, matching et encodage base64
│   ├── rpe_connector.py           # Jointure ultra-robuste GPS + RPE (Nom, Date, 'NR', skipna=True)
│   ├── visualizations.py          # Composants graphiques Plotly (Heatmaps, jauges, radars)
│   └── weekly_memory.py           # Mémoire sélective & rapports hebdomadaires multi-séances
│
├── pages/
│   ├── 1_Saisie_Quotidienne.py    # Formulaire mobile avec mémorisation et bouton sticky
│   ├── 2_Tableau_Individuel.py    # Bilan individuel, jauge ACWR et radar de bien-être
│   ├── 3_Tableau_Collectif.py     # Dashboard d'équipe, heatmap ACWR et alertes
│   ├── 4_Analyse_GPS.py           # Analyse Catapult, attribution boîtiers, export Excel & rapports
│   └── 5_Gestion_Joueurs.py       # Administration de l'effectif synchronisée Google Sheets
│
└── scripts/                       # Utilitaires de test et d'initialisation
    ├── init_gsheets.py
    ├── test_gsheets_connection.py
    └── test_write_saisies.py
```

---

## 🛠️ Installation et Lancement

### 1. Cloner le projet et créer l'environnement virtuel

```bash
git clone https://github.com/Quent123456/rca-unified-monitoring.git
cd rca-unified-monitoring
python -m venv .venv
# Sur Windows :
.venv\Scripts\activate
# Sur macOS/Linux :
source .venv/bin/activate
```

### 2. Installer les dépendances

```bash
pip install -r requirements.txt
```

### 3. Lancer l'application Streamlit

```bash
streamlit run app.py
```

L'application s'ouvre automatiquement dans votre navigateur à l'adresse `http://localhost:8501`.

---

## 🔑 Configuration de Google Sheets (Optionnel)

Pour relier l'application à votre feuille de calcul Google Sheets, créez un fichier `.streamlit/secrets.toml` à la racine :

```toml
spreadsheet_url = "https://docs.google.com/spreadsheets/d/VOTRE_ID_DE_FEUILLE/edit"

[gcp_service_account]
type = "service_account"
project_id = "votre-projet-gcp"
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "service-account@votre-projet.iam.gserviceaccount.com"
client_id = "..."
```

*Note : N'oubliez pas de partager votre feuille de calcul Google Sheets avec l'adresse e-mail `client_email` de votre compte de service avec les droits "Éditeur".*

---

## 📈 Guide des Métriques Clés

| Métrique | Description | Seuil / Interprétation |
| :--- | :--- | :--- |
| **Score de Hooper** | Somme (Sommeil + Fatigue + Courbatures + Stress) sur échelle 1 à 7 | $\le 10$ : Optimal \| $11-14$ : Vigilance \| $\ge 15$ : Alerte Fatigue |
| **Charge sRPE** | Durée de la séance (min) $\times$ RPE perçu (1 à 10) | Unités Arbitraires (UA) de charge interne |
| **HSR** | Course à Haute Intensité | Distance parcourue $> 20\text{ km/h}$ ($> 5.5\text{ m/s}$) |
| **Distance Sprint** | Vitesse maximale pure | Distance parcourue $> 25\text{ km/h}$ ($> 7.0\text{ m/s}$) |
| **ACWR** | Ratio Charge Aiguë (7 jours) / Charge Chronique (28 jours) | $0.8 - 1.3$ : *Sweet Spot* \| $> 1.5$ : Risque de blessure accru |

