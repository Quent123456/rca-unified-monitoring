import streamlit as st
import pandas as pd

from modules.gsheet_sync import (
    get_gsheet_client,
    get_worksheet_data,
    is_gsheet_configured,
    get_players,
    get_wellness_data
)

st.set_page_config(
    page_title="RCA Monitoring",
    page_icon="🏉",
    layout="wide"
)

st.title("🏉 RCA Unified Monitoring")
st.subheader("Plateforme de suivi de la performance, du wellness et des charges GPS")

st.markdown("---")

# Section 1 : Diagnostic & Test de Connexion Google Sheets
st.subheader("🔌 Statut de la Connexion Google Sheets")

col_status, col_actions = st.columns([2, 1])

with col_status:
    configured = is_gsheet_configured()
    if configured:
        client = get_gsheet_client()
        if client:
            st.success("✅ **Connexion Google Sheets établie avec succès !** Le client `gspread` est authentifié.")
            
            target_sheet = st.secrets.get("spreadsheet_url") or st.secrets.get("spreadsheet_name")
            if target_sheet:
                st.write(f"Feuille de calcul ciblée : `{target_sheet}`")
                try:
                    if target_sheet.startswith("http"):
                        sh = client.open_by_url(target_sheet)
                    else:
                        sh = client.open(target_sheet)
                    
                    worksheets = [ws.title for ws in sh.worksheets()]
                    st.write(f"📋 **Onglets détectés** : {', '.join([f'`{w}`' for w in worksheets])}")
                    
                    # Test rapide de lecture sur l'un des onglets
                    first_ws = worksheets[0]
                    test_df = get_worksheet_data(target_sheet, first_ws)
                    st.caption(f"Test de cache (@st.cache_data ttl=600) sur l'onglet `{first_ws}` : **{len(test_df)} ligne(s)** chargées.")
                except Exception as err:
                    st.warning(f"Client authentifié, mais impossible d'ouvrir la feuille : {err}")
            else:
                st.info("Aucun `spreadsheet_url` ou `spreadsheet_name` spécifié dans `st.secrets`.")
        else:
            st.error("❌ Échec lors de l'authentification avec les identifiants fournis dans `st.secrets`.")
    else:
        st.warning("⚠️ **Mode Démo Actif** : Aucune clé de compte de service n'est configurée dans `.streamlit/secrets.toml`.")
        with st.expander("ℹ️ Comment configurer vos secrets Google Sheets ?"):
            st.markdown("""
            Créez ou modifiez le fichier `.streamlit/secrets.toml` avec l'un des deux formats ci-dessous :

            **Option A - Compte de service GCP (`gcp_service_account`) :**
            ```toml
            spreadsheet_url = "https://docs.google.com/spreadsheets/d/VOTRE_ID_FEUILLE/edit"

            [gcp_service_account]
            type = "service_account"
            project_id = "votre-projet"
            private_key_id = "..."
            private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
            client_email = "service-account@votre-projet.iam.gserviceaccount.com"
            client_id = "..."
            ```

            **Option B - Streamlit Connections GSheets :**
            ```toml
            spreadsheet_url = "https://docs.google.com/spreadsheets/d/VOTRE_ID_FEUILLE/edit"

            [connections.gsheets]
            spreadsheet = "https://docs.google.com/spreadsheets/d/VOTRE_ID_FEUILLE/edit"
            service_account = '''{ ...votre json complet... }'''
            ```
            """)

with col_actions:
    if st.button("🔄 Rafraîchir le cache Google Sheets", use_container_width=True):
        get_worksheet_data.clear()
        st.toast("Cache vidé avec succès !", icon="✅")

st.markdown("---")

# Section 2 : Navigation Rapide
st.subheader("🚀 Modules disponibles")
nav1, nav2, nav3 = st.columns(3)
with nav1:
    st.markdown("### 📝 Saisie Quotidienne")
    st.write("Saisie rapide mobile avec mémoire, recherche dynamique et calcul de la Charge de Foster.")
with nav2:
    st.markdown("### 👤 Tableau Individuel")
    st.write("Profil athlète, jauge ACWR dynamique, radar de bien-être et historique des charges.")
with nav3:
    st.markdown("### 👥 Tableau Collectif")
    st.write("Heatmap d'équipe, matrice de risque ACWR, alertes fatigue et perception d'effort.")

nav4, nav5, nav6 = st.columns(3)
with nav4:
    st.markdown("### 🛰️ Analyse GPS")
    st.write("Import multi-constructeurs, volume, vitesse max, HSR (>20 km/h) et charge externe.")
with nav5:
    st.markdown("### 📄 Rapports PDF")
    st.write("Génération en un clic de rapports de performance A4 prêts pour l'impression.")
with nav6:
    st.markdown("### ⚙️ Gestion de l'Effectif")
    st.write("Ajout, modification et suppression synchronisés en direct sur Google Sheets.")

st.markdown("---")

# Section 3 : Aperçu Effectif
players_df = get_players()
wellness_df = get_wellness_data()

kpi_p1, kpi_p2, kpi_p3 = st.columns(3)
with kpi_p1:
    st.metric("Joueurs actifs dans l'effectif", len(players_df))
with kpi_p2:
    nb_avants = (players_df["groupe"] == "Avants").sum() if "groupe" in players_df else 0
    st.metric("Avants", nb_avants)
with kpi_p3:
    nb_arrieres = (players_df["groupe"] == "Arrières").sum() if "groupe" in players_df else 0
    st.metric("Arrières", nb_arrieres)
