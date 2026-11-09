import datetime
import pandas as pd
import streamlit as st

from modules.gsheet_sync import get_players, get_wellness_data, save_wellness_entry

st.set_page_config(
    page_title="Saisie Quotidienne - RCA Monitoring",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Saisie Quotidienne (Wellness & RPE)")
st.caption("Évaluation de l'état de forme (Indice de Hooper) et de la charge d'entraînement perçue (sRPE).")

players_df = get_players()
if players_df.empty:
    st.error("Aucun joueur trouvé dans l'effectif.")
    st.stop()

# Onglets : Saisie Individuelle | Saisie Rapide Groupe | Historique Récent
tab_indiv, tab_history = st.tabs(["📋 Formulaire Individuel", "📊 Saisies du Jour"])

with tab_indiv:
    with st.form("form_saisie_quotidienne", clear_on_submit=False):
        col_date, col_player = st.columns([1, 2])
        with col_date:
            date_saisie = st.date_input("Date d'évaluation", value=datetime.date.today())

        with col_player:
            player_options = [
                f"{row['prenom']} {row['nom']} ({row['poste']} - {row['groupe']})"
                for _, row in players_df.iterrows()
            ]
            selected_player_str = st.selectbox("Sélectionner un athlète", player_options)
            player_index = player_options.index(selected_player_str)
            current_player = players_df.iloc[player_index]

        st.markdown("---")
        st.subheader("1. État de forme - Questionnaire de Hooper")
        st.caption("Échelle de 1 (Très favorable) à 7 (Très défavorable).")

        col_h1, col_h2, col_h3, col_h4 = st.columns(4)
        with col_h1:
            sommeil = st.slider("🛌 Sommeil", min_value=1, max_value=7, value=2, help="1: Très réparateur, 7: Nuit blanche / très agité")
        with col_h2:
            fatigue = st.slider("⚡ Fatigue", min_value=1, max_value=7, value=2, help="1: Pleine forme, 7: Épuisé")
        with col_h3:
            courbatures = st.slider("🦵 Courbatures", min_value=1, max_value=7, value=2, help="1: Aucune douleur, 7: Courbatures sévères")
        with col_h4:
            stress = st.slider("🧠 Stress", min_value=1, max_value=7, value=1, help="1: Très serein, 7: Stress maximal")

        hooper_total = sommeil + fatigue + courbatures + stress
        if hooper_total <= 10:
            st.success(f"Score de Hooper : **{hooper_total}/28** (État de forme optimal)")
        elif hooper_total <= 15:
            st.warning(f"Score de Hooper : **{hooper_total}/28** (Vigilance - fatigue modérée)")
        else:
            st.error(f"Score de Hooper : **{hooper_total}/28** (Alerte rouge - fatigue prononcée)")

        st.markdown("---")
        st.subheader("2. Charge d'entraînement de la séance (sRPE)")

        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            session_type = st.selectbox(
                "Type de séance",
                ["Rugby Collectif", "Musculation", "Vitesse / Skills", "Capitaine", "Match", "Récupération"]
            )
        with col_s2:
            duree_min = st.number_input("Durée de la séance (minutes)", min_value=0, max_value=240, value=75, step=5)
        with col_s3:
            rpe = st.slider("RPE de séance (Borg CR-10)", min_value=1, max_value=10, value=6, help="1: Très facile, 10: Effort maximal absolu")

        srpe_charge = duree_min * rpe
        st.info(f"Charge de séance calculée (sRPE) : **{srpe_charge} UA** (Durée {duree_min} min × RPE {rpe})")

        st.markdown("---")
        st.subheader("3. Signalements médicaux & Commentaires")
        col_m1, col_m2 = st.columns([1, 2])
        with col_m1:
            douleur_loc = st.selectbox(
                "Zone de douleur ou gène",
                ["Aucune", "Ischios-jambiers", "Quadriceps", "Adducteurs", "Mollets", "Chevilles", "Genoux", "Épaules", "Dos", "Autre"]
            )
        with col_m2:
            commentaires = st.text_input("Commentaires staff / ressenti joueur", placeholder="Ex: Gêne légère en fin de fractionné")

        submitted = st.form_submit_button("💾 Enregistrer la saisie", use_container_width=True)

        if submitted:
            entry_dict = {
                "Date": date_saisie.isoformat(),
                "Joueur_ID": current_player["id"],
                "Nom": f"{current_player['prenom']} {current_player['nom']}",
                "Poste": current_player["poste"],
                "Groupe": current_player["groupe"],
                "Sommeil": int(sommeil),
                "Fatigue": int(fatigue),
                "Courbatures": int(courbatures),
                "Stress": int(stress),
                "Score_Hooper": int(hooper_total),
                "Session_Type": session_type,
                "RPE": int(rpe),
                "Duree_min": int(duree_min),
                "Charge_sRPE": int(srpe_charge),
                "Douleurs_Localisation": douleur_loc,
                "Commentaires": commentaires
            }

            if save_wellness_entry(entry_dict):
                st.success(f"Données enregistrées pour {current_player['prenom']} {current_player['nom']} !")
                st.rerun()

with tab_history:
    df_all = get_wellness_data()
    st.subheader("Dernières saisies enregistrées")
    if not df_all.empty:
        # Filtrer sur les 7 derniers jours par défaut
        st.dataframe(
            df_all.sort_values(by="Date", ascending=False).head(20),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("Aucune saisie trouvée.")
