import datetime
import pandas as pd
import streamlit as st

from modules.gsheet_sync import (
    get_players,
    get_wellness_data,
    save_wellness_entry,
    is_gsheet_configured
)

st.set_page_config(
    page_title="Saisie Quotidienne - RCA Monitoring",
    page_icon="📝",
    layout="wide"
)

# En-tête ergonomique optimisé mobile
st.title("📝 Saisie Quotidienne RPE & Wellness")
st.caption("Renseignez la charge perçue de la séance (Charge de Foster) et l'état de forme de l'athlète.")

# 1. Chargement et tri de l'effectif
players_df = get_players()
if players_df.empty:
    st.error("⚠️ Aucun joueur disponible dans l'effectif. Vérifiez la connexion Google Sheets ou l'onglet 'Effectif'.")
    st.stop()

# Garantir le tri alphabétique NOM puis Prénom
players_df = players_df.sort_values(
    by=["nom", "prenom"],
    key=lambda s: s.astype(str).str.upper()
).reset_index(drop=True)

# 2. Filtrage par groupe (Équipe A, Équipe B, etc.)
groupes_disponibles = sorted(list(set(players_df["groupe"].dropna().unique())))
if not groupes_disponibles:
    groupes_disponibles = ["Tous"]

col_filter_group, col_date = st.columns([2, 1])

with col_filter_group:
    filtre_groupe = st.radio(
        "🏷️ Filtrer par groupe",
        options=["Tous"] + groupes_disponibles,
        horizontal=True,
        help="Permet de restreindre rapidement la liste des joueurs sur mobile."
    )

with col_date:
    date_seance = st.date_input("📅 Date de la séance", value=datetime.date.today())

# Application du filtre groupe
df_filtered = players_df.copy()
if filtre_groupe != "Tous":
    df_filtered = df_filtered[df_filtered["groupe"] == filtre_groupe]

if df_filtered.empty:
    st.warning(f"Aucun joueur trouvé dans le groupe '{filtre_groupe}'.")
    st.stop()

# Création de la liste déroulante des athlètes
player_options = [
    f"{str(row['nom']).upper()} {str(row['prenom']).capitalize()} — {row['poste']} ({row['groupe']})"
    for _, row in df_filtered.iterrows()
]

st.markdown("### 👤 Sélection de l'athlète")
selected_player_str = st.selectbox(
    "Athlète à évaluer",
    options=player_options,
    label_visibility="collapsed"
)
selected_idx = player_options.index(selected_player_str)
current_player = df_filtered.iloc[selected_idx]

st.markdown("---")

# 3. Formulaire de saisie RPE & Charge de Foster
st.markdown("### 🏉 Paramètres de la Séance")

col_type, col_duree = st.columns([1, 1])

with col_type:
    type_seance = st.selectbox(
        "Type de séance",
        [
            "Rugby Collectif",
            "Musculation",
            "Vitesse / Skills",
            "Capitaine / Mise en place",
            "Match",
            "Récupération Active",
            "Autre"
        ]
    )

with col_duree:
    duree_minutes = st.number_input(
        "Durée effective de la séance (minutes)",
        min_value=5,
        max_value=300,
        value=75,
        step=5
    )

# Échelle de Foster / Borg CR-10 avec descriptions explicites
rpe_descriptions = {
    1: "1 - Repos / Très très facile",
    2: "2 - Facile",
    3: "3 - Modéré",
    4: "4 - Un peu dur",
    5: "5 - Dur",
    6: "6 - Difficile",
    7: "7 - Très dur",
    8: "8 - Épuisant",
    9: "9 - Presque maximal",
    10: "10 - Effort maximal absolu"
}

rpe_val = st.select_slider(
    "Intensité de la séance : RPE (Échelle de Borg CR-10)",
    options=list(range(1, 11)),
    value=5,
    format_func=lambda x: rpe_descriptions[x]
)

# Calcul automatique de la Charge de Foster
charge_foster = int(duree_minutes * rpe_val)

# Affichage direct de la Charge de Foster
col_card1, col_card2 = st.columns([1, 1])
with col_card1:
    st.metric(
        label="⚡ Charge de Foster calculée",
        value=f"{charge_foster} UA",
        delta=f"{duree_minutes} min × RPE {rpe_val}",
        help="Charge de Foster = Durée (min) × RPE (Borg CR-10)"
    )
with col_card2:
    if charge_foster < 300:
        charge_status = "Charge légère / Régénération"
        color = "#10B981"
    elif charge_foster <= 600:
        charge_status = "Charge moyenne / Développement"
        color = "#3B82F6"
    elif charge_foster <= 850:
        charge_status = "Charge élevée / Vigilance"
        color = "#F59E0B"
    else:
        charge_status = "Charge très lourde / Risque de surmenage"
        color = "#EF4444"
    st.markdown(
        f"""
        <div style="background-color: {color}15; border-left: 4px solid {color}; padding: 12px; border-radius: 6px; margin-top: 5px;">
            <b style="color: {color};">Niveau de charge :</b> {charge_status}
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")

# 4. Questionnaire Wellness Optionnel (Valeurs par défaut à 3)
with st.expander("💤 Questionnaire Wellness (Optionnel - État de forme)", expanded=False):
    st.caption("Évaluation de la récupération (échelle de 1 à 7, valeur neutre par défaut : 3).")
    
    col_w1, col_w2, col_w3 = st.columns(3)
    with col_w1:
        fatigue = st.slider("⚡ Fatigue", min_value=1, max_value=7, value=3, help="1: Très en forme, 3: Normal, 7: Épuisé")
    with col_w2:
        courbatures = st.slider("🦵 Courbatures", min_value=1, max_value=7, value=3, help="1: Aucune courbature, 3: Légères, 7: Sévères")
    with col_w3:
        sommeil = st.slider("🛌 Sommeil", min_value=1, max_value=7, value=3, help="1: Très réparateur, 3: Moyen, 7: Très mauvais")

    stress = st.slider("🧠 Niveau de Stress", min_value=1, max_value=7, value=2, help="1: Très détendu, 7: Extrêmement stressé")
    score_hooper = fatigue + courbatures + sommeil + stress

    col_m1, col_m2 = st.columns([1, 2])
    with col_m1:
        douleur_zone = st.selectbox(
            "Zone de douleur ou alerte médicale",
            ["Aucune", "Ischios-jambiers", "Quadriceps", "Adducteurs", "Mollets", "Chevilles", "Genoux", "Épaules", "Dos", "Autre"]
        )
    with col_m2:
        remarque = st.text_input("Commentaires / Ressenti", placeholder="Ex: Bonne séance, sensation de légères raideurs aux ischios")

st.markdown("<br>", unsafe_allow_html=True)

# 5. Bouton d'enregistrement ergonomique pleine largeur pour mobile
btn_submit = st.button("💾 ENREGISTRER LA SÉANCE", type="primary", use_container_width=True)

if btn_submit:
    with st.spinner("Enregistrement dans Google Sheets..."):
        entry_payload = {
            "Date": date_seance.isoformat(),
            "Joueur_ID": current_player.get("id", ""),
            "Nom": str(current_player["nom"]).upper(),
            "Prenom": str(current_player["prenom"]).capitalize(),
            "Poste": current_player.get("poste", "Rugbyman"),
            "Groupe": current_player.get("groupe", "Équipe A"),
            "Session_Type": type_seance,
            "Duree_min": int(duree_minutes),
            "RPE": int(rpe_val),
            "Charge_Foster": int(charge_foster),
            "Charge_sRPE": int(charge_foster),
            "Fatigue": int(fatigue),
            "Courbatures": int(courbatures),
            "Sommeil": int(sommeil),
            "Stress": int(stress),
            "Score_Hooper": int(score_hooper),
            "Douleurs_Localisation": douleur_zone,
            "Commentaires": remarque
        }

        success = save_wellness_entry(entry_payload)

        if success:
            st.toast("✅ Données enregistrées et cache actualisé !", icon="🏉")
            st.success(
                f"✅ **Séance enregistrée avec succès pour {current_player['prenom']} {current_player['nom']} !**  \n"
                f"• Charge de Foster : **{charge_foster} UA** ({duree_minutes} min × RPE {rpe_val})  \n"
                f"• Type de séance : **{type_seance}** | Date : **{date_seance}**"
            )
        else:
            st.error("Erreur lors de l'enregistrement. Veuillez vérifier les permissions de la feuille de calcul.")

st.markdown("---")

# Historique des saisies récentes pour contrôle rapide
st.subheader("📋 Dernières saisies enregistrées")
recent_df = get_wellness_data()
if not recent_df.empty:
    display_cols = [c for c in ["Date", "Nom", "Prenom", "Session_Type", "Duree_min", "RPE", "Charge_Foster", "Score_Hooper"] if c in recent_df.columns]
    if not display_cols:
        display_cols = recent_df.columns[:8]
    st.dataframe(
        recent_df.sort_values(by="Date", ascending=False).head(10)[display_cols],
        use_container_width=True,
        hide_index=True
    )
else:
    st.info("Aucune saisie récente disponible.")
