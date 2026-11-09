import datetime
import pandas as pd
import streamlit as st

from modules.gps_processing import calculate_acwr, get_mock_gps_data
from modules.gsheet_sync import get_players, get_wellness_data
from modules.pdf_reports import generate_player_pdf_report

st.set_page_config(
    page_title="Rapports PDF - RCA Monitoring",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Génération de Rapports PDF Individuels")
st.caption("Exportez en un clic une fiche de synthèse athlète combinant Wellness, Charge interne et Métriques GPS.")

players_df = get_players()
wellness_df = get_wellness_data()
gps_df = get_mock_gps_data()

if players_df.empty:
    st.error("Aucun joueur disponible.")
    st.stop()

col_select, col_period = st.columns([2, 1])

def format_full_name(row):
    p = str(row.get('prenom', '')).strip().capitalize()
    n = str(row.get('nom', '')).strip().upper()
    return f"{n} {p}".strip()

with col_select:
    player_names = [format_full_name(row) for _, row in players_df.iterrows()]
    selected_player_name = st.selectbox("Sélectionner un joueur pour le rapport", player_names)
    selected_idx = player_names.index(selected_player_name)
    player_row = players_df.iloc[selected_idx]

with col_period:
    period_days = st.selectbox("Période d'analyse", [7, 14, 30], index=0, format_func=lambda x: f"Derniers {x} jours")

cutoff_date = (datetime.date.today() - datetime.timedelta(days=period_days)).isoformat()

# Normalisation et filtrage des données joueur
if "Charge_sRPE" not in wellness_df.columns and "Charge_Foster" in wellness_df.columns:
    wellness_df["Charge_sRPE"] = wellness_df["Charge_Foster"]

player_nom = str(player_row.get("nom", "")).strip().upper()
player_prenom = str(player_row.get("prenom", "")).strip().upper()

def matches_player(val):
    v = str(val).strip().upper()
    return player_nom in v or player_prenom in v

p_wellness = wellness_df[
    wellness_df["Nom"].apply(matches_player) & (wellness_df["Date"] >= cutoff_date)
]
p_gps = gps_df[
    gps_df["player_name"].apply(matches_player) & (gps_df["date"] >= cutoff_date)
]

# Calcul des synthèses
wellness_summary = {
    "score_hooper_moy": round(p_wellness["Score_Hooper"].mean(), 1) if not p_wellness.empty else "N/A",
    "sommeil_moy": round(p_wellness["Sommeil"].mean(), 1) if not p_wellness.empty else "N/A",
    "fatigue_moy": round(p_wellness["Fatigue"].mean(), 1) if not p_wellness.empty else "N/A",
    "courbatures_moy": round(p_wellness["Courbatures"].mean(), 1) if not p_wellness.empty else "N/A",
    "stress_moy": round(p_wellness["Stress"].mean(), 1) if not p_wellness.empty else "N/A",
    "srpe_total": int(p_wellness["Charge_sRPE"].sum()) if not p_wellness.empty else "N/A"
}

# Calcul ACWR si possible
acwr_df = calculate_acwr(gps_df, metric="total_distance")
player_acwr = "N/A"
if not acwr_df.empty:
    p_acwr_series = acwr_df[acwr_df["player_name"] == selected_player_name]
    if not p_acwr_series.empty:
        player_acwr = p_acwr_series.sort_values(by="date")["acwr"].iloc[-1]

gps_summary = {
    "total_distance_m": int(p_gps["total_distance"].sum()) if not p_gps.empty else "N/A",
    "hsr_distance_m": int(p_gps["hsr_distance"].sum()) if not p_gps.empty else "N/A",
    "sprint_distance_m": int(p_gps["sprint_distance"].sum()) if not p_gps.empty else "N/A",
    "vmax_kmh": round(p_gps["max_speed"].max(), 1) if not p_gps.empty else "N/A",
    "intensity_m_min": round(p_gps["intensity_m_min"].mean(), 1) if not p_gps.empty else "N/A",
    "player_load_total": round(p_gps["player_load"].sum(), 1) if not p_gps.empty else "N/A",
    "acwr": player_acwr
}

st.markdown("---")
st.subheader("👁️ Aperçu de la Fiche Joueur")

# Cartes d'aperçu
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Poste / Groupe", f"{player_row['poste']}", player_row['groupe'])
with col2:
    st.metric("Score Hooper Moyen", f"{wellness_summary['score_hooper_moy']} / 28")
with col3:
    st.metric("Charge Totale sRPE", f"{wellness_summary['srpe_total']} UA")
with col4:
    st.metric("Distance Totale GPS", f"{gps_summary['total_distance_m']} m")

st.markdown("---")

# Bouton de génération et téléchargement
col_btn, _ = st.columns([1, 2])
with col_btn:
    pdf_bytes = generate_player_pdf_report(
        player_name=selected_player_name,
        player_info=player_row.to_dict(),
        wellness_summary=wellness_summary,
        gps_summary=gps_summary,
        period_str=f"Derniers {period_days} jours"
    )

    file_name = f"Rapport_RCA_{selected_player_name.replace(' ', '_')}_{datetime.date.today().isoformat()}.pdf"

    st.download_button(
        label="📥 Télécharger le rapport PDF",
        data=pdf_bytes,
        file_name=file_name,
        mime="application/pdf",
        use_container_width=True
    )

st.success("Le document PDF est prêt et mis en page selon la charte graphique officielle RCA.")
