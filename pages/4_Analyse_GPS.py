import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from modules.gps_processing import calculate_acwr, get_mock_gps_data, normalize_gps_dataframe

st.set_page_config(
    page_title="Analyse GPS - RCA Monitoring",
    page_icon="🛰️",
    layout="wide"
)

st.title("🛰️ Analyse des Données GPS & Charges Externes")
st.caption("Import, standardisation et analyse des métriques de déplacement et de vitesse.")

# Section d'import
with st.expander("📥 Importer un fichier GPS (CSV ou Excel)", expanded=False):
    uploaded_file = st.file_uploader(
        "Déposer l'export GPS de votre logiciel (Catapult, Titan, Statsports, etc.)",
        type=["csv", "xlsx"]
    )

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            raw_df = pd.read_csv(uploaded_file)
        else:
            raw_df = pd.read_excel(uploaded_file)
        gps_df = normalize_gps_dataframe(raw_df)
        st.success(f"Fichier '{uploaded_file.name}' importé et standardisé avec succès !")
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {e}")
        gps_df = get_mock_gps_data()
else:
    gps_df = get_mock_gps_data()
    st.info("💡 Données de démonstration chargées. Déposez un fichier ci-dessus pour analyser vos propres données.")

if gps_df.empty:
    st.warning("Aucune métrique GPS disponible.")
    st.stop()

# Barre latérale pour filtrage
with st.sidebar:
    st.header("⚙️ Filtres GPS")
    dates_dispo = sorted(gps_df["date"].unique(), reverse=True)
    selected_date = st.selectbox("Sélectionner une séance / match", ["Toutes les séances"] + list(dates_dispo))

    gps_groups = sorted([str(g) for g in gps_df["groupe"].dropna().unique() if str(g).strip()])
    if not gps_groups:
        gps_groups = ["Avants", "Arrières"]
    groupe_choice = st.radio("Groupe", ["Tous"] + gps_groups, horizontal=True)

df_view = gps_df.copy()
if selected_date != "Toutes les séances":
    df_view = df_view[df_view["date"] == selected_date]
if groupe_choice != "Tous":
    df_view = df_view[df_view["groupe"] == groupe_choice]

# KPIs de séance
st.markdown("### 📌 Synthèse de la séance")
kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

with kpi1:
    dist_moy = int(df_view["total_distance"].mean()) if "total_distance" in df_view else 0
    st.metric("Distance Moyenne", f"{dist_moy} m")

with kpi2:
    hsr_moy = int(df_view["hsr_distance"].mean()) if "hsr_distance" in df_view else 0
    st.metric("HSR Moyen (> 20 km/h)", f"{hsr_moy} m")

with kpi3:
    vmax_peak = round(df_view["max_speed"].max(), 1) if "max_speed" in df_view else 0
    st.metric("Vitesse Max Pic", f"{vmax_peak} km/h")

with kpi4:
    load_moy = round(df_view["player_load"].mean(), 1) if "player_load" in df_view else 0
    st.metric("PlayerLoad Moyen", f"{load_moy} UA")

with kpi5:
    intensite_moy = round(df_view["intensity_m_min"].mean(), 1) if "intensity_m_min" in df_view else 0
    st.metric("Intensité Moyenne", f"{intensite_moy} m/min")

st.markdown("---")

# Graphiques comparatifs
tab_sprint, tab_volume, tab_acwr = st.tabs(["⚡ Profil de Vitesse & HSR", "🏃 Volume & Intensité", "⚖️ Ratio ACWR (7j / 28j)"])

with tab_sprint:
    st.subheader("Courses Haute Intensité (HSR > 20 km/h) & Distance Sprint (> 25 km/h)")
    if "hsr_distance" in df_view and "sprint_distance" in df_view:
        fig_hsr = px.bar(
            df_view.sort_values(by="hsr_distance", ascending=False),
            x="player_name",
            y=["hsr_distance", "sprint_distance"],
            barmode="group",
            labels={"value": "Distance (mètres)", "player_name": "Joueur", "variable": "Métrique"},
            color_discrete_map={"hsr_distance": "#0D9488", "sprint_distance": "#F97316"},
            title="Distribution HSR vs Sprint par joueur"
        )
        fig_hsr.update_layout(xaxis_tickangle=-45, height=420, margin=dict(t=40, b=80))
        st.plotly_chart(fig_hsr, use_container_width=True)

with tab_volume:
    st.subheader("Volume de déplacement & Charge relative")
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        fig_dist = px.bar(
            df_view.sort_values(by="total_distance", ascending=False),
            x="player_name",
            y="total_distance",
            color="groupe",
            labels={"total_distance": "Distance Totale (m)", "player_name": "Joueur"},
            title="Distance Totale (m)"
        )
        fig_dist.update_layout(xaxis_tickangle=-45, height=380)
        st.plotly_chart(fig_dist, use_container_width=True)

    with col_v2:
        fig_int = px.scatter(
            df_view,
            x="intensity_m_min",
            y="player_load",
            color="groupe",
            hover_name="player_name",
            size="total_distance",
            color_discrete_map={"Avants": "#1E3A8A", "Arrières": "#0D9488"},
            labels={"intensity_m_min": "Intensité (m/min)", "player_load": "PlayerLoad (UA)"},
            title="Intensité Relative (m/min) vs PlayerLoad"
        )
        fig_int.update_layout(height=380)
        st.plotly_chart(fig_int, use_container_width=True)

with tab_acwr:
    st.subheader("Suivi du Ratio Aigu:Chronique (ACWR)")
    st.caption("Zone optimale de performance : 0.8 à 1.3 | Risque de blessure accru : > 1.5")

    acwr_df = calculate_acwr(gps_df, metric="total_distance")
    if not acwr_df.empty:
        latest_acwr = acwr_df.sort_values(by="date").groupby("player_name").last().reset_index()

        fig_acwr = px.scatter(
            latest_acwr,
            x="player_name",
            y="acwr",
            color="acwr",
            color_continuous_scale=["#10B981", "#F59E0B", "#EF4444"],
            range_color=[0.6, 1.8],
            labels={"acwr": "Ratio ACWR", "player_name": "Joueur"},
            title="Dernier Ratio ACWR calculé par joueur"
        )
        # Ajout des zones seuils
        fig_acwr.add_hrect(y0=0.8, y1=1.3, fillcolor="#10B981", opacity=0.15, line_width=0, annotation_text="Sweet Spot (0.8 - 1.3)")
        fig_acwr.add_hline(y=1.5, line_dash="dash", line_color="#EF4444", annotation_text="Danger Zone (> 1.5)")
        fig_acwr.update_layout(xaxis_tickangle=-45, height=420)
        st.plotly_chart(fig_acwr, use_container_width=True)
    else:
        st.info("Données insuffisantes pour calculer le ratio ACWR (nécessite plusieurs dates).")

# Tableau complet des données
st.markdown("---")
st.subheader("📑 Tableau des données détaillées")
st.dataframe(df_view, use_container_width=True, hide_index=True)
