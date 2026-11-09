import datetime
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from modules.gsheet_sync import get_players, get_wellness_data

st.set_page_config(
    page_title="Suivi Collectif - RCA Monitoring",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Suivi Collectif & Tableau de Bord Effectif")
st.caption("Surveillance globale de l'état de forme, des charges internes et de la disponibilité du groupe.")

players_df = get_players()
wellness_df = get_wellness_data()

if wellness_df.empty:
    st.info("Aucune donnée disponible. Veuillez effectuer des saisies dans la page 'Saisie Quotidienne'.")
    st.stop()

# Filtres dans la barre latérale
with st.sidebar:
    st.header("🔍 Filtres d'affichage")
    groupe_filter = st.selectbox("Groupe d'entraînement", ["Tous", "Avants", "Arrières"])
    
    plage_jours = st.slider("Période d'analyse (jours)", min_value=7, max_value=30, value=14, step=7)
    
    # Filtrage par date
    cutoff_date = (datetime.date.today() - datetime.timedelta(days=plage_jours)).isoformat()
    filtered_df = wellness_df[wellness_df["Date"] >= cutoff_date].copy()

    if groupe_filter != "Tous":
        filtered_df = filtered_df[filtered_df["Groupe"] == groupe_filter]

if filtered_df.empty:
    st.warning("Aucune donnée pour les critères sélectionnés.")
    st.stop()

# Calcul des KPIs clés
latest_date = filtered_df["Date"].max()
latest_entries = filtered_df[filtered_df["Date"] == latest_date]

hooper_moy = round(latest_entries["Score_Hooper"].mean(), 1) if not latest_entries.empty else 0
srpe_moy = int(latest_entries["Charge_sRPE"].mean()) if not latest_entries.empty else 0
alertes_hooper = (latest_entries["Score_Hooper"] >= 15).sum() if not latest_entries.empty else 0
alertes_douleurs = (latest_entries["Douleurs_Localisation"] != "Aucune").sum() if not latest_entries.empty else 0

# Bandeau de KPIs
kpi1, kpi2, kpi3, kpi4 = st.columns(4)

with kpi1:
    hooper_status = "Optimal" if hooper_moy <= 11 else ("Vigilance" if hooper_moy <= 14 else "Alerte Fatigue")
    st.metric(
        label="Score Hooper Moyen (Dernière séance)",
        value=f"{hooper_moy} / 28",
        delta=hooper_status,
        delta_color="inverse" if hooper_moy > 12 else "normal"
    )

with kpi2:
    st.metric(
        label="Charge sRPE Moyenne",
        value=f"{srpe_moy} UA",
        help="Moyenne Charge = Durée (min) × RPE"
    )

with kpi3:
    st.metric(
        label="Alertes Fatigue Élevée",
        value=f"{alertes_hooper} joueur(s)",
        delta="Hooper ≥ 15" if alertes_hooper > 0 else "Aucune alerte",
        delta_color="inverse" if alertes_hooper > 0 else "off"
    )

with kpi4:
    st.metric(
        label="Signalements Médicaux",
        value=f"{alertes_douleurs} joueur(s)",
        delta="Gênes / Douleurs" if alertes_douleurs > 0 else "Effectif disponible",
        delta_color="inverse" if alertes_douleurs > 0 else "normal"
    )

st.markdown("---")

# Graphiques principaux
col_graph1, col_graph2 = st.columns(2)

with col_graph1:
    st.subheader("📈 Évolution temporelle de la charge collective (sRPE)")
    daily_load = filtered_df.groupby("Date")["Charge_sRPE"].agg(["mean", "sum"]).reset_index()

    fig_load = go.Figure()
    fig_load.add_trace(go.Bar(
        x=daily_load["Date"],
        y=daily_load["sum"],
        name="Charge Totale Équipe",
        marker_color="#1E3A8A",
        opacity=0.7
    ))
    fig_load.add_trace(go.Scatter(
        x=daily_load["Date"],
        y=daily_load["mean"],
        name="Charge Moyenne / Joueur",
        yaxis="y2",
        mode="lines+markers",
        line=dict(color="#F97316", width=3)
    ))
    fig_load.update_layout(
        xaxis_title="Date",
        yaxis=dict(title="Charge Totale (UA)"),
        yaxis2=dict(title="Charge Moyenne (UA)", overlaying="y", side="right"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=20, r=20, t=30, b=20),
        height=380
    )
    st.plotly_chart(fig_load, use_container_width=True)

with col_graph2:
    st.subheader("🎯 État de Forme vs Charge (Dernière séance)")
    if not latest_entries.empty:
        fig_scatter = px.scatter(
            latest_entries,
            x="Score_Hooper",
            y="Charge_sRPE",
            color="Groupe",
            hover_name="Nom",
            size="RPE",
            color_discrete_map={"Avants": "#1E3A8A", "Arrières": "#0D9488"},
            labels={"Score_Hooper": "Indice de Hooper (Fatigue)", "Charge_sRPE": "Charge sRPE (UA)"}
        )
        # Lignes repères de vigilance
        fig_scatter.add_vline(x=14, line_dash="dash", line_color="#EF4444", annotation_text="Alerte Fatigue")
        fig_scatter.update_layout(
            margin=dict(l=20, r=20, t=30, b=20),
            height=380
        )
        st.plotly_chart(fig_scatter, use_container_width=True)
    else:
        st.info("Aucune donnée pour la dernière séance.")

# Tableau récapitulatif par joueur
st.markdown("---")
st.subheader("📋 Synthèse individuelle de l'effectif")

summary_by_player = filtered_df.groupby(["Nom", "Poste", "Groupe"]).agg(
    Score_Hooper_Moy=("Score_Hooper", "mean"),
    Sommeil_Moy=("Sommeil", "mean"),
    Fatigue_Moy=("Fatigue", "mean"),
    Courbatures_Moy=("Courbatures", "mean"),
    Stress_Moy=("Stress", "mean"),
    Charge_sRPE_Cumulee=("Charge_sRPE", "sum"),
    Derniere_Douleur=("Douleurs_Localisation", "last")
).reset_index()

summary_by_player["Score_Hooper_Moy"] = summary_by_player["Score_Hooper_Moy"].round(1)
summary_by_player["Sommeil_Moy"] = summary_by_player["Sommeil_Moy"].round(1)
summary_by_player["Fatigue_Moy"] = summary_by_player["Fatigue_Moy"].round(1)
summary_by_player["Courbatures_Moy"] = summary_by_player["Courbatures_Moy"].round(1)
summary_by_player["Stress_Moy"] = summary_by_player["Stress_Moy"].round(1)

# Mise en évidence visuelle des statuts
def highlight_status(row):
    color = "transparent"
    if row["Score_Hooper_Moy"] >= 14 or row["Derniere_Douleur"] != "Aucune":
        color = "#FEE2E2"  # Rouge léger
    elif row["Score_Hooper_Moy"] >= 11:
        color = "#FEF3C7"  # Jaune léger
    return [f"background-color: {color}"] * len(row)

styled_table = summary_by_player.style.apply(highlight_status, axis=1)
st.dataframe(styled_table, use_container_width=True, hide_index=True)
