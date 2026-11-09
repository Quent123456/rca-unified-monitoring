"""
pages/2_Tableau_Individuel.py
Dashboard individuel connecté à Google Sheets : KPI, alertes ACWR, graphiques évolution + bien-être.
"""

import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

from modules.calculations import (
    acute_load, chronic_load, acwr, monotony, strain,
    wellness_score, compute_rolling_metrics,
)
from modules.alerts import acwr_alert, monotony_alert, wellness_alert, strain_alert
from modules.visualizations import (
    plot_load_evolution, plot_wellness_radar,
    plot_acwr_gauge, plot_wellness_history,
)

st.set_page_config(page_title="Tableau Individuel", page_icon="👤", layout="wide")

st.title("👤 Tableau de bord individuel")

# ── Connexion à Google Sheets ──────────────────────────────────────────────────
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    # Lecture de l'onglet Effectif
    players_df = conn.read(worksheet="Effectif", ttl=0)
    players_df = players_df.dropna(how="all")
    
    # Lecture de l'onglet Saisies
    sessions_df = conn.read(worksheet="Saisies", ttl=0)
    sessions_df = sessions_df.dropna(how="all")
    gsheets_ok = True
except Exception as e:
    gsheets_ok = False
    st.error(f"Erreur de connexion à Google Sheets : {e}")
    st.stop()

# ── Sélection joueur ────────────────────────────────────────────────────────────
if players_df.empty:
    st.warning("Aucun joueur. Ajoutez des joueurs dans **Gestion des Joueurs**.")
    st.stop()

# Nettoyage et formatage
players_df["prenom"] = players_df["prenom"].fillna("")
players_df["nom"] = players_df["nom"].fillna("")
players_df["poste"] = players_df["poste"].fillna("Inconnu")

player_options = {
    f"{row['prenom']} {row['nom']} — {row['poste']}": f"{row['nom']} {row['prenom']}"
    for _, row in players_df.iterrows() if str(row['nom']).strip() != ""
}

col_sel1, col_sel2 = st.columns([2, 1])
with col_sel1:
    selected_label = st.selectbox("Sélectionner un joueur", list(player_options.keys()))
with col_sel2:
    period = st.selectbox("Période d'analyse", [28, 42, 56, 84], index=0,
                          format_func=lambda x: f"{x} derniers jours")

player_gsheets_name = player_options[selected_label]
player_name = selected_label.split(" —")[0]

# ── Chargement et Nettoyage des données Saisies ─────────────────────────────────
if sessions_df.empty:
    st.info(f"Aucune séance enregistrée. Commencez par la **Saisie Quotidienne**.")
    st.stop()

# Filtrer les données pour le joueur sélectionné
df_raw = sessions_df[sessions_df["Nom / Prenom"] == player_gsheets_name].copy()

if df_raw.empty:
    st.info(f"Aucune séance enregistrée pour {player_name}. Commencez par la **Saisie Quotidienne**.")
    st.stop()

# Conversion et formatage des données brutes GSheets pour les calculs
try:
    # Traitement de l'horodateur avec format flexible (pour gérer les secondes)
    df_raw["session_date"] = pd.to_datetime(df_raw["Horodateur"], format='mixed', dayfirst=True).dt.normalize()
    
    # Sécurisation des valeurs numériques
    df_raw["foster_load"] = pd.to_numeric(df_raw["Charge (UA)"], errors='coerce').fillna(0)
    df_raw["rpe"] = pd.to_numeric(df_raw["RPE (Chiffre)"], errors='coerce').fillna(0)
    df_raw["duration"] = pd.to_numeric(df_raw["Duree (min)"], errors='coerce').fillna(0)
    
    # Extraction du bien-être depuis la chaine composite (ex: "Fatigue:3 | Courbatures:3 | Sommeil:3")
    def extract_wellness(row_text, key):
        if pd.isna(row_text): return None
        try:
            parts = str(row_text).split("|")
            for p in parts:
                if key in p:
                    return float(p.split(":")[1].strip())
        except:
            pass
        return None

    df_raw["fatigue"] = df_raw["Fatigue | Courbatures | Sommeil"].apply(lambda x: extract_wellness(x, "Fatigue"))
    df_raw["courbatures"] = df_raw["Fatigue | Courbatures | Sommeil"].apply(lambda x: extract_wellness(x, "Courbatures"))
    df_raw["sommeil"] = df_raw["Fatigue | Courbatures | Sommeil"].apply(lambda x: extract_wellness(x, "Sommeil"))

    # Tri chronologique indispensable pour les calculs de moyennes glissantes
    df_raw = df_raw.sort_values(by="session_date")
    
except Exception as e:
    st.error(f"Erreur lors du formatage des données : {e}")
    st.stop()

if df_raw["foster_load"].sum() == 0:
    st.info(f"Toutes les charges enregistrées pour {player_name} sont à zéro.")
    st.stop()

# Application des métriques roulantes
df = compute_rolling_metrics(df_raw)
df_period = df.tail(period)

# Calculs actuels (snapshot dernier jour avec données)
last = df[df["foster_load"] > 0]
loads = df["foster_load"]

a     = acute_load(loads, 7)
c     = chronic_load(loads, 28)
ratio = acwr(a, c)
mono  = monotony(loads, 7)
st_   = strain(loads, 7)

# Bien-être moyen 7 derniers jours
well_recent = df.tail(7).dropna(subset=["fatigue", "courbatures", "sommeil"])
avg_well = (
    wellness_score(
        well_recent["fatigue"].mean(),
        well_recent["courbatures"].mean(),
        well_recent["sommeil"].mean(),
    ) if not well_recent.empty else None
)

# ── ALERTES ─────────────────────────────────────────────────────────────────────
alert     = acwr_alert(ratio)
em_mono, msg_mono  = monotony_alert(mono)
em_well, msg_well  = wellness_alert(avg_well)
em_st, msg_st      = strain_alert(st_)

# Bannière d'alerte principale
if alert.level == "danger":
    st.error(f"{alert.emoji} **ALERTE RISQUE** — {alert.message}")
elif alert.level == "warning":
    st.warning(f"{alert.emoji} **Vigilance** — {alert.message}")
elif alert.level == "underload":
    st.info(f"{alert.emoji} **{alert.label}** — {alert.message}")
else:
    st.success(f"{alert.emoji} **{alert.label}** — {alert.message}")

st.markdown("---")

# ── KPI CARDS ───────────────────────────────────────────────────────────────────
st.subheader("📊 Indicateurs clés")

k1, k2, k3, k4, k5, k6 = st.columns(6)

k1.metric("⚡ Charge aiguë 7j", f"{a:.0f} UA",
          help="Charge moyenne des 7 derniers jours")
k2.metric("📈 Charge chronique 28j", f"{c:.0f} UA",
          help="Charge moyenne des 28 derniers jours")
k3.metric(
    "⚖️ ACWR",
    f"{ratio:.2f}" if ratio else "N/A",
    delta=f"{ratio - 1.0:+.2f}" if ratio else None,
    delta_color="inverse" if ratio and ratio > 1 else "normal",
    help="Optimal : 0.8 – 1.3",
)
k4.metric("🔁 Monotonie", f"{mono:.2f}" if mono else "N/A",
          help="< 1.5 : bonne variabilité | > 2.0 : ⚠️")
k5.metric("🔥 Contrainte", f"{st_:.0f} UA",
          help="Charge hebdo × Monotonie. Seuil : 6000")
k6.metric("😴 Bien-être /5", f"{avg_well:.1f}" if avg_well else "N/A",
          help="Moyenne fatigue + courbatures + sommeil (7j)")

st.markdown("---")

# ── GRAPHIQUES ──────────────────────────────────────────────────────────────────
st.subheader("📉 Évolution de la charge")

col_g1, col_g2 = st.columns([3, 1])
with col_g1:
    fig_load = plot_load_evolution(df_period, player_name)
    st.plotly_chart(fig_load, use_container_width=True)
with col_g2:
    fig_gauge = plot_acwr_gauge(ratio, player_name)
    st.plotly_chart(fig_gauge, use_container_width=True)

    # Alertes textuelles compactes
    st.markdown(f"{em_mono} {msg_mono}")
    st.markdown(f"{em_st} {msg_st}")

st.markdown("---")
st.subheader("😴 Bien-être")

col_w1, col_w2 = st.columns(2)
with col_w1:
    fig_radar = plot_wellness_radar(df_period, player_name)
    st.plotly_chart(fig_radar, use_container_width=True)
with col_w2:
    fig_hist = plot_wellness_history(df_period, player_name)
    st.plotly_chart(fig_hist, use_container_width=True)

st.markdown(f"**Bilan bien-être :** {em_well} {msg_well}")

# ── TABLEAU DES DONNÉES BRUTES ──────────────────────────────────────────────────
st.markdown("---")
with st.expander("📋 Données brutes — Sessions"):
    display_df = df_period[["session_date", "rpe", "duration", "foster_load",
                             "acute_load_7", "chronic_load_28", "acwr",
                             "fatigue", "courbatures", "sommeil"]].copy()
    display_df["session_date"] = display_df["session_date"].dt.strftime("%d/%m/%Y")
    display_df.columns = ["Date", "RPE", "Durée (min)", "Charge (UA)",
                           "Aiguë 7j", "Chronique 28j", "ACWR",
                           "Fatigue", "Courbatures", "Sommeil"]
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    # Export CSV
    csv = display_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Télécharger CSV",
        data=csv,
        file_name=f"sessions_{player_name.replace(' ', '_')}.csv",
        mime="text/csv",
    )
