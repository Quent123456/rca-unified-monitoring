"""Module de traitement et d'analyse des métriques GPS.

Prise en charge des métriques clés rugby :
- Distance totale (m)
- Distance à haute intensité HSR (> 20 km/h)
- Distance sprint (> 25 km/h)
- Vitesse maximale (km/h)
- Accélérations (> 3 m/s²) & Décélérations (< -3 m/s²)
- PlayerLoad (UA)
- Intensité relative (m/min)
- Calcul de l'ACWR (Acute:Chronic Workload Ratio, 7j / 28j)
"""

from __future__ import annotations

import datetime
from typing import Dict, List, Optional
import numpy as np
import pandas as pd
import streamlit as st


COLUMN_MAPPING = {
    # Variations courantes de noms de colonnes dans les exports GPS
    "player_name": ["joueur", "player", "athlete", "nom", "nom du joueur", "player name"],
    "date": ["date", "session date", "date de séance", "date séance"],
    "session_type": ["session type", "type", "type de séance", "activite"],
    "duration_min": ["duration", "durée", "duree (min)", "minutes", "duration (mins)"],
    "total_distance": ["total distance (m)", "distance totale", "distance (m)", "distance", "tot dist"],
    "hsr_distance": ["hsr", "high speed running", "distance > 20km/h", "distance hsr (m)", "hsr distance"],
    "sprint_distance": ["sprint distance", "distance sprint (m)", "sprint", "distance > 25km/h"],
    "max_speed": ["max speed (km/h)", "vitesse max", "vmax", "top speed (km/h)", "vitesse max (km/h)"],
    "accelerations": ["accel", "accelerations", "accel > 3m/s2", "acc > 3 m/s²", "nb accelerations"],
    "decelerations": ["decel", "decelerations", "decel < -3m/s2", "dec < -3 m/s²", "nb decelerations"],
    "player_load": ["player load", "playerload", "charge totale", "load", "player_load"],
}


def normalize_gps_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise les noms de colonnes et types d'un fichier GPS importé."""
    df_clean = df.copy()
    lower_cols = {str(c).strip().lower(): c for c in df_clean.columns}

    resolved_cols: Dict[str, str] = {}
    for standard_name, possible_aliases in COLUMN_MAPPING.items():
        for alias in possible_aliases:
            if alias in lower_cols:
                resolved_cols[lower_cols[alias]] = standard_name
                break

    df_clean = df_clean.rename(columns=resolved_cols)

    # Calculs dérivés automatiques si absents
    if "total_distance" in df_clean.columns and "duration_min" in df_clean.columns:
        if "intensity_m_min" not in df_clean.columns:
            dur = pd.to_numeric(df_clean["duration_min"], errors="coerce").fillna(0)
            dist = pd.to_numeric(df_clean["total_distance"], errors="coerce").fillna(0)
            df_clean["intensity_m_min"] = np.where(dur > 0, dist / dur, 0).round(1)

    return df_clean


def get_mock_gps_data() -> pd.DataFrame:
    """Génère un jeu de données GPS simulé pour les tests et démonstrations."""
    if "mock_gps_data" in st.session_state:
        return st.session_state["mock_gps_data"]

    from modules.gsheet_sync import get_players

    players = get_players()
    today = datetime.date.today()
    records = []

    np.random.seed(42)

    for day_offset in range(28, -1, -1):
        cur_date = today - datetime.timedelta(days=day_offset)
        # Match le samedi (5), Entraînements mardi (1), jeudi (3), vendredi (4)
        if cur_date.weekday() not in [1, 3, 4, 5]:
            continue

        is_match = cur_date.weekday() == 5
        session_type = "Match" if is_match else ("Entraînement Collectif" if cur_date.weekday() in [1, 3] else "Capitaine / Vitesse")
        duration = 80 if is_match else (75 if cur_date.weekday() in [1, 3] else 45)

        for _, player in players.iterrows():
            is_back = player["groupe"] == "Arrières"

            # Les arrières parcourent généralement plus de distance et de HSR
            base_dist = (6500 if is_match else 4800) if is_back else (5800 if is_match else 4200)
            dist = int(np.random.normal(base_dist, 400))

            base_hsr = (550 if is_match else 350) if is_back else (250 if is_match else 140)
            hsr = max(0, int(np.random.normal(base_hsr, 60)))

            base_sprint = (120 if is_match else 70) if is_back else (40 if is_match else 15)
            sprint = max(0, int(np.random.normal(base_sprint, 25)))

            vmax = round(float(np.random.normal(31.5 if is_back else 27.5, 1.2)), 1)
            accel = int(np.random.normal(18 if is_back else 12, 4))
            decel = int(np.random.normal(16 if is_back else 14, 4))
            player_load = round(dist * 0.11 + accel * 2.5 + np.random.normal(0, 15), 1)

            records.append({
                "date": cur_date.isoformat(),
                "player_name": f"{player['prenom']} {player['nom']}",
                "joueur_id": player["id"],
                "groupe": player["groupe"],
                "poste": player["poste"],
                "session_type": session_type,
                "duration_min": duration,
                "total_distance": dist,
                "hsr_distance": hsr,
                "sprint_distance": sprint,
                "max_speed": vmax,
                "accelerations": max(0, accel),
                "decelerations": max(0, decel),
                "player_load": player_load,
                "intensity_m_min": round(dist / duration, 1) if duration > 0 else 0
            })

    df = pd.DataFrame(records)
    st.session_state["mock_gps_data"] = df
    return df


def calculate_acwr(df: pd.DataFrame, metric: str = "total_distance") -> pd.DataFrame:
    """Calcule le ratio charge aiguë / chronique (ACWR) par joueur sur base glissante (7j / 28j)."""
    if df.empty or metric not in df.columns:
        return pd.DataFrame()

    df_calc = df.copy()
    df_calc["date"] = pd.to_datetime(df_calc["date"])
    df_calc = df_calc.sort_values(by=["player_name", "date"])

    results = []
    for player, group in df_calc.groupby("player_name"):
        # Rééchantillonnage journalier pour avoir une vraie série temporelle
        daily = group.set_index("date").resample("D")[metric].sum().fillna(0)

        acute = daily.rolling(window=7, min_periods=1).mean()
        chronic = daily.rolling(window=28, min_periods=1).mean()
        acwr = acute / chronic.replace(0, np.nan)

        res_df = pd.DataFrame({
            "player_name": player,
            "date": daily.index,
            "daily_load": daily.values,
            "acute_load_7d": acute.values,
            "chronic_load_28d": chronic.values,
            "acwr": acwr.values.round(2)
        })
        results.append(res_df)

    if not results:
        return pd.DataFrame()

    return pd.concat(results, ignore_index=True)
