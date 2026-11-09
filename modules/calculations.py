"""
modules/calculations.py
Calculs de charge d'entraînement : Foster, ACWR, monotonie, contrainte, bien-être.
"""

import numpy as np
import pandas as pd
from typing import Optional


# ── CHARGE DE FOSTER ───────────────────────────────────────────────────────────

def foster_load(rpe: float, duration_min: float) -> float:
    """
    Charge de séance selon la méthode Foster (2001).
    Charge (UA) = RPE (Borg CR-10, 0-10) × Durée (minutes)
    """
    return round(rpe * duration_min, 1)


# ── CHARGES AIGUË / CHRONIQUE ──────────────────────────────────────────────────

def acute_load(loads: pd.Series, window: int = 7) -> float:
    """
    Charge aiguë : moyenne glissante sur `window` jours (défaut 7j).
    Utilise la moyenne des N dernières valeurs disponibles.
    """
    tail = loads.tail(window)
    if tail.empty:
        return 0.0
    return round(float(tail.mean()), 1)


def chronic_load(loads: pd.Series, window: int = 28) -> float:
    """
    Charge chronique : moyenne glissante sur `window` jours (défaut 28j).
    """
    tail = loads.tail(window)
    if tail.empty:
        return 0.0
    return round(float(tail.mean()), 1)


def acwr(acute: float, chronic: float) -> Optional[float]:
    """
    Acute:Chronic Workload Ratio.
    Retourne None si la charge chronique est 0 (évite division par zéro).
    """
    if chronic == 0:
        return None
    return round(acute / chronic, 3)


# ── MONOTONIE ET CONTRAINTE ────────────────────────────────────────────────────

def monotony(loads: pd.Series, window: int = 7) -> Optional[float]:
    """
    Monotonie de Foster = moyenne(7j) / écart-type(7j).
    Valeur élevée (>2) indique un manque de variabilité → fatigue accrue.
    """
    tail = loads.tail(window)
    if len(tail) < 2:
        return None
    std = float(tail.std())
    if std == 0:
        return None  # Toutes les séances identiques → monotonie infinie (évité)
    return round(float(tail.mean()) / std, 3)


def strain(loads: pd.Series, window: int = 7) -> float:
    """
    Contrainte de Foster = charge hebdomadaire × monotonie.
    Indicateur de stress cumulé.
    """
    tail = loads.tail(window)
    weekly_load = float(tail.sum())
    mono = monotony(loads, window)
    if mono is None:
        return 0.0
    return round(weekly_load * mono, 1)


# ── ROLLING METRICS (pour graphiques temporels) ────────────────────────────────

def compute_rolling_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcule pour chaque jour les métriques glissantes sur le DataFrame complet.
    Colonnes requises : foster_load (série chronologique triée).
    Retourne un DataFrame enrichi avec :
      - acute_load_7   : moyenne glissante 7j
      - chronic_load_28: moyenne glissante 28j
      - acwr           : ratio A:C
      - monotony_7     : monotonie 7j
      - strain_7       : contrainte 7j
    """
    df = df.copy().sort_values("session_date").reset_index(drop=True)
    loads = df["foster_load"]

    df["acute_load_7"]    = loads.rolling(window=7,  min_periods=1).mean().round(1)
    df["chronic_load_28"] = loads.rolling(window=28, min_periods=7).mean().round(1)
    df["acwr"]            = (df["acute_load_7"] / df["chronic_load_28"]).replace(
                                [np.inf, -np.inf], np.nan
                            ).round(3)
    df["monotony_7"]      = (
        loads.rolling(window=7, min_periods=2).mean()
        / loads.rolling(window=7, min_periods=2).std()
    ).replace([np.inf, -np.inf], np.nan).round(3)
    df["strain_7"]        = (
        loads.rolling(window=7, min_periods=1).sum() * df["monotony_7"]
    ).round(1)

    return df


# ── BIEN-ÊTRE ──────────────────────────────────────────────────────────────────

def wellness_score(fatigue: float, courbatures: float, sommeil: float) -> float:
    """
    Score de bien-être composite (1-5).
    Note : fatigue et courbatures sont inversés (1=mauvais, 5=bon).
    Score = moyenne(fatigue, courbatures, sommeil) / 5 → normalisé 0–1.
    """
    return round((fatigue + courbatures + sommeil) / 3, 2)


def compute_team_acwr_snapshot(sessions_df: pd.DataFrame) -> pd.DataFrame:
    """
    Pour chaque joueur, calcule les métriques actuelles (snapshot du jour).
    Retourne un DataFrame résumé par joueur.
    """
    results = []
    for (player_id, player_name), group in sessions_df.groupby(
        ["player_id", "player_name"]
    ):
        group = group.sort_values("session_date")
        loads = group["foster_load"]
        a = acute_load(loads, 7)
        c = chronic_load(loads, 28)
        ratio = acwr(a, c)
        mono  = monotony(loads, 7)
        st    = strain(loads, 7)

        # Bien-être moyen sur 7 derniers jours
        well_cols = ["fatigue", "courbatures", "sommeil"]
        recent_well = group.tail(7)[well_cols].dropna()
        avg_well = (
            wellness_score(
                recent_well["fatigue"].mean(),
                recent_well["courbatures"].mean(),
                recent_well["sommeil"].mean(),
            )
            if not recent_well.empty
            else None
        )

        results.append({
            "player_id":   player_id,
            "Joueur":      player_name,
            "Poste":       group["poste"].iloc[0],
            "Charge aiguë (7j)":   a,
            "Charge chronique (28j)": c,
            "ACWR":        ratio,
            "Monotonie":   mono,
            "Contrainte":  st,
            "Bien-être /5": avg_well,
        })

    return pd.DataFrame(results)
