"""
modules/visualizations.py
Graphiques Plotly réutilisables pour les dashboards individuel et collectif.
"""

import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import pandas as pd
import numpy as np
from typing import Optional


# ── COULEURS ───────────────────────────────────────────────────────────────────
COLOR_ACUTE   = "#6366F1"   # Violet — charge aiguë
COLOR_CHRONIC = "#10B981"   # Vert émeraude — charge chronique
COLOR_LOAD    = "#94A3B8"   # Gris bleu — barres charge quotidienne
COLOR_ACWR    = "#F59E0B"   # Ambre — courbe ACWR

ZONE_COLORS = {
    "underload": "rgba(59,130,246,0.10)",
    "optimal":   "rgba(34,197,94,0.10)",
    "warning":   "rgba(249,115,22,0.10)",
    "danger":    "rgba(239,68,68,0.10)",
}


# ── 1. ÉVOLUTION CHARGE AIGUË / CHRONIQUE / ACWR ──────────────────────────────

def plot_load_evolution(df: pd.DataFrame, player_name: str = "") -> go.Figure:
    """
    Graphique combiné :
    - Barres : charge quotidienne Foster
    - Courbes : charge aiguë 7j et chronique 28j
    - Axe secondaire : ACWR avec zones colorées
    """
    fig = make_subplots(
        rows=2, cols=1,
        shared_xaxes=True,
        row_heights=[0.65, 0.35],
        vertical_spacing=0.06,
        subplot_titles=("Charge quotidienne & moyennes glissantes", "ACWR (Ratio Aiguë:Chronique)"),
    )

    # — Barres charge quotidienne
    fig.add_trace(
        go.Bar(
            x=df["session_date"],
            y=df["foster_load"],
            name="Charge séance (UA)",
            marker_color=COLOR_LOAD,
            opacity=0.7,
        ),
        row=1, col=1,
    )

    # — Charge aiguë 7j
    fig.add_trace(
        go.Scatter(
            x=df["session_date"],
            y=df["acute_load_7"],
            name="Charge aiguë 7j",
            line=dict(color=COLOR_ACUTE, width=2.5),
            mode="lines",
        ),
        row=1, col=1,
    )

    # — Charge chronique 28j
    fig.add_trace(
        go.Scatter(
            x=df["session_date"],
            y=df["chronic_load_28"],
            name="Charge chronique 28j",
            line=dict(color=COLOR_CHRONIC, width=2.5, dash="dash"),
            mode="lines",
        ),
        row=1, col=1,
    )

    # — ACWR
    fig.add_trace(
        go.Scatter(
            x=df["session_date"],
            y=df["acwr"],
            name="ACWR",
            line=dict(color=COLOR_ACWR, width=2),
            mode="lines+markers",
            marker=dict(size=4),
        ),
        row=2, col=1,
    )

    # Zones ACWR
    x_range = [df["session_date"].min(), df["session_date"].max()]
    zones = [
        (0,   0.8,  "rgba(59,130,246,0.12)",  "Sous-charge"),
        (0.8, 1.3,  "rgba(34,197,94,0.12)",   "Optimal"),
        (1.3, 1.5,  "rgba(249,115,22,0.12)",  "Vigilance"),
        (1.5, 2.5,  "rgba(239,68,68,0.12)",   "Danger"),
    ]
    for y0, y1, color, label in zones:
        fig.add_hrect(
            y0=y0, y1=y1,
            fillcolor=color,
            line_width=0,
            row=2, col=1,
        )

    # Ligne ACWR = 1.5
    fig.add_hline(y=1.5, line_dash="dot", line_color="#EF4444", line_width=1, row=2, col=1)
    fig.add_hline(y=1.3, line_dash="dot", line_color="#F97316", line_width=1, row=2, col=1)
    fig.add_hline(y=0.8, line_dash="dot", line_color="#3B82F6", line_width=1, row=2, col=1)

    title = f"Évolution de la charge — {player_name}" if player_name else "Évolution de la charge"
    fig.update_layout(
        title=title,
        height=520,
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        plot_bgcolor="white",
        paper_bgcolor="white",
    )
    fig.update_yaxes(title_text="Charge (UA)", row=1, col=1)
    fig.update_yaxes(title_text="ACWR", range=[0, 2.5], row=2, col=1)
    fig.update_xaxes(showgrid=True, gridcolor="#F1F5F9", row=2, col=1)
    fig.update_yaxes(showgrid=True, gridcolor="#F1F5F9")

    return fig


# ── 2. RADAR BIEN-ÊTRE ─────────────────────────────────────────────────────────

def plot_wellness_radar(df: pd.DataFrame, player_name: str = "") -> go.Figure:
    """
    Graphique radar des indicateurs de bien-être des 7 derniers jours.
    """
    recent = df.tail(7).dropna(subset=["fatigue", "courbatures", "sommeil"])
    if recent.empty:
        fig = go.Figure()
        fig.add_annotation(text="Aucune donnée de bien-être disponible",
                           xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        return fig

    avg_fatigue     = recent["fatigue"].mean()
    avg_courbatures = recent["courbatures"].mean()
    avg_sommeil     = recent["sommeil"].mean()

    categories = ["Fatigue", "Courbatures", "Sommeil", "Fatigue"]  # Fermeture radar
    values     = [avg_fatigue, avg_courbatures, avg_sommeil, avg_fatigue]

    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=values,
        theta=categories,
        fill="toself",
        name="Bien-être moyen (7j)",
        line_color=COLOR_ACUTE,
        fillcolor="rgba(99,102,241,0.2)",
    ))
    # Zone de référence optimale (5/5)
    fig.add_trace(go.Scatterpolar(
        r=[5, 5, 5, 5],
        theta=categories,
        fill="toself",
        name="Référence (5/5)",
        line_color="#22C55E",
        fillcolor="rgba(34,197,94,0.05)",
        line_dash="dash",
    ))

    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
        title=f"Bien-être moyen 7j — {player_name}" if player_name else "Bien-être moyen 7j",
        height=380,
        showlegend=True,
        paper_bgcolor="white",
    )
    return fig


# ── 3. JAUGE ACWR ─────────────────────────────────────────────────────────────

def plot_acwr_gauge(ratio: Optional[float], player_name: str = "") -> go.Figure:
    """Jauge circulaire colorée pour l'ACWR actuel."""
    value = ratio if ratio is not None else 0
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=value,
        number={"suffix": "", "font": {"size": 40}},
        delta={"reference": 1.0, "increasing": {"color": "#EF4444"}, "decreasing": {"color": "#22C55E"}},
        title={"text": f"ACWR<br><sub>{player_name}</sub>", "font": {"size": 16}},
        gauge={
            "axis": {"range": [0, 2.5], "tickwidth": 1},
            "bar": {"color": "#6366F1", "thickness": 0.25},
            "steps": [
                {"range": [0, 0.8],   "color": "rgba(59,130,246,0.2)"},
                {"range": [0.8, 1.3], "color": "rgba(34,197,94,0.2)"},
                {"range": [1.3, 1.5], "color": "rgba(249,115,22,0.2)"},
                {"range": [1.5, 2.5], "color": "rgba(239,68,68,0.2)"},
            ],
            "threshold": {
                "line": {"color": "#EF4444", "width": 3},
                "thickness": 0.8,
                "value": 1.5,
            },
        },
    ))
    fig.update_layout(height=280, paper_bgcolor="white", margin=dict(t=60, b=20, l=20, r=20))
    return fig


# ── 4. HEATMAP ÉQUIPE ACWR ────────────────────────────────────────────────────

def plot_team_heatmap(snapshots_df: pd.DataFrame, sessions_all: pd.DataFrame) -> go.Figure:
    """
    Heatmap ACWR : lignes = joueurs, colonnes = 14 derniers jours.
    Couleurs : bleu < vert < orange < rouge selon les zones ACWR.
    """
    if sessions_all.empty:
        fig = go.Figure()
        fig.add_annotation(text="Aucune donnée disponible", xref="paper", yref="paper",
                           x=0.5, y=0.5, showarrow=False)
        return fig

    from modules.calculations import compute_rolling_metrics

    # Construire la matrice joueurs × jours
    last_14 = pd.date_range(end=pd.Timestamp.today().normalize(), periods=14, freq="D")
    players  = sessions_all[["player_id", "player_name"]].drop_duplicates()

    matrix_acwr = pd.DataFrame(index=players["player_name"].values, columns=last_14)
    matrix_acwr[:] = np.nan

    for _, row in players.iterrows():
        pid  = row["player_id"]
        pname = row["player_name"]
        pdata = sessions_all[sessions_all["player_id"] == pid].copy()
        pdata = compute_rolling_metrics(pdata)
        for day in last_14:
            day_row = pdata[pdata["session_date"] == day]
            if not day_row.empty:
                matrix_acwr.at[pname, day] = day_row["acwr"].values[0]

    matrix_acwr = matrix_acwr.astype(float)

    colorscale = [
        [0.0,  "#3B82F6"],   # Bleu : sous-charge
        [0.32, "#22C55E"],   # Vert : optimal bas
        [0.52, "#22C55E"],   # Vert : optimal haut
        [0.60, "#F97316"],   # Orange : vigilance
        [0.70, "#EF4444"],   # Rouge : danger
        [1.0,  "#7F1D1D"],   # Rouge foncé : extrême
    ]

    fig = go.Figure(go.Heatmap(
        z=matrix_acwr.values,
        x=[d.strftime("%d/%m") for d in last_14],
        y=matrix_acwr.index.tolist(),
        colorscale=colorscale,
        zmin=0,
        zmax=2.0,
        text=np.where(
            np.isnan(matrix_acwr.values),
            "Repos",
            np.round(matrix_acwr.values, 2).astype(str)
        ),
        texttemplate="%{text}",
        hovertemplate="Joueur: %{y}<br>Date: %{x}<br>ACWR: %{z:.2f}<extra></extra>",
        colorbar=dict(
            title="ACWR",
            tickvals=[0.4, 0.8, 1.3, 1.5, 2.0],
            ticktext=["0.4", "0.8 (min)", "1.3 (opt)", "1.5 (⚠️)", "2.0"],
        ),
    ))

    fig.update_layout(
        title="Heatmap ACWR — Effectif (14 derniers jours)",
        height=max(350, 50 * len(players) + 100),
        xaxis=dict(side="top"),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=150, r=50, t=80, b=40),
    )
    return fig


# ── 5. HISTORIQUE BIEN-ÊTRE (LIGNES INDIVIDUELLES) ───────────────────────────

def plot_wellness_history(df: pd.DataFrame, player_name: str = "") -> go.Figure:
    """Courbes temporelles fatigue / courbatures / sommeil."""
    fig = go.Figure()

    colors = {"fatigue": "#EF4444", "courbatures": "#F97316", "sommeil": "#6366F1"}
    labels = {"fatigue": "Fatigue", "courbatures": "Courbatures", "sommeil": "Sommeil"}

    for col, color in colors.items():
        subset = df.dropna(subset=[col])
        if subset.empty:
            continue
        fig.add_trace(go.Scatter(
            x=subset["session_date"],
            y=subset[col],
            name=labels[col],
            line=dict(color=color, width=2),
            mode="lines+markers",
            marker=dict(size=5),
        ))

    fig.add_hrect(y0=1, y1=2.5, fillcolor="rgba(239,68,68,0.08)", line_width=0,
                  annotation_text="Zone d'alerte", annotation_position="left")
    fig.update_layout(
        title=f"Historique bien-être — {player_name}" if player_name else "Historique bien-être",
        yaxis=dict(title="Score (1=mauvais, 5=excellent)", range=[0.5, 5.5]),
        xaxis=dict(title="Date"),
        height=350,
        hovermode="x unified",
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return fig
