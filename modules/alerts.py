"""
modules/alerts.py
Règles d'alertes et niveaux de risque basés sur les indicateurs de charge.
"""

from dataclasses import dataclass
from typing import Optional


@dataclass
class AlertLevel:
    level: str        # "optimal", "warning", "danger", "underload", "unknown"
    color: str        # Couleur hex Plotly
    emoji: str
    label: str
    message: str


# ── SEUILS ACWR ────────────────────────────────────────────────────────────────

ACWR_THRESHOLDS = {
    "underload": (None, 0.8),
    "optimal":   (0.8,  1.3),
    "warning":   (1.3,  1.5),
    "danger":    (1.5,  None),
}

ACWR_ALERTS = {
    "underload": AlertLevel(
        level="underload",
        color="#3B82F6",
        emoji="🔵",
        label="Sous-charge",
        message="L'ACWR est faible. Le joueur est en sous-charge — risque de déconditionnement.",
    ),
    "optimal": AlertLevel(
        level="optimal",
        color="#22C55E",
        emoji="🟢",
        label="Zone optimale",
        message="L'ACWR est dans la zone cible. Charge bien gérée.",
    ),
    "warning": AlertLevel(
        level="warning",
        color="#F97316",
        emoji="🟠",
        label="Vigilance",
        message="L'ACWR dépasse 1.3. Surveiller les signes de fatigue accumulée.",
    ),
    "danger": AlertLevel(
        level="danger",
        color="#EF4444",
        emoji="🔴",
        label="Risque élevé",
        message="L'ACWR dépasse 1.5 — risque de blessure significativement augmenté. Réduire la charge.",
    ),
    "unknown": AlertLevel(
        level="unknown",
        color="#9CA3AF",
        emoji="⚪",
        label="Données insuffisantes",
        message="Pas assez de données (minimum 7 jours de séances).",
    ),
}


def acwr_alert(ratio: Optional[float]) -> AlertLevel:
    """Retourne le niveau d'alerte pour un ratio ACWR donné."""
    if ratio is None:
        return ACWR_ALERTS["unknown"]
    if ratio < 0.8:
        return ACWR_ALERTS["underload"]
    elif ratio < 1.3:
        return ACWR_ALERTS["optimal"]
    elif ratio < 1.5:
        return ACWR_ALERTS["warning"]
    else:
        return ACWR_ALERTS["danger"]


# ── ALERTES MONOTONIE ──────────────────────────────────────────────────────────

def monotony_alert(mono: Optional[float]) -> tuple[str, str]:
    """
    Retourne (emoji, message) pour la monotonie.
    Seuil critique : > 2.0 (Foster & Dolan, 2001).
    """
    if mono is None:
        return "⚪", "Données insuffisantes"
    if mono < 1.5:
        return "🟢", f"Monotonie faible ({mono:.2f}) — bonne variabilité des séances."
    elif mono < 2.0:
        return "🟠", f"Monotonie modérée ({mono:.2f}) — varier les intensités recommandé."
    else:
        return "🔴", f"Monotonie élevée ({mono:.2f}) ⚠️ — manque de variété, risque de surentraînement."


# ── ALERTES BIEN-ÊTRE ──────────────────────────────────────────────────────────

def wellness_alert(score: Optional[float]) -> tuple[str, str]:
    """
    Retourne (emoji, message) pour le score bien-être (1-5).
    """
    if score is None:
        return "⚪", "Aucune donnée de bien-être saisie."
    if score >= 4.0:
        return "🟢", f"Bien-être excellent ({score:.1f}/5)."
    elif score >= 3.0:
        return "🟡", f"Bien-être correct ({score:.1f}/5) — à surveiller."
    elif score >= 2.0:
        return "🟠", f"Bien-être dégradé ({score:.1f}/5) — réduire l'intensité."
    else:
        return "🔴", f"Bien-être très faible ({score:.1f}/5) ⚠️ — repos obligatoire recommandé."


# ── ALERTES CONTRAINTE ─────────────────────────────────────────────────────────

def strain_alert(st: float) -> tuple[str, str]:
    """
    Retourne (emoji, message) pour la contrainte hebdomadaire.
    Seuil critique : > 6000 UA (Foster, 1998).
    """
    if st < 3000:
        return "🟢", f"Contrainte faible ({st:.0f} UA)."
    elif st < 6000:
        return "🟠", f"Contrainte modérée ({st:.0f} UA) — surveiller la récupération."
    else:
        return "🔴", f"Contrainte élevée ({st:.0f} UA) ⚠️ — risque de surentraînement."


# ── COULEUR ACWR pour tableaux ─────────────────────────────────────────────────

def acwr_color(ratio: Optional[float]) -> str:
    """Retourne la couleur hex correspondant au niveau ACWR."""
    return acwr_alert(ratio).color
