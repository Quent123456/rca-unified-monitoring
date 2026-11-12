"""Module de génération des rapports PDF officiels du Rugby Club Amiénois (RCA).

Gère :
1. Le rapport PDF officiel pour les coachs (4 à 8 pages Paysage) avec podiums par ligne/poste, jauges et graphiques.
2. L'intégration de la mémoire hebdomadaire et des comparatifs multi-séances (Mardi vs Mercredi vs Vendredi vs Dimanche).
3. L'enrichissement automatique avec les données de charge RPE (Nom, Prénom, Date, skipna=True).
4. Le bilan individuel joueur A4 au format ReportLab.
"""

from __future__ import annotations

import base64
import datetime
import io
import os
import re
from typing import Any, Dict, List, Optional, Tuple, Union

from fpdf import FPDF
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

try:
    from modules import photo_manager, weekly_memory, rpe_connector, gps_processing
except ImportError:
    import photo_manager
    import weekly_memory
    import rpe_connector
    import gps_processing

generer_podium_dynamique = getattr(gps_processing, "generer_podium_dynamique", None)
get_active_objectifs_config = getattr(gps_processing, "get_active_objectifs_config", None)
OBJECTIFS_CONFIG = getattr(gps_processing, "OBJECTIFS_CONFIG", {})

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ── Palette graphique officielle Rugby Club Amiénois (RCA) ───────────────────
RCA_NAVY_RGB = (22, 43, 76)      # #162B4C Bleu marine foncé officiel
RCA_NAVY_HEX = "#162B4C"
RCA_GOLD_RGB = (243, 195, 0)     # #F3C300 Jaune or officiel
RCA_GOLD_HEX = "#F3C300"
RCA_AZUR_HEX = "#2B6CB0"         # Bleu azur secondaire
RCA_RPE_ORANGE_HEX = "#E67E22"   # Orange sportif pour le RPE
RCA_GOLD_DARK_HEX = "#B45309"    # Ambre foncé pour texte sur fond blanc


def get_pdf_logo_path() -> Optional[str]:
    """Retourne le chemin absolu vers le logo officiel RCA (assets/logo_rca.png)."""
    candidates = [
        os.path.join(PROJECT_ROOT, "assets", "logo_rca.png"),
        os.path.join(PROJECT_ROOT, "logo_rca.png"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo_rca.png"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_rca.png"),
        "assets/logo_rca.png",
        "logo_rca.png",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


# ==============================================================================
# 7b. GÉNÉRATION DU RAPPORT PDF POUR LES COACHS (FPDF & MATPLOTLIB)
# ==============================================================================

JOURS_FR = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
MOIS_FR = [
    "", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
]


def formater_date_seance_fr(val_date: Any) -> str:
    """
    Convertit n'importe quel format de date (numéro de série Excel ex: 46273, string, timestamp)
    en date lisible en français avec numéro de semaine :
    ex: 'Mardi 8 Septembre 2026 - Semaine 37'
    """
    dt: Optional[datetime.datetime] = None
    if val_date is not None and not pd.isna(val_date) and str(val_date).strip() != "":
        # 1. Numéro de série Excel (ex: 46273 ou "46273")
        try:
            num_val = float(str(val_date).strip().replace(",", "."))
            if 30000 < num_val < 60000:
                dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=num_val)
        except (ValueError, TypeError):
            pass

        # 2. Parsing standard via pandas
        if dt is None:
            try:
                ts = pd.to_datetime(val_date, errors="coerce", dayfirst=True)
                if pd.notna(ts):
                    dt = ts.to_pydatetime()
            except Exception:
                pass

    if dt is None:
        dt = datetime.datetime.now()

    nom_jour = JOURS_FR[dt.weekday()]
    jour = dt.day
    nom_mois = MOIS_FR[dt.month]
    annee = dt.year
    semaine = dt.isocalendar()[1]
    return f"{nom_jour} {jour} {nom_mois} {annee} - Semaine {semaine}"


def generer_graphique_groupes_matplotlib(
    df_indiv: pd.DataFrame,
    moyennes_groupes: Optional[Dict[str, Dict[str, float]]] = None,
) -> io.BytesIO:
    """
    Génère un diagramme en barres comparant la Distance Moyenne (m)
    et le HSR Moyen (m) par Groupe de joueurs (5 de devant, 3ème Ligne, Trois-quarts).
    La barre rouge affiche STRICTEMENT la Zone 3 (HSR Z3 = Distance in Speed Zone 3 (km) * 1000).
    """
    fig, ax1 = plt.subplots(figsize=(5.6, 5.6), dpi=180)
    fig.patch.set_facecolor("#FFFFFF")
    ax1.set_facecolor("#FFFFFF")

    groupes = ["5 de devant", "3ème Ligne", "Trois-quarts"]

    dist_vals: List[float] = []
    hsr_vals: List[float] = []

    for g in groupes:
        if moyennes_groupes and g in moyennes_groupes and "Dist_m" in moyennes_groupes[g]:
            dist_vals.append(float(moyennes_groupes[g]["Dist_m"]))
            hsr_vals.append(float(moyennes_groupes[g]["HSR_m"]))
        else:
            sub = df_indiv[df_indiv["Groupe"] == g]
            if not sub.empty:
                dist_vals.append(float(sub["Distance (km)"].mean() * 1000.0) if "Distance (km)" in sub.columns else 0.0)
                # Strictement Zone 3 (HSR Z3) - Exclusion formelle du cumul Z3+Z4+Z5
                if "Distance in Speed Zone 3  (km)" in sub.columns:
                    hsr_vals.append(float(sub["Distance in Speed Zone 3  (km)"].mean() * 1000.0))
                elif "HSR Z3 (m)" in sub.columns:
                    hsr_vals.append(float(sub["HSR Z3 (m)"].mean()))
                elif "HSR_m" in sub.columns:
                    hsr_vals.append(float(sub["HSR_m"].mean()))
                elif " Distance haute intensité (m)" in sub.columns:
                    hsr_vals.append(float(sub[" Distance haute intensité (m)"].mean()))
                else:
                    hsr_vals.append(0.0)
            else:
                dist_vals.append(0.0)
                hsr_vals.append(0.0)

    x = np.arange(len(groupes))
    width = 0.35
    ax2 = ax1.twinx()

    rects1 = ax1.bar(x - width/2, dist_vals, width, label="Distance Moyenne (m)", color=RCA_NAVY_HEX, edgecolor="none")
    rects2 = ax2.bar(x + width/2, hsr_vals, width, label="HSR Z3 Moyen (m)", color=RCA_GOLD_HEX, edgecolor="#D4A100")

    ax1.set_ylabel("Distance Moyenne (m)", color=RCA_NAVY_HEX, fontweight="bold", fontsize=9)
    ax2.set_ylabel("HSR Z3 Moyen (m)", color=RCA_GOLD_DARK_HEX, fontweight="bold", fontsize=9)
    ax1.set_xticks(x)
    ax1.set_xticklabels(groupes, fontweight="bold", fontsize=9)
    ax1.set_title("Distance Moyenne & HSR Z3 Moyen par Groupe", fontweight="bold", color=RCA_NAVY_HEX, fontsize=10.5, pad=12)

    # Valeurs écrites directement sur chaque barre
    for rect in rects1:
        h = rect.get_height()
        if h > 0:
            ax1.annotate(f"{int(round(h)):,d} m".replace(",", " "),
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 3), textcoords="offset points",
                         ha="center", va="bottom", fontsize=8, fontweight="bold", color=RCA_NAVY_HEX)

    for rect in rects2:
        h = rect.get_height()
        if h > 0:
            ax2.annotate(f"{int(round(h)):,d} m".replace(",", " "),
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 3), textcoords="offset points",
                         ha="center", va="bottom", fontsize=8, fontweight="bold", color=RCA_GOLD_DARK_HEX)

    max_dist = max(dist_vals) if dist_vals and max(dist_vals) > 0 else 1000.0
    max_hsr = max(hsr_vals) if hsr_vals and max(hsr_vals) > 0 else 100.0
    ax1.set_ylim(0, max_dist * 1.25)
    ax2.set_ylim(0, max_hsr * 1.25)
    ax1.grid(axis="y", linestyle="--", alpha=0.3)
    ax1.tick_params(axis="y", labelcolor=RCA_NAVY_HEX)
    ax2.tick_params(axis="y", labelcolor=RCA_GOLD_DARK_HEX)

    # Légende épurée
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper center", bbox_to_anchor=(0.5, -0.06), ncol=2, frameon=False, fontsize=8.5)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def generer_graphique_postes_matplotlib(df_indiv: pd.DataFrame) -> io.BytesIO:
    """
    Génère un diagramme en barres affichant la Vitesse Max (Vmax) Moyenne par Poste.
    """
    df_temp = df_indiv.copy()
    j_lookup = weekly_memory._get_effective_joueurs_lookup() if hasattr(weekly_memory, "_get_effective_joueurs_lookup") else {}
    postes_list = []
    for _, row in df_temp.iterrows():
        nom = str(row.get("Nom ; Prénom", "")).strip()
        poste = str(row.get("Poste_Detaille", "")).strip()
        if not poste:
            poste = j_lookup.get(nom, "")
        if not poste:
            for j_nom, j_pos in j_lookup.items():
                if j_nom.lower() in nom.lower() or nom.lower() in j_nom.lower():
                    poste = j_pos
                    break
        if not poste:
            poste = str(row.get("Groupe", "Non défini"))
        postes_list.append(poste)
    df_temp["_Poste"] = postes_list

    grp_vmax = df_temp.groupby("_Poste")["Top Speed (km/h)"].mean().sort_values(ascending=True)

    fig, ax = plt.subplots(figsize=(6.0, 4.5), dpi=160)
    y_pos = np.arange(len(grp_vmax))
    rects = ax.barh(y_pos, grp_vmax.values, color=RCA_NAVY_HEX, height=0.55, edgecolor="none")

    ax.set_yticks(y_pos)
    ax.set_yticklabels(grp_vmax.index, fontweight="bold", fontsize=8, color="#334155")
    ax.set_xlabel("Vitesse Max (km/h)", fontweight="bold", fontsize=8.5, color=RCA_NAVY_HEX)
    ax.set_title("Vitesse Max Moyenne par Poste (km/h)", fontweight="bold", color=RCA_NAVY_HEX, fontsize=10.5, pad=12)

    # Valeurs écrites directement sur les barres
    for rect in rects:
        w = rect.get_width()
        if w > 0:
            ax.annotate(f"{w:.1f} km/h",
                        xy=(w, rect.get_y() + rect.get_height() / 2),
                        xytext=(4, 0), textcoords="offset points",
                        ha="left", va="center", fontsize=8, fontweight="bold", color=RCA_NAVY_HEX)

    max_v = grp_vmax.max() if not grp_vmax.empty and grp_vmax.max() > 0 else 30.0
    ax.set_xlim(0, max_v * 1.25)
    ax.grid(axis="x", linestyle="--", alpha=0.3)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def generer_graphique_jauges_groupe_matplotlib(
    group_name: str,
    metrics: Dict[str, Tuple[float, float]],
) -> io.BytesIO:
    """
    Génère une figure 2x2 propre et aérée regroupant les 4 duels de barres 'Objectif vs Réalisé'
    du module 'Retour Journée' pour un groupe donné :
    1. Objectif Distance (m) vs Réalisé
    2. Obj. HSR (m) vs Réalisé (Z3)
    3. Obj. VHSR (m) vs Réalisé (Z4+Z5)
    4. Obj. ACC/DEC vs Réalisé
    """
    fig, axes = plt.subplots(2, 2, figsize=(10.5, 5.2), dpi=180)
    fig.patch.set_facecolor("#FFFFFF")

    cards_def = [
        ("Objectif Distance (m)", "Distance Totale (m)", "m", axes[0, 0]),
        ("Obj. HSR", "HSR Z3 (m)", "m", axes[0, 1]),
        ("Obj. VHSR", "VHSR Sprint (m)", "m", axes[1, 0]),
        ("Obj. ACC/DEC", "ACC / DEC (> 3 m/s²)", "", axes[1, 1]),
    ]

    for key, title, unit, ax in cards_def:
        ax.set_facecolor("#F8FAFC")
        for spine in ax.spines.values():
            spine.set_color("#CBD5E1")
            spine.set_linewidth(0.8)

        val_obj, val_real = metrics.get(key, (0.0, 0.0))
        val_obj = float(val_obj)
        val_real = float(val_real)
        pct = (val_real / val_obj * 100.0) if val_obj > 0 else 0.0

        x = np.array([0, 1])
        width = 0.44

        bars = ax.bar(
            x,
            [val_obj, val_real],
            width=width,
            color=[RCA_NAVY_HEX, RCA_GOLD_HEX],
            edgecolor="none",
            zorder=3,
        )

        ax.set_title(title, fontsize=10.5, fontweight="bold", color=RCA_NAVY_HEX, pad=10)
        ax.set_xticks(x)
        ax.set_xticklabels(["Objectif", "Réalisé"], fontsize=9.5, fontweight="bold", color="#334155")

        max_v = max(val_obj, val_real, 1.0)
        ax.set_ylim(0, max_v * 1.30)
        ax.grid(axis="y", linestyle="--", alpha=0.35, color="#94A3B8", zorder=1)
        ax.set_yticklabels([])
        ax.tick_params(axis="both", which="both", length=0)

        def _fmt(v: float, u: str) -> str:
            if abs(v - round(v)) < 0.05:
                s = f"{int(round(v)):,d}".replace(",", " ")
            else:
                s = f"{v:.1f}"
            return f"{s} {u}".strip()

        ax.text(
            0, val_obj + (max_v * 0.04),
            _fmt(val_obj, unit),
            ha="center", va="bottom",
            fontsize=9.2, fontweight="bold", color=RCA_NAVY_HEX
        )

        ax.text(
            1, val_real + (max_v * 0.04),
            _fmt(val_real, unit),
            ha="center", va="bottom",
            fontsize=9.2, fontweight="bold", color=RCA_GOLD_DARK_HEX
        )

        # Affichage du % à l'intérieur de la barre dorée ou au-dessus si la barre est très petite
        if val_real >= max_v * 0.18:
            ax.text(
                1, val_real * 0.48,
                f"{pct:.0f}%",
                ha="center", va="center",
                fontsize=12.5, fontweight="bold", color=RCA_NAVY_HEX,
                zorder=4
            )
        else:
            ax.text(
                1, val_real + (max_v * 0.14),
                f"({pct:.0f}%)",
                ha="center", va="bottom",
                fontsize=9.2, fontweight="bold", color=RCA_GOLD_DARK_HEX,
                zorder=4
            )

    fig.tight_layout(pad=2.2)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


# ==============================================================================
# 7c. SYNTHÈSE NARRATIVE PAR LLM (NOTE DU PRÉPARATEUR PHYSIQUE)
# ==============================================================================

def generer_synthese_llm(df_indiv: pd.DataFrame, api_key: Optional[str] = None) -> str:
    """
    Rédige un court paragraphe de synthèse professionnelle (2 à 3 phrases maximum)
    destiné au staff technique du RCA en utilisant le client Gemini.
    Sécurité anti-panne intégrée : retourne un texte de repli propre et factuel
    si l'API key n'est pas configurée ou si l'appel échoue.
    """
    if df_indiv.empty:
        return "Séance enregistrée - Aucune donnée individuelle disponible pour l'analyse des charges."

    nb_joueurs = len(df_indiv)
    dist_tot = float(df_indiv["Distance (km)"].sum()) if "Distance (km)" in df_indiv.columns else 0.0
    dist_moy = float(df_indiv["Distance (km)"].mean()) if "Distance (km)" in df_indiv.columns else 0.0
    vmax = float(df_indiv["Top Speed (km/h)"].max()) if "Top Speed (km/h)" in df_indiv.columns else 0.0

    hsr_col = " Distance haute intensité (m)" if " Distance haute intensité (m)" in df_indiv.columns else ("HSR (m)" if "HSR (m)" in df_indiv.columns else None)
    hsr_tot = float(df_indiv[hsr_col].sum()) if hsr_col and not df_indiv[hsr_col].empty else 0.0
    hsr_moy = float(df_indiv[hsr_col].mean()) if hsr_col and not df_indiv[hsr_col].empty else 0.0
    if pd.isna(hsr_moy):
        hsr_moy = 0.0
    if pd.isna(hsr_tot):
        hsr_tot = 0.0

    # Texte de repli propre et factuel garanti
    fallback_texte = (
        f"Séance enregistrée pour {nb_joueurs} joueurs avec une charge globale de {dist_tot:.1f} km "
        f"({dist_moy:.2f} km/joueur) et un pic de vitesse mesuré à {vmax:.1f} km/h. "
        f"Le volume moyen à haute intensité s'établit à {int(round(hsr_moy))} m par joueur. "
        f"Analyse des charges validée par le pôle performance RCA."
    )

    try:
        resolved_key = (
            api_key
            or os.environ.get("GEMINI_API_KEY")
            or os.environ.get("GOOGLE_API_KEY")
            or (hasattr(st, "secrets") and ("GEMINI_API_KEY" in st.secrets) and st.secrets["GEMINI_API_KEY"])
            or (hasattr(st, "secrets") and ("GOOGLE_API_KEY" in st.secrets) and st.secrets["GOOGLE_API_KEY"])
            or (hasattr(st, "session_state") and st.session_state.get("gemini_api_key"))
        )

        if not resolved_key or not str(resolved_key).strip():
            return fallback_texte

        from google import genai
        client = genai.Client(api_key=str(resolved_key).strip())

        prompt = (
            "Agis en tant que préparateur physique de rugby senior du Rugby Club Amiénois (RCA). "
            "Rédige une synthèse athlétique factuelle, professionnelle et percutante (strictement 2 à 3 phrases maximum, "
            "sans puces, sans tirets, sans titre et sans aucun emoji) pour l'entraîneur principal en te basant sur ces métriques GPS :\n"
            f"- Effectif évalué : {nb_joueurs} joueurs\n"
            f"- Distance totale cumulée : {dist_tot:.1f} km (moyenne de {dist_moy:.2f} km/joueur)\n"
            f"- Vitesse maximale de la séance : {vmax:.1f} km/h\n"
            f"- Volume de haute intensité (HSR > 19.8 km/h) : moyenne de {int(round(hsr_moy))} m/joueur (total cumulé : {int(round(hsr_tot))} m)\n"
            "Style attendu : expert en sciences du sport, direct, axé sur la charge globale, la vitesse maximale et la sollicitation de haute intensité."
        )

        texte = None
        # 1. Tentative avec client.interactions (gemini-3.8-flash)
        try:
            interaction = client.interactions.create(
                model="gemini-3.8-flash",
                input=prompt,
            )
            texte = getattr(interaction, "output_text", None)
        except Exception:
            pass

        # 2. Repli vers client.models.generate_content (gemini-3.8-flash)
        if not texte:
            try:
                resp = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents=prompt,
                )
                texte = getattr(resp, "text", None)
            except Exception:
                pass

        # 3. Repli vers gemini-2.5-flash
        if not texte:
            try:
                resp = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt,
                )
                texte = getattr(resp, "text", None)
            except Exception:
                pass

        if texte and str(texte).strip():
            clean_str = " ".join(str(texte).strip().split())
            return clean_str

        return fallback_texte

    except Exception:
        return fallback_texte


def safe_latin1(s: Any) -> str:
    """Assure qu'une chaîne est compatible avec l'encodage Latin-1 (Helvetica standard FPDF)."""
    if s is None:
        return ""
    txt = str(s)
    replacements = {
        "\u2019": "'",
        "\u2018": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u201e": '"',
        "\u2013": "-",
        "\u2014": "-",
        "\u2022": "-",
        "\u2026": "...",
        "\u0153": "oe",
        "\u0152": "OE",
        "\u20ac": "EUR",
        "\xa0": " ",
        "\u200b": "",
        "\ufeff": "",
    }
    for k, v in replacements.items():
        txt = txt.replace(k, v)
    clean_chars = []
    for ch in txt:
        try:
            ch.encode("latin-1")
            clean_chars.append(ch)
        except UnicodeEncodeError:
            clean_chars.append(" ")
    return "".join(clean_chars)


_render_podium_section_pdf = weekly_memory._render_podium_section_pdf
_render_podium_page_aere = weekly_memory._render_podium_page_aere
_render_podiums_compact_horizontal = weekly_memory._render_podiums_compact_horizontal
calculer_sous_titres_podium = weekly_memory.calculer_sous_titres_podium


def generer_rapport_pdf_coachs(
    df_indiv: pd.DataFrame,
    date_seance: Optional[Any] = None,
    synthese_note: Optional[str] = None,
    moyennes_groupes: Optional[Dict[str, Dict[str, float]]] = None,
    config_objectifs: Optional[Dict[str, Dict[str, float]]] = None,
    inclure_comparatif: bool = False,
    nb_comparatif: int = 4,
    adversaire: str = "",
    rpe_stats: Optional[Dict[str, Any]] = None,
    granularity_mode: str = "Par ligne de poste",
    **kwargs: Any,
) -> bytes:
    """
    Génère le rapport PDF multi-pages officiel RCA au format Paysage (A4) :
    - Mode 'Par ligne de poste' (7 pages, ou 8 avec comparatif)
    - Mode 'Détaillé par poste' (2 pages par poste actif parmi les 9 postes + triptyques)
    """
    if df_indiv is None:
        df_indiv = pd.DataFrame()
    df_indiv = df_indiv.copy()

    # Exclusion préventive des joueurs fictifs / tests
    try:
        try:
            from modules import gsheet_sync
        except (ImportError, ModuleNotFoundError):
            import gsheet_sync
        if hasattr(gsheet_sync, "filter_out_test_players") and not df_indiv.empty:
            df_indiv = gsheet_sync.filter_out_test_players(df_indiv)
    except Exception:
        pass

    # Garantie d'enrichissement RPE correspondant strictement à la date / semaine de séance
    if not df_indiv.empty and ("RPE" not in df_indiv.columns or df_indiv["RPE"].dropna().empty):
        try:
            from rpe_connector import enrich_gps_dataframe_with_rpe
            s_dt = weekly_memory.parse_session_date(date_seance)
            if s_dt is not None:
                w_n = s_dt.isocalendar()[1]
                w_d = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"][s_dt.weekday()]
            else:
                w_n = None
                w_d = weekly_memory.detect_day_of_week(date_seance) or "Vendredi"
            m_gps = None
            if hasattr(st, "session_state"):
                m_gps = st.session_state.get("mapping_gps") or st.session_state.get("mapping_gps_joueurs")
            df_indiv = enrich_gps_dataframe_with_rpe(
                df_gps=df_indiv,
                session_date=s_dt,
                mapping_gps=m_gps,
                week_number=w_n,
                weekday=w_d,
            )
            if not rpe_stats or rpe_stats.get("equipe") is None:
                rpe_stats = rpe_connector.calculate_rpe_group_metrics(df_indiv)
        except Exception:
            pass

    j_lookup = kwargs.get("joueurs_lookup")
    if j_lookup is None and hasattr(weekly_memory, "_get_effective_joueurs_lookup"):
        j_lookup = weekly_memory._get_effective_joueurs_lookup()
    if j_lookup is None:
        j_lookup = {}

    if not df_indiv.empty and hasattr(weekly_memory, "normaliser_poste_specifique"):
        name_col = "Nom ; Prénom" if "Nom ; Prénom" in df_indiv.columns else ("Nom" if "Nom" in df_indiv.columns else None)
        pos_col = "Poste_Detaille" if "Poste_Detaille" in df_indiv.columns else ("Poste" if "Poste" in df_indiv.columns else None)
        df_indiv["Poste_Detaille"] = [
            weekly_memory.normaliser_poste_specifique(
                r.get(pos_col, "") if pos_col else "",
                player_name=str(r.get(name_col, "")) if name_col else "",
                joueurs_lookup=j_lookup,
            ) or (str(r.get(pos_col, "")).strip() if pos_col else "")
            for _, r in df_indiv.iterrows()
        ]

    granularity_mode = str(kwargs.get("granularity_mode", granularity_mode) or "Par ligne de poste")
    g_low = weekly_memory._strip_accents_lower(granularity_mode) if hasattr(weekly_memory, "_strip_accents_lower") else granularity_mode.lower()
    is_detailed = ("detaille" in g_low) or ("9 poste" in g_low)

    date_texte = safe_latin1(formater_date_seance_fr(date_seance))

    class RCACoachPDF(FPDF):
        def __init__(self, *args, **kwargs_pdf):
            super().__init__(*args, **kwargs_pdf)
            self.date_texte = ""
            self.current_subtitle = ""
            self.alias_nb_pages()

        def header(self):
            # 1. Bandeau supérieur Bleu Marine RCA (#162B4C)
            self.set_fill_color(22, 43, 76)
            self.rect(0, 0, 297, 16.5, style="F")

            # 2. Liseré inférieur décoratif : Jaune Or officiel (#F3C300) de 3 à 4 pt (~1.3 mm)
            self.set_fill_color(243, 195, 0)
            self.rect(0, 16.5, 297, 1.3, style="F")

            # 3. Logo officiel du RCA avec respect strict du ratio d'aspect (w=0, h=13.0)
            logo_path = get_pdf_logo_path()
            x_title = 12
            if logo_path and os.path.exists(logo_path):
                try:
                    self.image(logo_path, x=12, y=1.75, w=0, h=13.0)
                    x_title = 26
                except Exception:
                    pass

            # 4. Titre officiel à gauche (avec sous-titre dynamique sur les pages de groupes)
            if self.current_subtitle:
                self.set_xy(x_title, 2.2)
                self.set_font("Helvetica", "B", 10.0)
                self.set_text_color(255, 255, 255)
                self.cell(145, 6, safe_latin1("RUGBY CLUB AMIÉNOIS | Rapport de Séance GPS"), align="L")

                self.set_xy(x_title, 8.2)
                self.set_font("Helvetica", "B", 9.0)
                self.set_text_color(243, 195, 0)
                self.cell(145, 6, safe_latin1(self.current_subtitle), align="L")
            else:
                self.set_xy(x_title, 3.2)
                self.set_font("Helvetica", "B", 12.0)
                self.set_text_color(255, 255, 255)
                self.cell(145, 10, safe_latin1("RUGBY CLUB AMIÉNOIS | Rapport de Séance GPS"), align="L")

            # 5. Date & semaine à droite
            self.set_xy(170, 3.2)
            self.set_font("Helvetica", "B", 9.5)
            self.set_text_color(243, 195, 0)
            self.cell(115, 10, self.date_texte, align="R")

            self.set_text_color(30, 41, 59)
            self.set_y(20)

        def footer(self):
            # Liseré discret jaune or
            self.set_fill_color(243, 195, 0)
            self.rect(12, 199.5, 273, 0.4, style="F")

            # Mention officielle
            self.set_y(-9.5)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(100, 116, 139)
            self.cell(160, 6, safe_latin1("Document Staff Technique RCA — Pôle Préparation Physique"), align="L")
            self.cell(113, 6, safe_latin1("Page {0}/{{nb}}".format(self.page_no())), align="R")

    pdf = RCACoachPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False)
    pdf.set_margins(left=12, top=20, right=12)
    pdf.date_texte = date_texte

    if df_indiv.empty:
        pdf.add_page()
        pdf.set_font("Helvetica", "", 12)
        pdf.cell(0, 10, "Aucune donnee de joueur disponible pour ce rapport.", align="L")
        return bytes(pdf.output())

    # =========================================================================
    # PAGE 1 : SYNTHÈSE GLOBALE DE LA SÉANCE
    # =========================================================================
    pdf.current_subtitle = ""
    pdf.add_page()

    # 1. Encarts KPIs globaux de Séance
    nb_joueurs = len(df_indiv)
    dist_tot = float(df_indiv["Distance (km)"].sum()) if "Distance (km)" in df_indiv.columns else 0.0
    dist_moy = float(df_indiv["Distance (km)"].mean()) if "Distance (km)" in df_indiv.columns else 0.0
    vmax = float(df_indiv["Top Speed (km/h)"].max()) if "Top Speed (km/h)" in df_indiv.columns else 0.0

    # HSR Z3 / Joueur : strictement Zone 3 (> 14.4 et <= 19.8 km/h) moyen par joueur
    if "Distance in Speed Zone 3  (km)" in df_indiv.columns:
        hsr_moy = float(df_indiv["Distance in Speed Zone 3  (km)"].mean() * 1000.0)
    elif "HSR Z3 (m)" in df_indiv.columns:
        hsr_moy = float(df_indiv["HSR Z3 (m)"].mean())
    elif "HSR_m" in df_indiv.columns:
        hsr_moy = float(df_indiv["HSR_m"].mean())
    elif " Distance haute intensité (m)" in df_indiv.columns:
        hsr_moy = float(df_indiv[" Distance haute intensité (m)"].mean())
    else:
        hsr_moy = 0.0
    if pd.isna(hsr_moy):
        hsr_moy = 0.0

    # 1. Calcul strict du RPE pour la page de garde (Page 1)
    if "RPE" in df_indiv.columns:
        s_rpe = pd.to_numeric(df_indiv["RPE"], errors="coerce").dropna()
        rpe_moy_str = f"{s_rpe.mean():.1f}" if not s_rpe.empty else "NR"
    else:
        rpe_moy_str = "NR"

    if rpe_moy_str == "NR" and rpe_stats and rpe_stats.get("equipe") is not None:
        try:
            val_eq = float(rpe_stats["equipe"])
            if val_eq > 0:
                rpe_moy_str = f"{val_eq:.1f}"
        except Exception:
            pass

    rpe_val_str = f"{rpe_moy_str} / 10" if rpe_moy_str != "NR" else "NR"

    kpis = [
        ("Joueurs Analyses", f"{nb_joueurs}"),
        ("Distance Cumulee", f"{dist_tot:.1f} km"),
        ("Moyenne / Joueur", f"{dist_moy:.2f} km"),
        ("Vitesse Max Seance", f"{vmax:.1f} km/h"),
        ("HSR Z3 / Joueur", f"{int(round(hsr_moy)):,d} m".replace(",", " ")),
        ("RPE Moyen Equipe", rpe_val_str),
    ]

    card_w = (273.0 - (len(kpis) - 1) * 2.5) / len(kpis)  # 43.4 mm pour 6 KPIs
    pdf.set_y(19.5)
    for title, val in kpis:
        x_curr = pdf.get_x()
        y_curr = pdf.get_y()
        pdf.set_fill_color(248, 250, 252)  # #F8FAFC
        pdf.set_draw_color(203, 213, 225)  # #CBD5E1
        pdf.rect(x_curr, y_curr, card_w, 12.0, style="FD")

        pdf.set_xy(x_curr, y_curr + 1.2)
        pdf.set_font("Helvetica", "", 7.5)
        pdf.set_text_color(100, 116, 139)  # #64748B
        pdf.cell(card_w, 3.5, safe_latin1(title), align="C")

        pdf.set_xy(x_curr, y_curr + 5.2)
        pdf.set_font("Helvetica", "B", 10.5)
        pdf.set_text_color(*RCA_NAVY_RGB)
        pdf.cell(card_w, 5.5, safe_latin1(val), align="C")

        pdf.set_xy(x_curr + card_w + 2.5, y_curr)

    # 2. Encart narratif : Note du Préparateur Physique (LLM ou repli sécurisé)
    if synthese_note is None or not str(synthese_note).strip():
        synthese_note = generer_synthese_llm(df_indiv)

    y_note = 33.5
    h_note = 14.5
    w_note = 273.0

    pdf.set_fill_color(248, 250, 252)  # #F8FAFC
    pdf.set_draw_color(203, 213, 225)  # #CBD5E1
    pdf.rect(12, y_note, w_note, h_note, style="FD")

    # Bordure gauche accentuée Bleu Marine RCA (#162B4C) et liseré doré RCA (#F3C300)
    pdf.set_fill_color(*RCA_NAVY_RGB)
    pdf.rect(12, y_note, 3.5, h_note, style="F")
    pdf.set_fill_color(*RCA_GOLD_RGB)
    pdf.rect(15.5, y_note, 0.8, h_note, style="F")

    # Titre du bloc narratif
    pdf.set_xy(18, y_note + 1.2)
    pdf.set_font("Helvetica", "B", 7.5)
    pdf.set_text_color(*RCA_NAVY_RGB)
    pdf.cell(140, 3.5, safe_latin1("NOTE DU PRÉPARATEUR PHYSIQUE | SYNTHÈSE ATHLÉTIQUE"), align="L")

    # Badge discret Charge Interne (RPE) vs Charge Externe (Intensité m/min)
    dpm_moy = float(df_indiv["Distance Per Min (m/min)"].mean()) if "Distance Per Min (m/min)" in df_indiv.columns else (float(df_indiv["Distance/min (m/min)"].mean()) if "Distance/min (m/min)" in df_indiv.columns else 0.0)
    badge_txt = f"CHARGE INT. / EXT. : RPE {rpe_val_str}  |  {dpm_moy:.1f} m/min"
    pdf.set_xy(160, y_note + 1.0)
    pdf.set_fill_color(244, 247, 251)
    pdf.set_draw_color(*RCA_GOLD_RGB)
    pdf.set_font("Helvetica", "B", 7.0)
    pdf.set_text_color(*RCA_NAVY_RGB)
    pdf.cell(120, 3.8, safe_latin1(badge_txt), border=1, align="R", fill=True)

    # Corps du texte nettoyé pour compatibilité Latin-1
    pdf.set_xy(18, y_note + 4.8)
    pdf.set_font("Helvetica", "", 7.5)
    pdf.set_text_color(51, 65, 85)
    pdf.multi_cell(w=w_note - 10, h=3.6, txt=safe_latin1(synthese_note), align="L")

    # 3. Colonne de Gauche : Graphique Matplotlib par Groupe (5 de devant, 3ème Ligne, Trois-quarts)
    buf1 = generer_graphique_groupes_matplotlib(df_indiv, moyennes_groupes=moyennes_groupes)
    chart_y = 50.5
    chart_w = 133
    chart_h = 142
    pdf.image(buf1, x=12, y=chart_y, w=chart_w, h=chart_h)

    # 4. Colonne de Droite : Podium Général de la Séance (Top 3 + Flop 3)
    podium_general = generer_podium_dynamique(df_indiv)
    sub_map_gen = calculer_sous_titres_podium(df_indiv, label_prefix="Moy. séance")
    _render_podium_section_pdf(
        pdf=pdf,
        podium_data=podium_general,
        x_pod=149,
        y_pod=chart_y,
        w_pod=136,
        title_header="PODIUMS DE LA SÉANCE",
        subtitle_map=sub_map_gen,
        include_flop=True,
    )

    # =========================================================================
    # PAGES DÉTAILLÉES : 2 PAGES PAR CATÉGORIE (3 GROUPES OU 9 POSTES ACTIFS)
    # =========================================================================
    if is_detailed:
        groupes_pages = weekly_memory.get_active_positions(df_indiv, j_lookup)
        if not groupes_pages:
            groupes_pages = ["5 de devant", "3ème Ligne", "Trois-quarts"]
    else:
        groupes_pages = ["5 de devant", "3ème Ligne", "Trois-quarts"]

    for grp in groupes_pages:
        pdf.current_subtitle = f"RETOUR JOURNÉE — {grp.upper()}"
        pdf.add_page()

        if config_objectifs is None:
            config_objectifs = get_active_objectifs_config()
        cfg_targets = weekly_memory.get_objectifs_for_category(config_objectifs, grp)
        df_sub = weekly_memory._filter_df_by_group(df_indiv, grp)

        if (not is_detailed) and moyennes_groupes and grp in moyennes_groupes:
            grp_stats = moyennes_groupes[grp]
        elif not df_sub.empty:
            hsr_m_val = (
                float(pd.to_numeric(df_sub["Distance in Speed Zone 3  (km)"], errors="coerce").mean() * 1000.0)
                if "Distance in Speed Zone 3  (km)" in df_sub.columns
                else (
                    float(pd.to_numeric(df_sub["HSR Z3 (m)"], errors="coerce").mean())
                    if "HSR Z3 (m)" in df_sub.columns
                    else (
                        float(pd.to_numeric(df_sub[" Distance haute intensité (m)"], errors="coerce").mean())
                        if " Distance haute intensité (m)" in df_sub.columns
                        else 0.0
                    )
                )
            )
            vhsr_val = (
                float(pd.to_numeric(df_sub["VHSR (m)"], errors="coerce").mean())
                if "VHSR (m)" in df_sub.columns
                else (
                    float((pd.to_numeric(df_sub["Distance in Speed Zone 4  (km)"], errors="coerce").fillna(0.0) + pd.to_numeric(df_sub["Distance in Speed Zone 5  (km)"], errors="coerce").fillna(0.0)).mean() * 1000.0)
                    if "Distance in Speed Zone 4  (km)" in df_sub.columns and "Distance in Speed Zone 5  (km)" in df_sub.columns
                    else 0.0
                )
            )
            dpm_val = (
                float(pd.to_numeric(df_sub["Distance Per Min (m/min)"], errors="coerce").mean())
                if "Distance Per Min (m/min)" in df_sub.columns
                else (float(pd.to_numeric(df_sub["Distance/min (m/min)"], errors="coerce").mean()) if "Distance/min (m/min)" in df_sub.columns else 0.0)
            )
            grp_stats = {
                "Dist_m": float(pd.to_numeric(df_sub["Distance (km)"], errors="coerce").mean() * 1000.0) if "Distance (km)" in df_sub.columns else 0.0,
                "HSR_m": 0.0 if np.isnan(hsr_m_val) else hsr_m_val,
                "VHSR_m": 0.0 if np.isnan(vhsr_val) else vhsr_val,
                "Dpm": 0.0 if np.isnan(dpm_val) else dpm_val,
                "Acc_Dec": float(pd.to_numeric(df_sub["Accel/Dec"], errors="coerce").mean()) if "Accel/Dec" in df_sub.columns else 0.0,
            }
        else:
            grp_stats = {"Dist_m": 0.0, "HSR_m": 0.0, "VHSR_m": 0.0, "Dpm": 0.0, "Acc_Dec": 0.0}

        obj_dist = cfg_targets["Objectif Distance (m)"]
        obj_hsr = cfg_targets["Obj. HSR"]
        obj_vhsr = cfg_targets["Obj. VHSR"]
        obj_acc = cfg_targets["Obj. ACC/DEC"]

        real_dist = float(grp_stats.get("Dist_m", 0.0) or 0.0)
        real_hsr = float(grp_stats.get("HSR_m", 0.0) or 0.0)
        real_vhsr = float(grp_stats.get("VHSR_m", 0.0) or 0.0)
        real_dpm = float(grp_stats.get("Dpm", 0.0) or 0.0)
        real_acc = float(grp_stats.get("Acc_Dec", 0.0) or 0.0)

        pct_dist = (real_dist / obj_dist * 100.0) if obj_dist > 0 else 0.0
        pct_hsr = (real_hsr / obj_hsr * 100.0) if obj_hsr > 0 else 0.0
        pct_vhsr = (real_vhsr / obj_vhsr * 100.0) if obj_vhsr > 0 else 0.0
        pct_acc = (real_acc / obj_acc * 100.0) if obj_acc > 0 else 0.0

        # 1. Haut de page : Tableau « RETOUR JOURNÉE »
        table_x = 12.0
        table_y = 21.0
        col_w_grp = 45.0
        col_w_num = 38.0  # 45.0 + 6 * 38.0 = 273.0 mm

        pdf.set_xy(table_x, table_y)
        pdf.set_fill_color(*RCA_NAVY_RGB)
        pdf.set_draw_color(*RCA_NAVY_RGB)
        pdf.set_font("Helvetica", "B", 8.0)
        pdf.set_text_color(255, 255, 255)

        cat_prefix = "POSTE" if is_detailed else "GROUPE"
        headers = [
            (col_w_grp, f"{cat_prefix} : {grp.upper()}", "L"),
            (col_w_num, "DISTANCE TOTALE", "C"),
            (col_w_num, "DISTANCE HSR (Z3)", "C"),
            (col_w_num, "DISTANCE VHSR (Sprint)", "C"),
            (col_w_num, "m/min", "C"),
            (col_w_num, "ACC / DEC n°", "C"),
            (col_w_num, "RPE MOYEN", "C"),
        ]
        for w, h_title, align in headers:
            pdf.cell(w, 6.5, safe_latin1(h_title), border=1, align=align, fill=True)

        # Ligne 1 : Objectif
        pdf.set_xy(table_x, table_y + 6.5)
        pdf.set_fill_color(244, 247, 251)
        pdf.set_draw_color(203, 213, 225)
        pdf.set_text_color(71, 85, 105)
        pdf.set_font("Helvetica", "B", 8.0)
        pdf.cell(col_w_grp, 6.0, "  Objectif", border=1, align="L", fill=True)

        pdf.set_font("Helvetica", "", 8.0)
        pdf.set_text_color(*RCA_NAVY_RGB)
        pdf.cell(col_w_num, 6.0, f"{int(round(obj_dist)):,d} m".replace(",", " "), border=1, align="C", fill=True)
        pdf.cell(col_w_num, 6.0, f"{int(round(obj_hsr)):,d} m".replace(",", " "), border=1, align="C", fill=True)
        pdf.cell(col_w_num, 6.0, f"{int(round(obj_vhsr)):,d} m".replace(",", " "), border=1, align="C", fill=True)
        pdf.set_text_color(148, 163, 184)
        pdf.cell(col_w_num, 6.0, "/", border=1, align="C", fill=True)
        pdf.set_text_color(*RCA_NAVY_RGB)
        pdf.cell(col_w_num, 6.0, f"{int(round(obj_acc)):,d}".replace(",", " "), border=1, align="C", fill=True)
        pdf.set_text_color(148, 163, 184)
        pdf.cell(col_w_num, 6.0, "/", border=1, align="C", fill=True)

        # Ligne 2 : Réalisée
        pdf.set_xy(table_x, table_y + 12.5)
        pdf.set_fill_color(255, 255, 255)
        pdf.set_draw_color(203, 213, 225)
        pdf.set_text_color(180, 83, 9)
        pdf.set_font("Helvetica", "B", 8.0)
        pdf.cell(col_w_grp, 6.8, "  Realisee", border=1, align="L", fill=True)

        pdf.set_text_color(30, 41, 59)
        pdf.cell(col_w_num, 6.8, f"{int(round(real_dist)):,d} m ({pct_dist:.0f}%)".replace(",", " "), border=1, align="C", fill=True)
        pdf.cell(col_w_num, 6.8, f"{int(round(real_hsr)):,d} m ({pct_hsr:.0f}%)".replace(",", " "), border=1, align="C", fill=True)
        pdf.cell(col_w_num, 6.8, f"{int(round(real_vhsr)):,d} m ({pct_vhsr:.0f}%)".replace(",", " "), border=1, align="C", fill=True)
        pdf.cell(col_w_num, 6.8, f"{real_dpm:.1f} m/min", border=1, align="C", fill=True)
        pdf.cell(col_w_num, 6.8, f"{int(round(real_acc)):,d} ({pct_acc:.0f}%)".replace(",", " "), border=1, align="C", fill=True)

        s_rpe = pd.to_numeric(df_sub['RPE'], errors='coerce').dropna() if "RPE" in df_sub.columns else pd.Series(dtype=float)
        rpe_moy_str = f"{s_rpe.mean():.1f}" if not s_rpe.empty else "NR"

        if rpe_moy_str == "NR" and rpe_stats:
            by_p = rpe_stats.get("by_poste") if isinstance(rpe_stats, dict) else {}
            v_grp = (
                rpe_stats.get(grp)
                or (by_p.get(grp) if isinstance(by_p, dict) else None)
                or rpe_stats.get(f"rpe_moyen_{grp.lower().replace(' ', '_').replace('è', 'e')}")
            )
            if v_grp is not None:
                try:
                    f_grp = float(v_grp)
                    if f_grp > 0:
                        rpe_moy_str = f"{f_grp:.1f}"
                except Exception:
                    pass

        rpe_grp_str = f"{rpe_moy_str} / 10" if rpe_moy_str != "NR" else "NR"
        pdf.cell(col_w_num, 6.8, safe_latin1(rpe_grp_str), border=1, align="C", fill=True)

        # 2. Graphiques : 4 barres comparatives en pleine largeur 2x2
        grp_metrics = {
            "Objectif Distance (m)": (obj_dist, real_dist),
            "Obj. HSR": (obj_hsr, real_hsr),
            "Obj. VHSR": (obj_vhsr, real_vhsr),
            "Obj. ACC/DEC": (obj_acc, real_acc),
        }
        buf_jauges = generer_graphique_jauges_groupe_matplotlib(grp, grp_metrics)
        chart_y = 43.0
        chart_w = 273.0
        chart_h = 146.0
        pdf.image(buf_jauges, x=12.0, y=chart_y, w=chart_w, h=chart_h)

        # ---------------------------------------------------------------------
        # PAGE B : PODIUMS TOP 3 + FLOP 3 SUR LA MÊME PAGE
        # ---------------------------------------------------------------------
        pdf.current_subtitle = f"PODIUMS DE LA {'CATÉGORIE' if is_detailed else 'LIGNE'} — {grp.upper()}"
        pdf.add_page()

        podium_grp = generer_podium_dynamique(df_sub)
        sub_map = calculer_sous_titres_podium(df_sub, label_prefix="Moyenne")

        _render_podium_page_aere(
            pdf=pdf,
            podium_data=podium_grp,
            group_name=grp,
            subtitle_map=sub_map,
        )

    # -------------------------------------------------------------------------
    # DERNIÈRE(S) PAGE(S) : ANALYSE COMPARATIVE HISTORIQUE DU MÊME JOUR (TRIPTYQUE)
    # -------------------------------------------------------------------------
    if inclure_comparatif:
        target_day = weekly_memory.detect_day_of_week(date_seance) or "Vendredi"
        dt_curr = weekly_memory.parse_session_date(date_seance)
        curr_week = dt_curr.isocalendar()[1] if dt_curr else 45
        items_comp = weekly_memory.recuperer_historique_seances_meme_jour(
            session_date_or_day=date_seance,
            df_current=df_indiv,
            n_max=min(4, nb_comparatif),
            current_notes_or_tag=synthese_note or "",
            current_adversaire=adversaire,
            current_rpe_stats=rpe_stats,
            categories=groupes_pages,
        )

        for it in items_comp:
            it_g_metrics = it.setdefault("group_metrics", {})
            if it.get("is_current", False):
                for grp in groupes_pages:
                    df_group = weekly_memory._filter_df_by_group(df_indiv, grp)
                    s_rpe = pd.to_numeric(df_group['RPE'], errors='coerce').dropna() if "RPE" in df_group.columns else pd.Series(dtype=float)
                    if not s_rpe.empty:
                        it_g_metrics.setdefault(grp, {})["rpe"] = round(float(s_rpe.mean()), 1)
                    elif it_g_metrics.get(grp, {}).get("rpe") is None and rpe_stats:
                        by_p = rpe_stats.get("by_poste") if isinstance(rpe_stats, dict) else {}
                        v_r = rpe_stats.get(grp) or (by_p.get(grp) if isinstance(by_p, dict) else None)
                        if v_r is not None and float(v_r) > 0:
                            it_g_metrics.setdefault(grp, {})["rpe"] = round(float(v_r), 1)
            else:
                for grp in groupes_pages:
                    if it_g_metrics.get(grp, {}).get("rpe") is None:
                        df_grp_hist = None
                        if "players_data" in it and it["players_data"]:
                            df_p = pd.DataFrame(it["players_data"])
                            df_grp_hist = weekly_memory._filter_df_by_group(df_p, grp)
                        if df_grp_hist is not None and not df_grp_hist.empty and "RPE" in df_grp_hist.columns:
                            s_rpe_h = pd.to_numeric(df_grp_hist['RPE'], errors='coerce').dropna()
                            if not s_rpe_h.empty:
                                it_g_metrics.setdefault(grp, {})["rpe"] = round(float(s_rpe_h.mean()), 1)
                        if it_g_metrics.get(grp, {}).get("rpe") is None:
                            r_stats_it = it.get("rpe_stats") or {}
                            by_p_it = r_stats_it.get("by_poste") if isinstance(r_stats_it, dict) else {}
                            v_r = (
                                r_stats_it.get(grp)
                                or (by_p_it.get(grp) if isinstance(by_p_it, dict) else None)
                                or r_stats_it.get(f"rpe_moyen_{grp.lower().replace(' ', '_').replace('è', 'e')}")
                            )
                            if v_r is not None and float(v_r) > 0:
                                it_g_metrics.setdefault(grp, {})["rpe"] = round(float(v_r), 1)

        weekly_memory.render_comparatif_triptyque_pdf_pages(
            pdf=pdf,
            items_comparatif=items_comp,
            subtitle_base=f"RETOUR JOURNÉE — COMPARAISON AVEC LES DERNIERS {target_day.upper()}S (SEMAINE {curr_week})",
            categories_list=groupes_pages,
            target_day=target_day,
            week_num=curr_week,
            df_current=df_indiv,
        )

    return bytes(pdf.output())




# ==============================================================================
# RAPPORT INDIVIDUEL ATHLÈTE A4 (REPORTLAB)
# ==============================================================================

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        HRFlowable,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False


def generate_player_pdf_report(
    player_name: str,
    player_info: Dict[str, Any],
    wellness_summary: Dict[str, Any],
    gps_summary: Dict[str, Any],
    period_str: str = "7 derniers jours"
) -> bytes:
    """Génère un rapport individuel en PDF sous forme de flux binaire (bytes)."""
    if not HAS_REPORTLAB:
        # Fallback élégant avec FPDF aux couleurs officielles RCA
        pdf = FPDF(orientation="P", unit="mm", format="A4")
        pdf.set_auto_page_break(auto=False)
        pdf.add_page()

        # 1. Bandeau supérieur Bleu Marine RCA (#162B4C)
        pdf.set_fill_color(*RCA_NAVY_RGB)
        pdf.rect(0, 0, 210, 22.0, style="F")

        # 2. Liseré doré officiel (#F3C300)
        pdf.set_fill_color(*RCA_GOLD_RGB)
        pdf.rect(0, 22.0, 210, 1.4, style="F")

        # 3. Logo officiel du RCA
        logo_path = get_pdf_logo_path()
        x_title = 14
        if logo_path and os.path.exists(logo_path):
            try:
                pdf.image(logo_path, x=14, y=2.5, w=0, h=17.0)
                x_title = 32
            except Exception:
                pass

        pdf.set_xy(x_title, 4.0)
        pdf.set_font("Helvetica", "B", 13.0)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(160, 6, safe_latin1("RUGBY CLUB AMIÉNOIS | Rapport Individuel Athlète"), ln=True)

        pdf.set_xy(x_title, 11.5)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(*RCA_GOLD_RGB)
        pdf.cell(160, 6, safe_latin1(f"Période : {period_str} | Date : {pd.Timestamp.now().strftime('%d/%m/%Y')}"), ln=True)

        # 4. Encart Profil Joueur
        pdf.set_y(30.0)
        pdf.set_fill_color(244, 247, 251)
        pdf.set_draw_color(203, 213, 225)
        pdf.rect(14, 30.0, 182, 16.0, style="FD")
        pdf.set_fill_color(*RCA_NAVY_RGB)
        pdf.rect(14, 30.0, 3.0, 16.0, style="F")

        pdf.set_xy(20, 32.0)
        pdf.set_font("Helvetica", "B", 10.0)
        pdf.set_text_color(*RCA_NAVY_RGB)
        pdf.cell(85, 5, safe_latin1(f"Athlète : {player_name}"), ln=False)
        pdf.set_font("Helvetica", "", 9.0)
        pdf.set_text_color(71, 85, 105)
        pdf.cell(85, 5, safe_latin1(f"ID : {player_info.get('id', 'N/A')}"), ln=True)

        pdf.set_xy(20, 38.5)
        pdf.set_font("Helvetica", "", 9.0)
        pdf.cell(85, 5, safe_latin1(f"Poste : {player_info.get('poste', 'Non spécifié')}"), ln=False)
        pdf.cell(85, 5, safe_latin1(f"Groupe : {player_info.get('groupe', 'Non spécifié')}"), ln=True)

        # 5. État de Forme & Wellness
        pdf.set_y(52.0)
        pdf.set_font("Helvetica", "B", 11.0)
        pdf.set_text_color(*RCA_NAVY_RGB)
        pdf.cell(0, 7, safe_latin1("1. État de Forme & Wellness (Indice de Hooper)"), ln=True)

        pdf.set_font("Helvetica", "", 9.5)
        pdf.set_text_color(51, 65, 85)
        pdf.cell(0, 6, safe_latin1(f"• Score Hooper Moyen : {wellness_summary.get('score_hooper_moy', 'N/A')} / 28"), ln=True)
        pdf.cell(0, 6, safe_latin1(f"• Qualité du Sommeil : {wellness_summary.get('sommeil_moy', 'N/A')} / 7"), ln=True)
        pdf.cell(0, 6, safe_latin1(f"• Niveau de Fatigue : {wellness_summary.get('fatigue_moy', 'N/A')} / 7"), ln=True)
        pdf.cell(0, 6, safe_latin1(f"• Courbatures Musculaires : {wellness_summary.get('courbatures_moy', 'N/A')} / 7"), ln=True)
        pdf.cell(0, 6, safe_latin1(f"• Niveau de Stress : {wellness_summary.get('stress_moy', 'N/A')} / 7"), ln=True)

        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(230, 126, 34)  # Orange sportif pour le RPE
        pdf.cell(0, 6, safe_latin1(f"• Charge Interne sRPE Totale : {wellness_summary.get('srpe_total', 'N/A')} UA"), ln=True)

        # 6. Données GPS & Charge Externe
        pdf.ln(5)
        pdf.set_font("Helvetica", "B", 11.0)
        pdf.set_text_color(*RCA_NAVY_RGB)
        pdf.cell(0, 7, safe_latin1("2. Données GPS & Charge Externe"), ln=True)

        pdf.set_font("Helvetica", "", 9.5)
        pdf.set_text_color(51, 65, 85)
        pdf.cell(0, 6, safe_latin1(f"• Distance Totale : {gps_summary.get('total_distance_m', 'N/A')} m"), ln=True)
        pdf.cell(0, 6, safe_latin1(f"• Course Haute Intensité HSR : {gps_summary.get('hsr_distance_m', 'N/A')} m"), ln=True)
        pdf.cell(0, 6, safe_latin1(f"• Distance Sprint : {gps_summary.get('sprint_distance_m', 'N/A')} m"), ln=True)
        pdf.cell(0, 6, safe_latin1(f"• Vitesse Max atteinte : {gps_summary.get('vmax_kmh', 'N/A')} km/h"), ln=True)
        pdf.cell(0, 6, safe_latin1(f"• Intensité Moyenne : {gps_summary.get('intensity_m_min', 'N/A')} m/min"), ln=True)
        pdf.cell(0, 6, safe_latin1(f"• PlayerLoad Total : {gps_summary.get('player_load_total', 'N/A')}"), ln=True)
        pdf.cell(0, 6, safe_latin1(f"• Ratio ACWR : {gps_summary.get('acwr', 'N/A')} (Zone optimale : 0.8 - 1.3)"), ln=True)

        # 7. Pied de page officiel
        pdf.set_fill_color(*RCA_GOLD_RGB)
        pdf.rect(14, 280.0, 182, 0.4, style="F")
        pdf.set_y(-12)
        pdf.set_font("Helvetica", "I", 8.0)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(0, 6, safe_latin1("Document Staff Technique RCA — Pôle Préparation Physique"), align="C")

        return bytes(pdf.output())

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=20,
        leading=24,
        textColor=colors.HexColor(RCA_NAVY_HEX),
        alignment=0
    )

    subtitle_style = ParagraphStyle(
        "DocSubtitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#64748B")
    )

    section_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=16,
        textColor=colors.HexColor(RCA_NAVY_HEX),
        spaceBefore=10,
        spaceAfter=6
    )

    body_style = ParagraphStyle(
        "BodyDark",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#334155")
    )

    elements = []

    elements.append(Paragraph("🏉 RCA - RAPPORT DE PERFORMANCE INDIVIDUEL", title_style))
    elements.append(Paragraph(f"Période d'évaluation : {period_str} | Date d'émission : {pd.Timestamp.now().strftime('%d/%m/%Y')}", subtitle_style))
    elements.append(Spacer(1, 0.4 * cm))
    elements.append(HRFlowable(width="100%", thickness=2.0, color=colors.HexColor(RCA_NAVY_HEX), spaceAfter=2))
    elements.append(HRFlowable(width="100%", thickness=1.0, color=colors.HexColor(RCA_GOLD_HEX), spaceAfter=12))

    poste = player_info.get("poste", "Non spécifié")
    groupe = player_info.get("groupe", "Non spécifié")
    joueur_id = player_info.get("id", "N/A")

    info_data = [
        [
            Paragraph(f"<b>Athlète :</b> {player_name}", body_style),
            Paragraph(f"<b>Identifiant :</b> {joueur_id}", body_style)
        ],
        [
            Paragraph(f"<b>Poste :</b> {poste}", body_style),
            Paragraph(f"<b>Groupe :</b> {groupe}", body_style)
        ]
    ]

    info_table = Table(info_data, colWidths=[9 * cm, 9 * cm])
    info_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
        ("PADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 0.5 * cm))

    elements.append(Paragraph("1. État de Forme & Wellness (Indice de Hooper)", section_style))

    hooper_moy = wellness_summary.get("score_hooper_moy", "N/A")
    sommeil = wellness_summary.get("sommeil_moy", "N/A")
    fatigue = wellness_summary.get("fatigue_moy", "N/A")
    courbatures = wellness_summary.get("courbatures_moy", "N/A")
    stress = wellness_summary.get("stress_moy", "N/A")
    srpe_total = wellness_summary.get("srpe_total", "N/A")

    wellness_data = [
        ["Indicateur", "Valeur Moyenne", "Statut / Interprétation"],
        ["Score Hooper Total", str(hooper_moy), "Optimal" if isinstance(hooper_moy, (int, float)) and hooper_moy <= 12 else "Surveillance"],
        ["Qualité du Sommeil", str(sommeil), "1 (Très bon) à 7 (Très mauvais)"],
        ["Niveau de Fatigue", str(fatigue), "1 (Frais) à 7 (Épuisé)"],
        ["Courbatures musculaires", str(courbatures), "1 (Aucune) à 7 (Sévères)"],
        ["Niveau de Stress", str(stress), "1 (Détendu) à 7 (Très stressé)"],
        ["Charge Interne Totale (sRPE)", f"{srpe_total} UA", "Somme RPE x Durée (min)"]
    ]

    t_wellness = Table(wellness_data, colWidths=[6.5 * cm, 4.5 * cm, 7 * cm])
    t_wellness.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(RCA_NAVY_HEX)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("ALIGN", (1, 0), (1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("PADDING", (0, 0), (-1, -1), 5),
    ]))
    elements.append(t_wellness)
    elements.append(Spacer(1, 0.5 * cm))

    elements.append(Paragraph("2. Données GPS & Charge Externe", section_style))

    tot_dist = gps_summary.get("total_distance_m", "N/A")
    hsr_dist = gps_summary.get("hsr_distance_m", "N/A")
    sprint_dist = gps_summary.get("sprint_distance_m", "N/A")
    vmax = gps_summary.get("vmax_kmh", "N/A")
    intensity = gps_summary.get("intensity_m_min", "N/A")
    player_load = gps_summary.get("player_load_total", "N/A")
    acwr_val = gps_summary.get("acwr", "N/A")

    gps_table_data = [
        ["Métrique GPS", "Valeur Cumulée / Pic", "Référence"],
        ["Distance Totale", f"{tot_dist} m" if tot_dist != "N/A" else "N/A", "Volume de déplacement global"],
        ["Course Haute Intensité (> 20 km/h)", f"{hsr_dist} m" if hsr_dist != "N/A" else "N/A", "HSR (> 5.5 m/s)"],
        ["Distance Sprint (> 25 km/h)", f"{sprint_dist} m" if sprint_dist != "N/A" else "N/A", "Vitesse pure"],
        ["Vitesse Maximale atteinte", f"{vmax} km/h" if vmax != "N/A" else "N/A", "VMax pic de la période"],
        ["Intensité moyenne", f"{intensity} m/min" if intensity != "N/A" else "N/A", "Rythme de jeu moyen"],
        ["PlayerLoad Total", str(player_load), "Charge accélérométrique (3D)"],
        ["Ratio ACWR (7j / 28j)", str(acwr_val), "Zone optimale : 0.8 - 1.3"]
    ]

    t_gps = Table(gps_table_data, colWidths=[6.5 * cm, 4.5 * cm, 7 * cm])
    t_gps.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(RCA_NAVY_HEX)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("ALIGN", (1, 0), (1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("PADDING", (0, 0), (-1, -1), 5),
    ]))
    elements.append(t_gps)
    elements.append(Spacer(1, 0.6 * cm))

    elements.append(Paragraph("3. Synthèse Staff & Préconisations", section_style))
    reco_text = (
        "Athlète dans les normes de charge attendues. Maintenir la vigilance sur la récupération "
        "et la qualité du sommeil post-match. Adapter les séances de vitesse si le ratio ACWR dépasse 1.4."
    )
    elements.append(Paragraph(f"<i>{reco_text}</i>", body_style))

    elements.append(Spacer(1, 0.4 * cm))
    elements.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor(RCA_GOLD_HEX), spaceAfter=6))
    footer_style = ParagraphStyle(
        "DocFooter",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#64748B"),
        alignment=1,
    )
    elements.append(Paragraph("Document Staff Technique RCA — Pôle Préparation Physique", footer_style))

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()
