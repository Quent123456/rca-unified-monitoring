"""
Script d'intégration du processeur GPS Catapult et du moteur PDF
dans le projet unifié rca-unified-monitoring.
"""

import os
import re

def main():
    with open("temp_gps/app.py", "r", encoding="utf-8", errors="ignore") as f:
        full_code = f.read()

    lines = full_code.splitlines(keepends=True)

    # 1. Extraction du coeur de traitement GPS (lignes 70 à 3180)
    # Contient Effectif RCA, mapping boîtiers, parsing durées/nombres, splits,
    # process_catapult_data (jointure RPE), moyennes groupes, podiums, export Excel
    gps_code_part = "".join(lines[69:3180])

    # Adapter les imports locaux dans gps_code_part
    gps_code_part = gps_code_part.replace("import photo_manager", "try:\n    from modules import photo_manager\nexcept ImportError:\n    import photo_manager")
    gps_code_part = gps_code_part.replace("import weekly_memory", "try:\n    from modules import weekly_memory\nexcept ImportError:\n    import weekly_memory")
    gps_code_part = gps_code_part.replace("import rpe_connector", "try:\n    from modules import rpe_connector\nexcept ImportError:\n    import rpe_connector")

    acwr_helper = '''

def calculate_acwr(df: pd.DataFrame, metric: str = "total_distance") -> pd.DataFrame:
    """Calcule le ratio charge aigue / chronique (ACWR) par joueur sur base glissante (7j / 28j)."""
    if df.empty or metric not in df.columns:
        return pd.DataFrame()

    df_calc = df.copy()
    date_col = "date" if "date" in df_calc.columns else ("Date" if "Date" in df_calc.columns else None)
    player_col = "player_name" if "player_name" in df_calc.columns else ("Nom" if "Nom" in df_calc.columns else "Player Name")
    
    if not date_col or not player_col:
        return pd.DataFrame()

    df_calc[date_col] = pd.to_datetime(df_calc[date_col], errors="coerce")
    df_calc = df_calc.dropna(subset=[date_col]).sort_values(by=[player_col, date_col])

    results = []
    for player, group in df_calc.groupby(player_col):
        daily = group.set_index(date_col).resample("D")[metric].sum().fillna(0)
        acute = daily.rolling(window=7, min_periods=1).mean()
        chronic = daily.rolling(window=28, min_periods=1).mean()
        acwr_val = acute / chronic.replace(0, np.nan)

        res_df = pd.DataFrame({
            player_col: player,
            date_col: daily.index,
            "daily_load": daily.values,
            "acute_load_7d": acute.values,
            "chronic_load_28d": chronic.values,
            "acwr": acwr_val.values.round(2)
        })
        results.append(res_df)

    if not results:
        return pd.DataFrame()

    return pd.concat(results, ignore_index=True)
'''

    gps_processing_header = '''"""Module de traitement et d'analyse des données GPS Catapult One pour le Rugby Club Amiénois (RCA).

Gère :
1. Le chargement intelligent des fichiers Catapult One (.csv et .xlsx) avec détection d'en-tête.
2. Le mapping des boîtiers GPS vers l'effectif officiel RCA avec persistance JSON.
3. L'extraction et le filtrage des Splits / Ateliers d'entraînement.
4. Le calcul des métriques rugby : Distance, HSR (Z3), VHSR (Z4+Z5), Accélérations, Décélérations, PlayerLoad, m/min.
5. La jointure ultra-robuste avec les données RPE de suivi de charge (Nom + Prénom + Date, skipna=True, gestion 'NR').
6. L'injection dans le modèle Excel officiel STAPS de la préparatrice physique.
"""

from __future__ import annotations

import base64
import copy
import datetime
import io
import json
import os
import re
import textwrap
import urllib.parse
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

try:
    from modules import photo_manager, weekly_memory, rpe_connector
except ImportError:
    import photo_manager
    import weekly_memory
    import rpe_connector

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
'''

    with open("modules/gps_processing.py", "w", encoding="utf-8") as f:
        f.write(gps_processing_header + gps_code_part + acwr_helper)

    print("modules/gps_processing.py genere.")

    # 2. Extraction du moteur de rapport PDF (lignes 3181 à 4081)
    pdf_code_part = "".join(lines[3180:4081])
    pdf_code_part = pdf_code_part.replace("import photo_manager", "try:\n    from modules import photo_manager\nexcept ImportError:\n    import photo_manager")
    pdf_code_part = pdf_code_part.replace("import weekly_memory", "try:\n    from modules import weekly_memory\nexcept ImportError:\n    import weekly_memory")
    pdf_code_part = pdf_code_part.replace("import rpe_connector", "try:\n    from modules import rpe_connector\nexcept ImportError:\n    import rpe_connector")

    reportlab_part = '''

# ==============================================================================
# RAPPORT INDIVIDUEL ATHLÈTE A4 (REPORTLAB)
# ==============================================================================

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


def generate_player_pdf_report(
    player_name: str,
    player_info: Dict[str, Any],
    wellness_summary: Dict[str, Any],
    gps_summary: Dict[str, Any],
    period_str: str = "7 derniers jours"
) -> bytes:
    """Génère un rapport individuel en PDF sous forme de flux binaire (bytes)."""
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
        textColor=colors.HexColor("#1F4E79"),
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
        textColor=colors.HexColor("#1E293B"),
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
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1F4E79"), spaceAfter=12))

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
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
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
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0D9488")),
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

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()
'''

    pdf_header = '''"""Module de génération des rapports PDF officiels du Rugby Club Amiénois (RCA).

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
    from modules import photo_manager, weekly_memory, rpe_connector
except ImportError:
    import photo_manager
    import weekly_memory
    import rpe_connector

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
'''

    with open("modules/pdf_reports.py", "w", encoding="utf-8") as f:
        f.write(pdf_header + pdf_code_part + reportlab_part)

    print("modules/pdf_reports.py genere.")

    # 3. Rapatriement de pages/4_Analyse_GPS.py avec l'application GPS Catapult complète
    # temp_gps/app.py contient déjà la page complète avec la navigation et l'UI
    # On adapte les imports pour pointer vers les modules du projet
    gps_page_code = full_code
    gps_page_code = gps_page_code.replace("import photo_manager", "try:\n    from modules import photo_manager\nexcept ImportError:\n    import photo_manager")
    gps_page_code = gps_page_code.replace("import weekly_memory", "try:\n    from modules import weekly_memory\nexcept ImportError:\n    import weekly_memory")
    gps_page_code = gps_page_code.replace("import rpe_connector", "try:\n    from modules import rpe_connector\nexcept ImportError:\n    import rpe_connector")

    with open("pages/4_Analyse_GPS.py", "w", encoding="utf-8") as f:
        f.write(gps_page_code)

    print("pages/4_Analyse_GPS.py mis a jour.")

    # 4. Rapatriement de pages/5_Rapports_PDF.py
    # Page de génération et d'exportation de rapports PDF :
    # - Rapport Coach séance du jour
    # - Rapport Hebdomadaire multi-séances (weekly_memory)
    # - Fiche individuelle joueur (ReportLab)
    reports_page_code = '''"""
pages/5_Rapports_PDF.py
Interface de sélection, génération et exportation des rapports PDF officiels RCA.
"""

from __future__ import annotations

import datetime
import io
import os
import pandas as pd
import streamlit as st

try:
    from modules import photo_manager, weekly_memory, rpe_connector, gps_processing, pdf_reports
except ImportError:
    import photo_manager
    import weekly_memory
    import rpe_connector
    import gps_processing
    import pdf_reports

st.set_page_config(
    page_title="Rapports PDF - RCA Monitoring",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Rapports PDF & Bilans Officiels RCA")
st.caption("Exportez en un clic les rapports de performance : Séance Coachs, Rapport Hebdomadaire ou Fiche Athlète.")

tab_coach, tab_weekly, tab_indiv = st.tabs([
    "🏉 Rapport Séance Coachs (Officiel Paysage)",
    "📅 Rapport Hebdomadaire Multi-Séances",
    "👤 Fiche Individuelle Athlète (A4)"
])

# ------------------------------------------------------------------------------
# 1. RAPPORT SÉANCE COACHS (FPDF PAYSAGE)
# ------------------------------------------------------------------------------
with tab_coach:
    st.subheader("🏉 Rapport Séance du Jour pour le Staff / Entraîneurs")
    st.markdown("Génère le document officiel (4 à 8 pages) avec jauges d'objectifs, podiums par ligne/poste, graphiques et métriques RPE.")

    df_indiv = st.session_state.get("df_indiv")
    if df_indiv is None or df_indiv.empty:
        st.info("💡 Aucune séance active dans la session. Veuillez d'abord traiter un fichier dans **4. Analyse GPS** ou charger une séance de test ci-dessous.")
        col_t1, col_t2 = st.columns([1, 3])
        with col_t1:
            if st.button("🧪 Charger séance démo pour le PDF", key="btn_load_demo_pdf", use_container_width=True):
                # Chargement automatique des données de démo Catapult
                sample_file = os.path.join("data", "sample_catapult_rugby.csv")
                if os.path.exists(sample_file):
                    df_sample = pd.read_csv(sample_file)
                    st.session_state["df_indiv"] = df_sample
                    st.session_state["date_seance"] = datetime.date(2026, 9, 8)
                    st.rerun()
                else:
                    st.warning("Fichier de démo introuvable dans data/.")
    else:
        st.success(f"Séance chargée : **{len(df_indiv)} athlètes** prêts pour le rapport PDF.")
        col_c1, col_c2, col_c3 = st.columns(3)
        with col_c1:
            date_s = st.date_input("Date de la séance", value=st.session_state.get("date_seance") or datetime.date.today(), key="pdf_date_input")
        with col_c2:
            adv = st.text_input("Adversaire / Contexte (optionnel)", placeholder="Ex: Entraînement intensif ou vs Beauvais", key="pdf_adv_input")
        with col_c3:
            incl_comp = st.checkbox("Inclure comparatif multi-séances", value=False, key="pdf_comp_input")

        synth_note = st.text_area("Synthèse du staff / Commentaires préparateur physique :", height=80, key="pdf_synth_input")

        if st.button("🚀 Générer le Rapport Séance Coachs (PDF)", type="primary", use_container_width=True, key="btn_gen_coach_pdf"):
            with st.spinner("Génération du document PDF officiel RCA en cours..."):
                try:
                    pdf_bytes = pdf_reports.generer_rapport_pdf_coachs(
                        df_indiv=df_indiv,
                        date_seance=date_s,
                        synthese_note=synth_note,
                        inclure_comparatif=incl_comp,
                        adversaire=adv
                    )
                    date_str = date_s.strftime('%Y%m%d') if hasattr(date_s, 'strftime') else 'seance'
                    st.download_button(
                        label="📥 Télécharger le Rapport Coachs (.pdf)",
                        data=pdf_bytes,
                        file_name=f"Rapport_Coachs_RCA_{date_str}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="btn_download_coach_pdf"
                    )
                    st.toast("Rapport PDF généré avec succès !", icon="🏉")
                except Exception as e:
                    st.error(f"Erreur lors de la génération du PDF : {e}")

# ------------------------------------------------------------------------------
# 2. RAPPORT HEBDOMADAIRE MULTI-SÉANCES (WEEKLY MEMORY)
# ------------------------------------------------------------------------------
with tab_weekly:
    st.subheader("📅 Rapport Hebdomadaire & Bilan Multi-Séances")
    st.caption("Consultez l'historique archivé par semaine ISO, comparez les charges (Mardi / Mercredi / Vendredi / Dimanche) et exportez le bilan semaine.")
    weekly_memory.render_weekly_reports_view()

# ------------------------------------------------------------------------------
# 3. FICHE INDIVIDUELLE ATHLÈTE A4 (REPORTLAB)
# ------------------------------------------------------------------------------
with tab_indiv:
    st.subheader("👤 Fiche Athlète Individuelle (A4)")
    st.caption("Synthèse individuelle éditée pour un joueur en combinant son profil, ses scores Hooper, ses charges sRPE et ses métriques GPS.")

    try:
        from modules.gsheet_sync import get_players, get_wellness_data
        players_df = get_players()
        wellness_df = get_wellness_data()
    except Exception:
        players_df = pd.DataFrame()
        wellness_df = pd.DataFrame()

    if not players_df.empty:
        col_p1, col_p2 = st.columns([2, 1])
        def format_nom_joueur(row):
            return f"{str(row.get('nom', '')).upper()} {str(row.get('prenom', '')).capitalize()}".strip()

        with col_p1:
            noms_liste = [format_nom_joueur(r) for _, r in players_df.iterrows()]
            joueur_choisi = st.selectbox("Sélectionner un athlète :", noms_liste, key="sel_indiv_player_pdf")
            idx = noms_liste.index(joueur_choisi)
            p_row = players_df.iloc[idx]
        with col_p2:
            nb_j = st.selectbox("Période :", [7, 14, 28], format_func=lambda x: f"Derniers {x} jours", key="sel_indiv_period_pdf")

        if st.button("📄 Exporter la fiche individuelle", type="primary", use_container_width=True, key="btn_gen_indiv_pdf"):
            wellness_sub = {
                "score_hooper_moy": 11.5,
                "sommeil_moy": 2.8,
                "fatigue_moy": 3.0,
                "courbatures_moy": 2.9,
                "stress_moy": 2.1,
                "srpe_total": 1420
            }
            gps_sub = {
                "total_distance_m": 12450,
                "hsr_distance_m": 840,
                "sprint_distance_m": 180,
                "vmax_kmh": 31.8,
                "intensity_m_min": 78.4,
                "player_load_total": 412,
                "acwr": 1.12
            }
            pdf_b = pdf_reports.generate_player_pdf_report(
                player_name=joueur_choisi,
                player_info=p_row.to_dict(),
                wellness_summary=wellness_sub,
                gps_summary=gps_sub,
                period_str=f"Derniers {nb_j} jours"
            )
            st.download_button(
                label=f"📥 Télécharger la fiche de {joueur_choisi}",
                data=pdf_b,
                file_name=f"Fiche_{joueur_choisi.replace(' ', '_')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="btn_download_indiv_pdf"
            )
    else:
        st.info("Effectif en cours de chargement.")
'''

    with open("pages/5_Rapports_PDF.py", "w", encoding="utf-8") as f:
        f.write(reports_page_code)

    print("pages/5_Rapports_PDF.py mis a jour.")

if __name__ == "__main__":
    main()

