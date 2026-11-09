"""Module de génération de rapports PDF professionnels avec ReportLab.

Permet de générer :
- Un bilan individuel joueur (Wellness Hooper, Charge sRPE, Métriques GPS).
- Un rapport de séance ou bilan hebdomadaire prêt pour l'impression ou l'export.
"""

from __future__ import annotations

import io
from typing import Any, Dict, Optional
import pandas as pd

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

    # Styles personnalisés
    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=20,
        leading=24,
        textColor=colors.HexColor("#1E3A8A"),  # Bleu marine RCA
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

    # En-tête du club / document
    elements.append(Paragraph("🏉 RCA - RAPPORT DE PERFORMANCE INDIVIDUEL", title_style))
    elements.append(Paragraph(f"Période d'évaluation : {period_str} | Date d'émission : {pd.Timestamp.now().strftime('%d/%m/%Y')}", subtitle_style))
    elements.append(Spacer(1, 0.4 * cm))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1E3A8A"), spaceAfter=12))

    # Fiche Joueur (Tableau 2 colonnes)
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

    # Section 1 : État de Forme & Wellness (Hooper)
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
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
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

    # Section 2 : Charge Externe & Données GPS
    elements.append(Paragraph("2. Données GPS & Charge Externe", section_style))

    tot_dist = gps_summary.get("total_distance_m", "N/A")
    hsr_dist = gps_summary.get("hsr_distance_m", "N/A")
    sprint_dist = gps_summary.get("sprint_distance_m", "N/A")
    vmax = gps_summary.get("vmax_kmh", "N/A")
    intensity = gps_summary.get("intensity_m_min", "N/A")
    player_load = gps_summary.get("player_load_total", "N/A")
    acwr = gps_summary.get("acwr", "N/A")

    gps_table_data = [
        ["Métrique GPS", "Valeur Cumulée / Pic", "Référence"],
        ["Distance Totale", f"{tot_dist} m" if tot_dist != "N/A" else "N/A", "Volume de déplacement global"],
        ["Course Haute Intensité (> 20 km/h)", f"{hsr_dist} m" if hsr_dist != "N/A" else "N/A", "HSR (> 5.5 m/s)"],
        ["Distance Sprint (> 25 km/h)", f"{sprint_dist} m" if sprint_dist != "N/A" else "N/A", "Vitesse pure"],
        ["Vitesse Maximale atteinte", f"{vmax} km/h" if vmax != "N/A" else "N/A", "VMax pic de la période"],
        ["Intensité moyenne", f"{intensity} m/min" if intensity != "N/A" else "N/A", "Rythme de jeu moyen"],
        ["PlayerLoad Total", str(player_load), "Charge accélérométrique (3D)"],
        ["Ratio ACWR (7j / 28j)", str(acwr), "Zone optimale : 0.8 - 1.3"]
    ]

    t_gps = Table(gps_table_data, colWidths=[6.5 * cm, 4.5 * cm, 7 * cm])
    t_gps.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0D9488")),  # Vert canard / Teal pro
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

    # Notes et recommandations staff
    elements.append(Paragraph("3. Synthèse Staff & Préconisations", section_style))
    reco_text = (
        "Athlète dans les normes de charge attendues. Maintenir la vigilance sur la récupération "
        "et la qualité du sommeil post-match. Adapter les séances de vitesse si le ratio ACWR dépasse 1.4."
    )
    elements.append(Paragraph(f"<i>{reco_text}</i>", body_style))

    # Construction du PDF
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()
