"""
Application Streamlit - Portail d'Accueil Unifié
Club : Rugby Club Amiénois (RCA)
Description :
    Portail centralisé de monitoring athlétique intégrant :
    1. Saisie Quotidienne RPE & Wellness (Foster & Hooper)
    2. Tableau de Bord Individuel & ACWR
    3. Tableau de Bord Collectif & Matrice de Risque
    4. Traitement & Analyse GPS Catapult One, Modèle STAPS et Rapports PDF
    5. Gestion et Administration de l'Effectif
"""

import os
import streamlit as st

st.set_page_config(
    page_title="RCA Unified Monitoring",
    page_icon="🏉",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Styles CSS personnalisés pour l'identité visuelle RCA ─────────────────────
st.markdown(
    """
    <style>
    .stApp {
        font-family: 'Segoe UI', Montserrat, Tahoma, sans-serif;
    }
    .main-title {
        color: #1F4E79;
        font-size: 2.3rem;
        font-weight: 800;
        margin-bottom: 0.15rem;
        letter-spacing: -0.5px;
        line-height: 1.15;
    }
    .sub-title {
        color: #4A5568;
        font-size: 1.1rem;
        margin-bottom: 0.4rem;
        font-weight: 400;
    }
    .badge-club {
        display: inline-block;
        background-color: #EBF3FA;
        color: #1F4E79;
        font-weight: 700;
        font-size: 0.85rem;
        padding: 4px 12px;
        border-radius: 20px;
        border: 1px solid #BFD7EA;
        margin-bottom: 0.8rem;
    }
    .module-card-title {
        color: #1F4E79;
        font-size: 1.18rem;
        font-weight: 700;
        margin-bottom: 0.35rem;
    }
    .module-card-desc {
        color: #556270;
        font-size: 0.92rem;
        line-height: 1.45;
        min-height: 58px;
        margin-bottom: 0.75rem;
    }

    /* ── Cartes de modules entièrement cliquables avec effet hover ── */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(a[data-testid="stPageLink-NavLink"]) {
        position: relative;
        transition: all 0.2s ease-in-out;
        cursor: pointer;
        border-radius: 12px;
        background-color: #FFFFFF;
        border: 1.5px solid #E2E8F0;
    }

    div[data-testid="stVerticalBlockBorderWrapper"]:has(a[data-testid="stPageLink-NavLink"]):hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 22px rgba(31, 78, 121, 0.12), 0 2px 6px rgba(0, 0, 0, 0.04);
        border-color: #1F4E79 !important;
    }

    /* Étirer le lien a[data-testid="stPageLink-NavLink"] sur l'intégralité du conteneur de la carte */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(a[data-testid="stPageLink-NavLink"]) a[data-testid="stPageLink-NavLink"]::after {
        content: "";
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        z-index: 10;
        border-radius: 12px;
    }

    /* Bouton d'action harmonisé en bas de chaque carte */
    div[data-testid="stVerticalBlockBorderWrapper"] a[data-testid="stPageLink-NavLink"] {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
        background-color: #F0F4F8;
        color: #1F4E79 !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
        padding: 8px 14px !important;
        border-radius: 8px !important;
        text-decoration: none !important;
        border: 1px solid #D0DFEE;
        transition: all 0.2s ease;
        margin-top: auto;
    }

    div[data-testid="stVerticalBlockBorderWrapper"]:has(a[data-testid="stPageLink-NavLink"]):hover a[data-testid="stPageLink-NavLink"] {
        background-color: #1F4E79 !important;
        color: #FFFFFF !important;
        border-color: #1F4E79 !important;
    }

    @media (prefers-color-scheme: dark) {
        div[data-testid="stVerticalBlockBorderWrapper"]:has(a[data-testid="stPageLink-NavLink"]) {
            background-color: #1A202C;
            border-color: #2D3748;
        }
        div[data-testid="stVerticalBlockBorderWrapper"]:has(a[data-testid="stPageLink-NavLink"]):hover {
            border-color: #63B3ED !important;
            box-shadow: 0 10px 22px rgba(0, 0, 0, 0.4);
        }
        div[data-testid="stVerticalBlockBorderWrapper"] a[data-testid="stPageLink-NavLink"] {
            background-color: #2D3748;
            color: #90CDF4 !important;
            border-color: #4A5568;
        }
        div[data-testid="stVerticalBlockBorderWrapper"]:has(a[data-testid="stPageLink-NavLink"]):hover a[data-testid="stPageLink-NavLink"] {
            background-color: #2B6CB0 !important;
            color: #FFFFFF !important;
            border-color: #2B6CB0 !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── En-tête officiel avec Logo RCA ────────────────────────────────────────────
logo_path = None
for candidate in ["assets/logo_rca.png", "logo_rca.png"]:
    if os.path.exists(candidate):
        logo_path = candidate
        break

header_col1, header_col2 = st.columns([1, 8])

with header_col1:
    if logo_path:
        st.image(logo_path, width=105)
    else:
        st.markdown("<div style='font-size: 4rem; text-align: center;'>🏉</div>", unsafe_allow_html=True)

with header_col2:
    st.markdown('<div class="badge-club">🏉 Rugby Club Amiénois • Fédérale 2</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-title">RCA Unified Monitoring</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-title">Plateforme centralisée de suivi athlétique — RPE, Wellness & Métriques GPS Catapult</div>',
        unsafe_allow_html=True,
    )

st.divider()

# ── Grille des 5 modules principaux ───────────────────────────────────────────
st.subheader("🚀 Modules de Performance & Suivi")
st.caption("Sélectionnez un espace pour accéder aux formulaires, tableaux de bord et outils d'analyse :")

row1_col1, row1_col2, row1_col3 = st.columns(3)

with row1_col1:
    with st.container(border=True):
        st.markdown('<div class="module-card-title">📝 1. Saisie Quotidienne</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="module-card-desc">'
            'Formulaire joueur ergonomique (mobile) : RPE Foster (effort × durée), '
            'questionnaire bien-être Hooper (sommeil, courbatures, fatigue) et mémorisation de groupe.'
            '</div>',
            unsafe_allow_html=True,
        )
        st.page_link(
            "pages/1_Saisie_Quotidienne.py",
            label="Accéder à la Saisie",
            icon="📝",
            use_container_width=True,
        )

with row1_col2:
    with st.container(border=True):
        st.markdown('<div class="module-card-title">👤 2. Tableau Individuel</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="module-card-desc">'
            'Profil athlète détaillé, jauge ACWR dynamique (aiguë vs chronique), '
            'radar de bien-être, historique de charge et alertes de surmenage.'
            '</div>',
            unsafe_allow_html=True,
        )
        st.page_link(
            "pages/2_Tableau_Individuel.py",
            label="Ouvrir le Profil Individuel",
            icon="👤",
            use_container_width=True,
        )

with row1_col3:
    with st.container(border=True):
        st.markdown('<div class="module-card-title">👥 3. Tableau Collectif</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="module-card-desc">'
            'Vue globale de l\'effectif, heatmap d\'entraînement, matrice de risque ACWR, '
            'comparaisons Avants / Arrières et détection des baisses de forme.'
            '</div>',
            unsafe_allow_html=True,
        )
        st.page_link(
            "pages/3_Tableau_Collectif.py",
            label="Consulter la Vue Équipe",
            icon="👥",
            use_container_width=True,
        )

st.write("")

row2_col1, row2_col2 = st.columns(2)

with row2_col1:
    with st.container(border=True):
        st.markdown('<div class="module-card-title">🛰️ 4. Analyse GPS & Rapports</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="module-card-desc">'
            'Processeur Catapult One en 3 étapes : attribution des boîtiers, métriques rugby '
            '(PlayerLoad, HSR, sprints), injection dans le modèle officiel STAPS et rapports hebdomadaires.'
            '</div>',
            unsafe_allow_html=True,
        )
        st.page_link(
            "pages/4_Analyse_GPS.py",
            label="Lancer l'Analyse GPS",
            icon="🛰️",
            use_container_width=True,
        )

with row2_col2:
    with st.container(border=True):
        st.markdown('<div class="module-card-title">⚙️ 5. Gestion de l\'Effectif</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="module-card-desc">'
            'Administration de l\'effectif synchronisée en direct avec Google Sheets : '
            'ajout de nouveaux joueurs, mise à jour des postes, groupes d\'entraînement et photos.'
            '</div>',
            unsafe_allow_html=True,
        )
        st.page_link(
            "pages/5_Gestion_Joueurs.py",
            label="Gérer l'Effectif",
            icon="⚙️",
            use_container_width=True,
        )

# ── Pied de page sobre ────────────────────────────────────────────────────────
st.divider()
footer_c1, footer_c2 = st.columns([3, 1])
with footer_c1:
    st.caption("🏉 **Rugby Club Amiénois** — Pôle Préparation Physique & Performance Sportive")
with footer_c2:
    st.caption("Version 2.0 • Architecture Unifiée")
