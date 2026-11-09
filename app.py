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
    .rca-card-box {
        border: 1.5px solid rgba(49, 51, 63, 0.2);
        border-radius: 0.65rem;
        padding: 1.25rem;
        margin-bottom: 0.8rem;
        background-color: #FFFFFF;
        transition: all 0.2s ease-in-out;
        cursor: pointer;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        min-height: 210px;
        height: 100%;
        box-sizing: border-box;
    }
    .rca-card-box:hover {
        transform: translateY(-4px);
        border-color: #1F4E79 !important;
        box-shadow: 0 10px 24px rgba(31, 78, 121, 0.18) !important;
    }
    @media (prefers-color-scheme: dark) {
        .rca-card-box {
            background-color: #1A202C !important;
            border-color: #2D3748 !important;
        }
        .rca-card-box p {
            color: #A0AEC0 !important;
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

# ── Grille des 5 modules principaux (Cartes HTML 100% cliquables) ──────────────
st.subheader("🚀 Modules de Performance & Suivi")
st.caption("Cliquez directement sur une carte pour ouvrir le module correspondant :")

def render_module_card(title: str, desc: str, url_path: str, btn_text: str) -> str:
    return f"""
<a href="{url_path}" target="_self" style="text-decoration: none; color: inherit; display: block; height: 100%;">
  <div class="rca-card-box"
       onmouseover="this.style.transform='translateY(-4px)'; this.style.borderColor='#1F4E79'; this.style.boxShadow='0 10px 24px rgba(31,78,121,0.18)';"
       onmouseout="this.style.transform='translateY(0)'; this.style.borderColor='rgba(49, 51, 63, 0.2)'; this.style.boxShadow='0 2px 8px rgba(0,0,0,0.04)';">
    <div>
      <h3 style="margin-top: 0; margin-bottom: 0.45rem; color: #1F4E79; font-size: 1.15rem; font-weight: 700;">{title}</h3>
      <p style="font-size: 0.91rem; color: #555; margin-bottom: 0.9rem; line-height: 1.45;">{desc}</p>
    </div>
    <div style="display: inline-block; background-color: #1F4E79; color: white; padding: 0.4rem 0.85rem; border-radius: 5px; font-size: 0.85rem; font-weight: 600; text-align: center; width: fit-content;">{btn_text} →</div>
  </div>
</a>
"""

row1_col1, row1_col2, row1_col3 = st.columns(3)

with row1_col1:
    st.markdown(
        render_module_card(
            title="📝 1. Saisie Quotidienne",
            desc="Formulaire mobile RPE Foster (effort × durée) & questionnaire bien-être Hooper (sommeil, fatigue, courbatures).",
            url_path="/Saisie_Quotidienne",
            btn_text="Ouvrir la saisie",
        ),
        unsafe_allow_html=True,
    )

with row1_col2:
    st.markdown(
        render_module_card(
            title="👤 2. Tableau Individuel",
            desc="Profil athlète détaillé, jauge ACWR dynamique (charge aiguë vs chronique), radar de bien-être et alertes de surmenage.",
            url_path="/Tableau_Individuel",
            btn_text="Ouvrir le profil individuel",
        ),
        unsafe_allow_html=True,
    )

with row1_col3:
    st.markdown(
        render_module_card(
            title="👥 3. Tableau Collectif",
            desc="Vue globale de l'effectif, heatmap d'entraînement, matrice de risque ACWR et comparaisons Avants / Arrières.",
            url_path="/Tableau_Collectif",
            btn_text="Consulter la vue équipe",
        ),
        unsafe_allow_html=True,
    )

st.write("")

row2_col1, row2_col2 = st.columns(2)

with row2_col1:
    st.markdown(
        render_module_card(
            title="🛰️ 4. Analyse GPS & Rapports",
            desc="Traitement Catapult One en 3 étapes : boîtiers, métriques rugby (PlayerLoad, HSR, sprints), injection Excel STAPS et rapports hebdo.",
            url_path="/Analyse_GPS",
            btn_text="Lancer l'analyse GPS",
        ),
        unsafe_allow_html=True,
    )

with row2_col2:
    st.markdown(
        render_module_card(
            title="⚙️ 5. Gestion de l'Effectif",
            desc="Administration de l'effectif synchronisée en direct avec Google Sheets : ajout, modification des postes, groupes et photos.",
            url_path="/Gestion_Joueurs",
            btn_text="Gérer l'effectif",
        ),
        unsafe_allow_html=True,
    )

# ── Pied de page sobre ────────────────────────────────────────────────────────
st.divider()
footer_c1, footer_c2 = st.columns([3, 1])
with footer_c1:
    st.caption("🏉 **Rugby Club Amiénois** — Pôle Préparation Physique & Performance Sportive")
with footer_c2:
    st.caption("Version 2.0 • Architecture Unifiée")
