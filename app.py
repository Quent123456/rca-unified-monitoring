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

# ── Initialisation de la charte visuelle officielle et de la Sidebar RCA ──────
try:
    from modules.ui_theme import setup_sidebar
    setup_sidebar()
except Exception:
    # Fallback si exécution isolée
    pass

# Forçage du fond d'écran tech/data sur la racine HTML et conteneurs Streamlit
st.markdown("""
<style>
/* Forçage du fond sur la racine HTML et tous les conteneurs Streamlit */
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"], .stApp, section.main {
    background-color: #F4F7FB !important;
    background-image: radial-gradient(#162B4C 1.1px, transparent 1.1px), radial-gradient(#F3C300 1.1px, #F4F7FB 1.1px) !important;
    background-size: 28px 28px !important;
    background-position: 0 0, 14px 14px !important;
    background-attachment: fixed !important;
}

/* Conserver les cartes et conteneurs avec un fond blanc opaque net */
div[data-testid="stVerticalBlockBorderWrapper"],
div[data-testid="stForm"],
.stCard {
    background-color: #FFFFFF !important;
    border: 1px solid rgba(22, 43, 76, 0.15) !important;
    box-shadow: 0 4px 12px rgba(22, 43, 76, 0.05) !important;
}

/* Sidebar blanche contrastée */
section[data-testid="stSidebar"] {
    background-color: #FFFFFF !important;
    border-right: 1px solid rgba(22, 43, 76, 0.12) !important;
}
</style>
""", unsafe_allow_html=True)

# ── En-tête officiel avec Logo RCA ────────────────────────────────────────────
logo_path = None
for candidate in ["assets/logo_rca.png", "logo_rca.png"]:
    if os.path.exists(candidate):
        logo_path = candidate
        break

header_col1, header_col2 = st.columns([1, 8])

with header_col1:
    if logo_path:
        st.image(logo_path, width=110)
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
       onmouseover="this.style.transform='translateY(-4px)'; this.style.borderColor='#162B4C'; this.style.boxShadow='0 12px 24px rgba(22,43,76,0.18)'; const b=this.querySelector('.rca-card-btn'); if(b){{ b.style.backgroundColor='#F3C300'; b.style.color='#162B4C'; b.style.borderColor='#F3C300'; }}"
       onmouseout="this.style.transform='translateY(0)'; this.style.borderColor='rgba(22, 43, 76, 0.2)'; this.style.boxShadow='0 2px 8px rgba(0,0,0,0.04)'; const b=this.querySelector('.rca-card-btn'); if(b){{ b.style.backgroundColor='#162B4C'; b.style.color='#FFFFFF'; b.style.borderColor='#162B4C'; }}">
    <div>
      <h3 style="margin-top: 0; margin-bottom: 0.45rem; color: #162B4C; font-size: 1.15rem; font-weight: 700;">{title}</h3>
      <p style="font-size: 0.91rem; color: #4A5568; margin-bottom: 0.9rem; line-height: 1.45;">{desc}</p>
    </div>
    <div class="rca-card-btn" style="display: inline-block; background-color: #162B4C; color: #FFFFFF; border: 1px solid #162B4C; padding: 0.45rem 0.95rem; border-radius: 6px; font-size: 0.85rem; font-weight: 700; text-align: center; width: fit-content; transition: all 0.2s ease;">{btn_text} →</div>
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
