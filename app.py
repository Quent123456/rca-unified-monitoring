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

# ── Injection du layout officiel RCA & Bannière universelle ───────────────────
try:
    from modules.ui import inject_rca_page_layout
    inject_rca_page_layout("RCA Unified Monitoring", "Plateforme centralisée RPE & Catapult GPS")
except Exception:
    pass

# ── Grille des 5 modules principaux (Cartes HTML 100% cliquables) ──────────────
st.subheader("🚀 Modules de Performance & Suivi")
st.caption("Cliquez directement sur une carte pour ouvrir le module correspondant :")

def render_module_card(title: str, desc: str, url_path: str, btn_text: str) -> str:
    return f"""
<a href="{url_path}" target="_self" style="text-decoration: none; color: inherit; display: block; height: 100%;">
  <div class="rca-card">
    <div>
      <h3>{title}</h3>
      <p>{desc}</p>
    </div>
    <div class="rca-card-btn">{btn_text}</div>
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
            btn_text="📝 Ouvrir la Saisie →",
        ),
        unsafe_allow_html=True,
    )

with row1_col2:
    st.markdown(
        render_module_card(
            title="👤 2. Tableau Individuel",
            desc="Profil athlète détaillé, jauge ACWR dynamique (charge aiguë vs chronique), radar de bien-être et alertes de surmenage.",
            url_path="/Tableau_Individuel",
            btn_text="👤 Profil Individuel →",
        ),
        unsafe_allow_html=True,
    )

with row1_col3:
    st.markdown(
        render_module_card(
            title="👥 3. Tableau Collectif",
            desc="Vue globale de l'effectif, heatmap d'entraînement, matrice de risque ACWR et comparaisons Avants / Arrières.",
            url_path="/Tableau_Collectif",
            btn_text="👥 Voir l'Équipe →",
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
            btn_text="🛰️ Lancer l'Analyse GPS →",
        ),
        unsafe_allow_html=True,
    )

with row2_col2:
    st.markdown(
        render_module_card(
            title="⚙️ 5. Gestion de l'Effectif",
            desc="Administration de l'effectif synchronisée en direct avec Google Sheets : ajout, modification des postes, groupes et photos.",
            url_path="/Gestion_Joueurs",
            btn_text="⚙️ Gérer l'Effectif →",
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
