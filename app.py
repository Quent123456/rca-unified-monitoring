"""
Application Streamlit - Portail d'Accueil Unifié
Club : Rugby Club Amiénois (RCA)
Description :
    Portail centralisé de monitoring athlétique intégrant :
    1. Saisie Quotidienne RPE 
    2. Tableau de Bord Individuel & ACWR
    3. Tableau de Bord Collectif & Matrice de Risque
    4. Traitement & Analyse GPS Catapult One et Rapports PDF
    5. Gestion et Administration de l'Effectif
"""

import textwrap
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
    inject_rca_page_layout(
        "RCA Unified Monitoring",
        "Plateforme centralisée RPE & Catapult GPS",
        enable_tactile_ui=False,
    )
except Exception:
    pass

# ── Grille des 5 modules principaux (Navigation native Streamlit multi-pages) ─
st.subheader("Modules de Performance & Suivi")
st.caption("Cliquez directement sur une carte pour ouvrir le module correspondant :")

MODULE_CARDS = [
    {
        "slug": "saisie",
        "title": "📝 1. Saisie Quotidienne",
        "desc": "Formulaire mobile RPE Foster (effort × durée) & questionnaire bien-être Hooper (sommeil, fatigue, courbatures).",
        "page_path": "pages/1_Saisie_Quotidienne.py",
        "btn_text": "📝 Ouvrir la Saisie →",
    },
    {
        "slug": "individuel",
        "title": "👤 2. Tableau Individuel",
        "desc": "Profil athlète détaillé, jauge ACWR dynamique (charge aiguë vs chronique), radar de bien-être et alertes de surmenage.",
        "page_path": "pages/2_Tableau_Individuel.py",
        "btn_text": "👤 Profil Individuel →",
    },
    {
        "slug": "collectif",
        "title": "👥 3. Tableau Collectif",
        "desc": "Vue globale de l'effectif, heatmap d'entraînement, matrice de risque ACWR et comparaisons Avants / Arrières.",
        "page_path": "pages/3_Tableau_Collectif.py",
        "btn_text": "👥 Voir l'Équipe →",
    },
    {
        "slug": "gps",
        "title": "🛰️ 4. Analyse GPS & Rapports",
        "desc": "Traitement Catapult One en 3 étapes : boîtiers, métriques rugby (PlayerLoad, HSR, sprints), injection Excel STAPS et rapports hebdo.",
        "page_path": "pages/4_Analyse_GPS.py",
        "btn_text": "🛰️ Lancer l'Analyse GPS →",
    },
    {
        "slug": "effectif",
        "title": "⚙️ 5. Gestion de l'Effectif",
        "desc": "Administration de l'effectif synchronisée en direct avec Google Sheets : ajout, modification des postes, groupes et photos.",
        "page_path": "pages/5_Gestion_Joueurs.py",
        "btn_text": "⚙️ Gérer l'Effectif →",
    },
]


def render_module_card(col, card: dict) -> None:
    """Affiche une carte d'accueil avec redirection multi-pages native Streamlit."""
    slug = card["slug"]
    with col:
        with st.container(key=f"home_card_{slug}"):
            card_html = textwrap.dedent(f"""
<div class="rca-card-body">
<h3>{card['title']}</h3>
<p>{card['desc']}</p>
</div>
""").strip()
            st.markdown(card_html, unsafe_allow_html=True)
            if st.button(card["btn_text"], key=f"home_btn_{slug}", use_container_width=True):
                st.switch_page(card["page_path"])


row1_cols = st.columns(3)
for col, card in zip(row1_cols, MODULE_CARDS[:3]):
    render_module_card(col, card)

st.write("")

row2_cols = st.columns(2)
for col, card in zip(row2_cols, MODULE_CARDS[3:]):
    render_module_card(col, card)

# ── Pied de page sobre ────────────────────────────────────────────────────────
st.divider()
footer_c1, footer_c2 = st.columns([3, 1])
with footer_c1:
    st.caption("🏉 **Rugby Club Amiénois** — Pôle Préparation Physique & Performance Sportive")
with footer_c2:
    st.caption("Version 2.0 • Portail Multi-Pages")
