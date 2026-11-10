"""
modules/ui_theme.py
Module centralisé d'identité visuelle et d'habillage UI pour le Rugby Club Amiénois (RCA).

Fournit :
1. apply_rca_theme() : Thème global sobre et moderne (fond fixe bandes angulaires RCA, filigrane, cartes contrastées).
2. render_rca_header() : Bandeau d'en-tête officiel RCA au sommet de chaque page.
3. inject_rca_css() : Styles complets (navigation sidebar, typographie, cartes, interactions).
4. render_sidebar_header() : Logo officiel centré, titre officiel et badge Fédérale 2 au sommet de la sidebar.
5. render_sidebar_footer() : Boîte d'état Pôle Performance, statuts systèmes, objectif hebdo et signature.
6. setup_sidebar() : Initialise l'ensemble de l'habillage sidebar et CSS en un seul appel.
"""

import base64
import functools
import os
import streamlit as st

try:
    from modules.ui import inject_rca_page_layout, inject_rca_tactile_ui
except Exception:
    inject_rca_page_layout = None
    inject_rca_tactile_ui = None

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@functools.lru_cache(maxsize=1)
def get_rca_logo_base64() -> str:
    """Charge et encode en Base64 le logo officiel RCA pour affichage HTML instantané."""
    candidates = [
        os.path.join(PROJECT_ROOT, "assets", "logo_rca.png"),
        os.path.join(PROJECT_ROOT, "logo_rca.png"),
        "assets/logo_rca.png",
        "logo_rca.png",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                with open(path, "rb") as f:
                    return base64.b64encode(f.read()).decode("utf-8")
            except Exception:
                pass
    return ""


def apply_rca_theme() -> None:
    """Applique le thème officiel RCA (fond fixe bandes RCA angulaires, filigrane et cartes contrastées)."""
    b64_logo = get_rca_logo_base64()
    logo_watermark_url = (
        f"data:image/png;base64,{b64_logo}"
        if b64_logo
        else "https://raw.githubusercontent.com/Quent123456/rca-unified-monitoring/main/assets/logo_rca.png"
    )

    st.markdown(
        f"""
        <style>
        /* Fond fixe avec bandes angulaires bleu/jaune et filigrane */
        html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"], .stApp {
            background-color: #F5F7FB !important;
            background-image: 
                /* Bandes graphiques sportives angulaires */
                linear-gradient(135deg, rgba(22, 43, 76, 0.05) 0%, rgba(22, 43, 76, 0.05) 25%, transparent 25%, transparent 50%, rgba(243, 195, 0, 0.06) 50%, rgba(243, 195, 0, 0.06) 75%, transparent 75%, transparent),
                /* Dégradé doux d'ambiance */
                linear-gradient(to bottom, #FFFFFF 0%, #EEF2F8 100%) !important;
            background-size: 120px 120px, 100% 100% !important;
            background-attachment: fixed !important;
        }

        /* En option : filigrane centré très discret du logo RCA */
        [data-testid="stAppViewContainer"]::before {{
            content: "";
            position: fixed;
            top: 50%;
            left: 55%;
            transform: translate(-50%, -50%);
            width: 520px;
            height: 520px;
            background-image: url('{logo_watermark_url}');
            background-repeat: no-repeat;
            background-position: center;
            background-size: contain;
            opacity: 0.035;
            pointer-events: none;
            z-index: 0;
        }}

        /* Les conteneurs, tableaux et cartes restent blancs et lisibles à 100% */
        div[data-testid="stVerticalBlockBorderWrapper"],
        div[data-testid="stForm"],
        .stCard {{
            background-color: #FFFFFF !important;
            border-radius: 10px !important;
            border: 1px solid rgba(22, 43, 76, 0.14) !important;
            box-shadow: 0 4px 18px rgba(22, 43, 76, 0.06) !important;
            position: relative;
            z-index: 1;
        }}

        /* Barre latérale Bleu RCA dégradé avec liseré or */
        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #1F4E79 0%, #153452 100%) !important;
            border-right: 3px solid #D4AF37 !important;
            position: relative;
            z-index: 10;
        }}
        section[data-testid="stSidebar"] * {{
            color: #FFFFFF !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a {{
            color: #FFFFFF !important;
            font-weight: 600 !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a span {{
            color: #FFFFFF !important;
        }}

        /* Boutons primaires aux couleurs du RCA */
        button[kind="primary"] {{
            background-color: #F3C300 !important;
            color: #162B4C !important;
            font-weight: 700 !important;
            border: none !important;
        }}
        button[kind="primary"]:hover {{
            background-color: #DEB200 !important;
            color: #162B4C !important;
        }}

        /* Boutons des cartes de la page d'accueil */
        .rca-card-btn {{
            background-color: #162B4C !important;
            color: #FFFFFF !important;
            border: 1px solid #162B4C !important;
            display: inline-block !important;
            padding: 0.45rem 1rem !important;
            border-radius: 6px !important;
            font-size: 0.85rem !important;
            font-weight: 700 !important;
            text-align: center !important;
            width: fit-content !important;
            transition: all 0.2s ease !important;
        }}
        .rca-card-box:hover .rca-card-btn {{
            background-color: #F3C300 !important;
            color: #162B4C !important;
            border-color: #F3C300 !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_rca_header(page_title: str = "") -> None:
    """Affiche le bandeau d'en-tête officiel RCA au sommet de chaque page."""
    b64_logo = get_rca_logo_base64()
    logo_html = (
        f'<img src="data:image/png;base64,{b64_logo}" style="width: 52px; height: auto; filter: drop-shadow(0 2px 4px rgba(0,0,0,0.12)); vertical-align: middle;" alt="Logo RCA" />'
        if b64_logo
        else '<span style="font-size: 2.2rem; vertical-align: middle;">🏉</span>'
    )

    badge_html = ""
    if page_title and str(page_title).strip():
        badge_html = f'''
        <span style="display: inline-block; background-color: #162B4C; color: #FFFFFF; font-size: 0.80rem; font-weight: 700; padding: 4px 12px; border-radius: 16px; border-left: 3px solid #F3C300; margin-left: 0.65rem; vertical-align: middle;">
            {page_title}
        </span>
        '''

    st.markdown(
        f"""
        <div style="background-color: #FFFFFF; border-radius: 10px; border: 1px solid rgba(22, 43, 76, 0.14); border-left: 6px solid #F3C300; padding: 0.85rem 1.3rem; margin-bottom: 1.25rem; box-shadow: 0 4px 16px rgba(22, 43, 76, 0.05); display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 0.8rem; position: relative; z-index: 2;">
            <div style="display: flex; align-items: center; gap: 1.1rem;">
                {logo_html}
                <div>
                    <div style="display: flex; align-items: center; flex-wrap: wrap; gap: 0.35rem;">
                        <span style="font-size: 1.3rem; font-weight: 900; color: #162B4C; letter-spacing: 0.6px; text-transform: uppercase;">
                            RUGBY CLUB AMIÉNOIS
                        </span>
                        {badge_html}
                    </div>
                    <div style="font-size: 0.88rem; color: #4A5568; margin-top: 0.15rem; font-weight: 500;">
                        Plateforme Officielle de Suivi de Charge, Wellness &amp; Performance Athlétique
                    </div>
                </div>
            </div>
            <div style="display: flex; align-items: center; gap: 0.45rem;">
                <span style="display: inline-block; background-color: #E8EFF8; color: #162B4C; font-weight: 700; font-size: 0.76rem; padding: 3px 10px; border-radius: 12px; border: 1px solid #BFD7EA;">
                    🏉 Fédérale 2
                </span>
                <span style="display: inline-block; background-color: #FEF3C7; color: #92400E; font-weight: 700; font-size: 0.76rem; padding: 3px 10px; border-radius: 12px; border: 1px solid #FDE68A;">
                    ⚡ Pôle Performance
                </span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def inject_rca_css() -> None:
    """Injecte la feuille de styles CSS officielle RCA dans le document Streamlit."""
    apply_rca_theme()
    st.markdown(
        """
        <style>
        /* ── Typographie globale ─────────────────────────────────────── */
        html, body, [class*="css"] {
            font-family: 'Segoe UI', Montserrat, -apple-system, BlinkMacSystemFont, Tahoma, sans-serif;
        }

        /* ── Habillage Sidebar & Navigation ─────────────────────────── */
        [data-testid="stSidebarContent"] {
            display: flex !important;
            flex-direction: column !important;
            padding-top: 1rem !important;
            padding-bottom: 1.5rem !important;
        }

        /* Permet aux enfants directs de stSidebarUserContent de participer au flex layout */
        [data-testid="stSidebarUserContent"],
        [data-testid="stSidebarUserContent"] > [data-testid="stVerticalBlock"],
        [data-testid="stSidebarUserContent"] > div {
            display: contents !important;
        }

        /* 1. En-tête Sidebar : au sommet au-dessus du menu de navigation */
        div[data-testid="stElementContainer"]:has(.rca-sidebar-header),
        div[data-testid="stVerticalBlockBorderWrapper"]:has(.rca-sidebar-header),
        div:has(.rca-sidebar-header) {
            order: -10 !important;
            margin-bottom: 0.5rem !important;
        }

        /* 2. Menu de Navigation latéral au centre */
        [data-testid="stSidebarNav"] {
            order: -5 !important;
            padding-top: 0.4rem !important;
            padding-bottom: 0.8rem !important;
            border-top: 1px solid #E2E8F0 !important;
            border-bottom: 1px solid #E2E8F0 !important;
            margin-bottom: 0.8rem !important;
        }

        [data-testid="stSidebarNav"] ul {
            padding: 0 !important;
            margin: 0 !important;
            gap: 0.35rem !important;
        }

        [data-testid="stSidebarNav"] li,
        [data-testid="stSidebarNavLinkContainer"] {
            margin-bottom: 0.3rem !important;
        }

        /* Liens de navigation : agrandis, lisibles et modernes */
        [data-testid="stSidebarNav"] li a,
        [data-testid="stSidebarNavLink"] {
            font-size: 1.05rem !important;
            font-weight: 600 !important;
            padding: 0.65rem 1rem !important;
            border-radius: 8px !important;
            color: #162B4C !important;
            border-left: 4px solid transparent !important;
            transition: all 0.2s ease-in-out !important;
            background-color: transparent !important;
            display: flex !important;
            align-items: center !important;
            text-decoration: none !important;
        }

        [data-testid="stSidebarNav"] li a span,
        [data-testid="stSidebarNavLink"] span {
            font-size: 1.02rem !important;
            font-weight: 600 !important;
            color: #162B4C !important;
            transition: color 0.2s ease-in-out !important;
        }

        /* Au survol : fond doux RCA, texte marine et liseré doré RCA */
        [data-testid="stSidebarNav"] li a:hover,
        [data-testid="stSidebarNavLink"]:hover {
            background-color: #E8EFF8 !important;
            color: #162B4C !important;
            border-left: 4px solid #F3C300 !important;
            transform: translateX(3px) !important;
        }

        [data-testid="stSidebarNav"] li a:hover span,
        [data-testid="stSidebarNavLink"]:hover span {
            color: #162B4C !important;
        }

        /* Onglet actif : fond bleu marine foncé RCA, texte blanc, liseré doré RCA */
        [data-testid="stSidebarNav"] li a[aria-current="page"],
        [data-testid="stSidebarNavLink"][aria-current="page"],
        [data-testid="stSidebarNav"] li.active a {
            background-color: #162B4C !important;
            color: #FFFFFF !important;
            border-left: 4px solid #F3C300 !important;
            box-shadow: 0 3px 8px rgba(22, 43, 76, 0.22) !important;
        }

        [data-testid="stSidebarNav"] li a[aria-current="page"] span,
        [data-testid="stSidebarNavLink"][aria-current="page"] span,
        [data-testid="stSidebarNav"] li.active a span {
            color: #FFFFFF !important;
            font-weight: 700 !important;
        }

        [data-testid="stSidebarNav"] li a[aria-current="page"] svg,
        [data-testid="stSidebarNavLink"][aria-current="page"] svg {
            fill: #F3C300 !important;
        }

        /* 3. Contenu utilisateur ordinaire (widgets page) */
        [data-testid="stSidebarUserContent"] > [data-testid="stVerticalBlock"] > div {
            order: 0;
        }

        /* 4. Pied de Sidebar : fixé en bas */
        div[data-testid="stElementContainer"]:has(.rca-sidebar-footer),
        div[data-testid="stVerticalBlockBorderWrapper"]:has(.rca-sidebar-footer),
        div:has(.rca-sidebar-footer) {
            order: 10 !important;
            margin-top: auto !important;
            padding-top: 0.8rem !important;
        }

        /* ── Styles spécifiques Accueil RCA (app.py) ─────────────────── */
        .main-title {
            color: #162B4C;
            font-size: 2.3rem;
            font-weight: 800;
            margin-bottom: 0.15rem;
            letter-spacing: -0.5px;
            line-height: 1.15;
        }
        .sub-title {
            color: #4A5568;
            font-size: 1.05rem;
            margin-bottom: 0.4rem;
            font-weight: 400;
        }
        .badge-club {
            display: inline-block;
            background-color: #E8EFF8;
            color: #162B4C;
            font-weight: 700;
            font-size: 0.82rem;
            padding: 3px 12px;
            border-radius: 20px;
            border: 1px solid #BFD7EA;
            margin-bottom: 0.6rem;
        }
        .rca-card-box {
            border: 1.5px solid rgba(22, 43, 76, 0.2);
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
            border-color: #162B4C !important;
            box-shadow: 0 12px 24px rgba(22, 43, 76, 0.18) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar_header() -> None:
    """Affiche l'en-tête officiel RCA dans la barre latérale."""
    b64_logo = get_rca_logo_base64()
    img_html = (
        f'<img src="data:image/png;base64,{b64_logo}" style="width: 135px; max-width: 85%; height: auto; margin: 0 auto 0.6rem auto; display: block; filter: drop-shadow(0 3px 6px rgba(0,0,0,0.12));" alt="Logo RCA" />'
        if b64_logo
        else '<div style="font-size: 3.5rem; text-align: center; margin-bottom: 0.4rem;">🏉</div>'
    )

    st.sidebar.markdown(
        f"""
        <div class="rca-sidebar-header" style="text-align: center; padding: 0.4rem 0 0.8rem 0;">
            {img_html}
            <div style="color: #162B4C; font-weight: 800; font-size: 1.12rem; letter-spacing: 0.5px; line-height: 1.2; text-transform: uppercase;">
                RUGBY CLUB AMIÉNOIS
            </div>
            <div style="display: inline-block; background-color: #E8EFF8; color: #162B4C; font-weight: 700; font-size: 0.76rem; padding: 2px 10px; border-radius: 12px; border: 1px solid #BFD7EA; margin-top: 0.35rem;">
                🏉 Fédérale 2
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar_footer() -> None:
    """Affiche le pied de page structuré RCA dans la barre latérale."""
    with st.sidebar.container(border=True):
        st.markdown(
            """
            <div class="rca-sidebar-footer">
                <div style="font-weight: 700; color: #162B4C; font-size: 0.92rem; margin-bottom: 0.5rem; display: flex; align-items: center; gap: 0.35rem;">
                    ⚡ Pôle Performance RCA
                </div>
                <div style="font-size: 0.8rem; color: #4A5568; line-height: 1.6; margin-bottom: 0.65rem;">
                    <div>🟢 Google Sheets : <strong style="color: #162B4C;">Connecté</strong></div>
                    <div>🟢 GPS Engine : <strong style="color: #162B4C;">Opérationnel</strong></div>
                </div>
                <div style="background-color: #F4F7FB; border-left: 3px solid #F3C300; padding: 0.5rem 0.65rem; border-radius: 4px; font-size: 0.78rem; margin-bottom: 0.65rem;">
                    <span style="font-weight: 700; color: #162B4C;">🎯 Objectif Hebdo :</span><br/>
                    <em style="color: #2D3748;">Intensité & Répétition des efforts</em>
                </div>
                <div style="font-size: 0.72rem; color: #718096; text-align: center; border-top: 1px solid #E2E8F0; padding-top: 0.45rem;">
                    Préparation Physique • Saison 2026-2027
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def setup_sidebar() -> None:
    """Active le thème officiel RCA, injecte le CSS et habille la sidebar."""
    inject_rca_css()
    render_sidebar_header()
    render_sidebar_footer()
