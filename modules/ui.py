"""
modules/ui.py
Module universel d'habillage UI, de bannière officielle et de mise en page pour le Rugby Club Amiénois (RCA).
Charte officielle :
- Bleu RCA : #1F4E79
- Or/Jaune RCA : #D4AF37
- Fond dégradé : #F4F6F9 vers #E9EEF5
"""

import base64
import functools
import os
import streamlit as st

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load_tactile_css() -> str:
    """Charge le CSS tactile centralisé depuis assets/style.css."""
    candidates = [
        os.path.join(PROJECT_ROOT, "assets", "style.css"),
        "assets/style.css",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return f.read()
            except Exception:
                pass
    return ""


def inject_rca_tactile_ui() -> None:
    """Injecte la surcouche CSS tactile RCA (inputs, pastilles, sliders, cartes athlètes, boutons) depuis assets/style.css."""
    css = _load_tactile_css()
    if css:
        st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def get_base64_logo() -> str:
    """Cherche le logo RCA local s'il existe, sinon fallback transparent ou SVG / GitHub."""
    possible_paths = [
        os.path.join(PROJECT_ROOT, "assets", "logo_rca.png"),
        os.path.join(PROJECT_ROOT, "logo_rca.png"),
        os.path.join(PROJECT_ROOT, "assets", "logo.png"),
        "assets/logo_rca.png",
        "logo_rca.png",
        "assets/logo.png",
    ]
    for p in possible_paths:
        if os.path.exists(p):
            try:
                with open(p, "rb") as f:
                    return f"data:image/png;base64,{base64.b64encode(f.read()).decode()}"
            except Exception:
                pass
    return "https://raw.githubusercontent.com/Quent123456/rca-unified-monitoring/main/assets/logo_rca.png"


def inject_rca_page_layout(
    page_title: str = "RCA Unified Monitoring",
    page_subtitle: str = "Pôle Performance & Préparation Physique",
    enable_tactile_ui: bool = True,
) -> None:
    """
    Injecte les styles officiels RCA :
    1. Fond global dégradé
    2. Sidebar dégradée Bleu RCA (#1F4E79) avec liseré doré (#D4AF37)
    3. Bandeau supérieur officiel (Header Banner RCA) avec logo et typographie
    4. Cartes d'accueil interactives avec animation au survol (hover)
    5. Surcouche CSS tactile mobile/desktop (assets/style.css) si enable_tactile_ui=True
    """
    logo_src = get_base64_logo()

    st.markdown(
        f"""
    <style>
        /* 1. Fond global de l'application (dégradé bleu-gris très doux + maillage fin discret) */
        html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stApp"] {{
            background-color: #E9EEF5 !important;
            background-image:
                linear-gradient(rgba(31, 78, 121, 0.045) 1px, transparent 1px),
                linear-gradient(90deg, rgba(31, 78, 121, 0.045) 1px, transparent 1px),
                linear-gradient(135deg, #F4F6F9 0%, #E9EEF5 55%, #DFE7F2 100%) !important;
            background-size: 32px 32px, 32px 32px, 100% 100% !important;
            background-attachment: fixed !important;
        }}
        [data-testid="stMain"],
        [data-testid="stMainBlockContainer"],
        section.main,
        [data-testid="stHeader"],
        [data-testid="stAppHeader"] {{
            background: transparent !important;
            background-color: transparent !important;
        }}
        
        /* 2. Sidebar stylisée aux couleurs du RCA avec largeur 320px, logo, respiration et onglets agrandis */
        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #1F4E79 0%, #153452 100%) !important;
            border-right: 3px solid #D4AF37 !important;
        }}
        @media (min-width: 768px) {{
            section[data-testid="stSidebar"]:not([aria-expanded="false"]) {{
                width: 320px !important;
                min-width: 320px !important;
                max-width: 320px !important;
            }}
            section[data-testid="stSidebar"]:not([aria-expanded="false"]) > div {{
                width: 320px !important;
            }}
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {{
            padding: 0.75rem 1.1rem 1.5rem 1.1rem !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarHeader"] {{
            height: 2.25rem !important;
            min-height: 2.25rem !important;
            margin-bottom: 0.25rem !important;
            padding: 0.25rem 0.25rem 0 0.25rem !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] svg {{
            color: #FFFFFF !important;
            fill: #FFFFFF !important;
        }}
        section[data-testid="stSidebar"] h1,
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3,
        section[data-testid="stSidebar"] p,
        section[data-testid="stSidebar"] label,
        section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
        section[data-testid="stSidebar"] .st-emotion-cache-17lntkn {{
            color: #FFFFFF !important;
        }}

        /* Logo officiel RCA centré + sous-titre "RCA Performance" au sommet de la navigation */
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"]::before {{
            content: "RCA Performance";
            display: block;
            width: 100%;
            background-image: url("{logo_src}");
            background-repeat: no-repeat;
            background-position: center 2px;
            background-size: 94px auto;
            padding-top: 130px;
            padding-bottom: 14px;
            margin-bottom: 18px;
            text-align: center;
            font-size: 1.12rem;
            font-weight: 800;
            letter-spacing: 0.7px;
            color: #D4AF37 !important;
            border-bottom: 1px solid rgba(212, 175, 55, 0.35);
            filter: drop-shadow(0 2px 6px rgba(0, 0, 0, 0.28));
        }}

        /* Espacement vertical aéré des liens pour remplir harmonieusement la sidebar */
        section[data-testid="stSidebar"] [data-testid="stSidebarNavItems"] {{
            display: flex !important;
            flex-direction: column !important;
            gap: 14px !important;
            padding: 4px 0 18px 0 !important;
            margin: 0 !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarNavItems"] > li {{
            margin: 0 !important;
            padding: 0 !important;
        }}

        /* Onglets agrandis : hauteur, padding et police renforcés */
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a,
        section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"] {{
            min-height: 54px !important;
            padding: 12px 16px !important;
            border-radius: 10px !important;
            background-color: rgba(255, 255, 255, 0.06) !important;
            border: 1px solid rgba(255, 255, 255, 0.09) !important;
            border-left: 4px solid transparent !important;
            margin: 0 !important;
            display: flex !important;
            align-items: center !important;
            text-decoration: none !important;
            transition: all 0.22s cubic-bezier(0.4, 0, 0.2, 1) !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a span,
        section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"] span {{
            color: #F8FAFC !important;
            font-size: 1.02rem !important;
            font-weight: 700 !important;
            line-height: 1.35 !important;
            letter-spacing: 0.25px !important;
            transition: color 0.22s ease !important;
        }}

        /* Effet au survol : fond bleu foncé + liseré et texte or #D4AF37 */
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a:hover,
        section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"]:hover {{
            background-color: #102A43 !important;
            border: 1px solid rgba(212, 175, 55, 0.45) !important;
            border-left: 5px solid #D4AF37 !important;
            transform: translateX(4px) !important;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.22) !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a:hover span,
        section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"]:hover span {{
            color: #D4AF37 !important;
        }}

        /* Style actif net : fond bleu nuit foncé + liseré et texte jaune or #D4AF37 */
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"],
        section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"][aria-current="page"] {{
            background-color: #0B2036 !important;
            border: 1px solid rgba(212, 175, 55, 0.65) !important;
            border-left: 5px solid #D4AF37 !important;
            box-shadow: 0 4px 14px rgba(0, 0, 0, 0.28) !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] span,
        section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"][aria-current="page"] span {{
            color: #D4AF37 !important;
            font-weight: 700 !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarNavSeparator"] {{
            border-bottom-color: rgba(212, 175, 55, 0.28) !important;
            margin: 0.75rem 0 !important;
        }}
        section[data-testid="stSidebar"] button[kind="primary"],
        section[data-testid="stSidebar"] button[kind="primary"] * {{
            background-color: #D4AF37 !important;
            color: #1F4E79 !important;
            font-weight: 700 !important;
            border: none !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] {{
            background-color: rgba(255, 255, 255, 0.12) !important;
            color: #FFFFFF !important;
            border: 1px solid rgba(255, 255, 255, 0.28) !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"],
        section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] * {{
            color: #1F4E79 !important;
        }}
        
        /* 3. Bandeau supérieur (Header Banner RCA) */
        .rca-top-banner {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            background: linear-gradient(90deg, #1F4E79 0%, #296499 100%);
            border-bottom: 3px solid #D4AF37;
            border-radius: 12px;
            padding: 14px 24px;
            margin-bottom: 25px;
            box-shadow: 0 4px 15px rgba(31, 78, 121, 0.18);
        }}
        .rca-top-banner-content {{
            display: flex;
            align-items: center;
            gap: 18px;
        }}
        .rca-top-banner img {{
            height: 60px;
            filter: drop-shadow(0px 2px 4px rgba(0,0,0,0.3));
        }}
        .rca-top-banner-text h2 {{
            color: #FFFFFF !important;
            font-size: 1.4rem;
            margin: 0;
            font-weight: 700;
            letter-spacing: 0.5px;
        }}
        .rca-top-banner-text p {{
            color: #D4AF37 !important;
            margin: 2px 0 0 0;
            font-size: 0.9rem;
            font-weight: 500;
        }}
        .rca-badge-pill {{
            background: rgba(212, 175, 55, 0.15);
            border: 1px solid #D4AF37;
            color: #FFFFFF !important;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 0.8rem;
            font-weight: 600;
        }}

        /* 4. Cartes d'accueil interactives avec animation au survol (hover) */
        .rca-card {{
            background: #FFFFFF;
            border-radius: 12px;
            padding: 20px;
            border: 1px solid #E1E8F0;
            border-top: 4px solid #1F4E79;
            box-shadow: 0 4px 12px rgba(0,0,0,0.05);
            transition: all 0.28s cubic-bezier(0.4, 0, 0.2, 1);
            cursor: pointer;
            height: 100%;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            box-sizing: border-box;
            min-height: 215px;
        }}
        .rca-card:hover {{
            transform: translateY(-6px) scale(1.02);
            box-shadow: 0 12px 24px rgba(31, 78, 121, 0.18);
            border-top: 4px solid #D4AF37;
        }}
        .rca-card:hover .rca-card-btn {{
            background-color: #D4AF37 !important;
            color: #1F4E79 !important;
            border-color: #D4AF37 !important;
        }}
        .rca-card h3 {{
            color: #1F4E79 !important;
            font-size: 1.15rem;
            font-weight: 700;
            margin-top: 0;
            margin-bottom: 8px;
        }}
        .rca-card p {{
            color: #556980 !important;
            font-size: 0.88rem;
            line-height: 1.45;
            margin-bottom: 1rem;
        }}
        .rca-card-btn {{
            background-color: #1F4E79;
            color: #FFFFFF;
            border: 1px solid #1F4E79;
            padding: 0.45rem 1rem;
            border-radius: 6px;
            font-size: 0.85rem;
            font-weight: 700;
            text-align: center;
            width: fit-content;
            transition: all 0.2s ease;
        }}

        /* Cartes génériques blanches de la page d'accueil */
        .stCard {{
            background-color: #FFFFFF !important;
            border-radius: 10px !important;
            border: 1px solid rgba(31, 78, 121, 0.14) !important;
            box-shadow: 0 4px 14px rgba(31, 78, 121, 0.05) !important;
        }}
    </style>

    <div class="rca-top-banner">
        <div class="rca-top-banner-content">
            <img src="{logo_src}" alt="RCA Logo" />
            <div class="rca-top-banner-text">
                <h2>{page_title}</h2>
                <p>{page_subtitle}</p>
            </div>
        </div>
        <div class="rca-badge-pill">
            🏉 RCA Fédérale 2
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    if enable_tactile_ui:
        inject_rca_tactile_ui()

