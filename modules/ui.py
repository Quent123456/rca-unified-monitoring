"""
modules/ui.py
Module universel d'habillage UI, de bannière officielle et de mise en page pour le Rugby Club Amiénois (RCA).
Charte officielle :
- Bleu RCA : #1F4E79
- Or/Jaune RCA : #D4AF37
- Fond dégradé : #F4F6F9 vers #E9EEF5
"""

import base64
import os
import streamlit as st

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


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
) -> None:
    """
    Injecte les styles officiels RCA :
    1. Fond global dégradé
    2. Sidebar dégradée Bleu RCA (#1F4E79) avec liseré doré (#D4AF37)
    3. Bandeau supérieur officiel (Header Banner RCA) avec logo et typographie
    4. Cartes d'accueil interactives avec animation au survol (hover)
    """
    logo_src = get_base64_logo()

    st.markdown(
        f"""
    <style>
        /* 1. Fond global de l'application */
        .stApp, [data-testid="stAppViewContainer"], [data-testid="stApp"] {{
            background: linear-gradient(135deg, #F4F6F9 0%, #E9EEF5 100%) !important;
        }}
        
        /* 2. Sidebar stylisée aux couleurs du RCA */
        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #1F4E79 0%, #153452 100%) !important;
            border-right: 3px solid #D4AF37 !important;
        }}
        section[data-testid="stSidebar"] * {{
            color: #FFFFFF !important;
        }}
        section[data-testid="stSidebar"] .st-emotion-cache-17lntkn {{
            color: #FFFFFF !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a {{
            color: #FFFFFF !important;
            font-weight: 600 !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a span {{
            color: #FFFFFF !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a:hover {{
            background-color: rgba(212, 175, 55, 0.20) !important;
            border-left: 4px solid #D4AF37 !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] {{
            background-color: rgba(212, 175, 55, 0.25) !important;
            border-left: 4px solid #D4AF37 !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] span {{
            color: #D4AF37 !important;
            font-weight: 700 !important;
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

        /* Conteneurs et formulaires lisibles */
        div[data-testid="stVerticalBlockBorderWrapper"],
        div[data-testid="stForm"],
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
