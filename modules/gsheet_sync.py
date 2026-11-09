"""Module de synchronisation et de gestion du cache Google Sheets pour RCA Monitoring.

Ce module gère :
1. L'authentification via `st.secrets` (support de "gcp_service_account" et "connections.gsheets").
2. La récupération optimisée des données avec cache (@st.cache_data(ttl=600)) pour respecter les quotas de requêtes.
3. Le stockage et la mise à jour des saisies (Wellness & GPS).
4. Un fallback automatique vers des données de démonstration en local.
"""

from __future__ import annotations

import datetime
import json
from typing import Any, Dict, List, Optional
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st

# Scopes nécessaires pour Google Drive & Google Sheets
GSHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Postes standard Rugby
RUGBY_POSITIONS = [
    "Pilier gauche", "Talonneur", "Pilier droit",
    "Deuxième ligne", "Troisième ligne aile", "Troisième ligne centre",
    "Demi de mêlée", "Demi d'ouverture",
    "Ailier gauche", "Premier centre", "Deuxième centre", "Ailier droit",
    "Arrière"
]

MOCK_PLAYERS = [
    {"id": "J01", "nom": "Dupont", "prenom": "Antoine", "poste": "Demi de mêlée", "groupe": "Arrières"},
    {"id": "J02", "nom": "Alldritt", "prenom": "Grégory", "poste": "Troisième ligne centre", "groupe": "Avants"},
    {"id": "J03", "nom": "Ntamack", "prenom": "Romain", "poste": "Demi d'ouverture", "groupe": "Arrières"},
    {"id": "J04", "nom": "Atonio", "prenom": "Uini", "poste": "Pilier droit", "groupe": "Avants"},
    {"id": "J05", "nom": "Marchand", "prenom": "Julien", "poste": "Talonneur", "groupe": "Avants"},
    {"id": "J06", "nom": "Flament", "prenom": "Thibaud", "poste": "Deuxième ligne", "groupe": "Avants"},
    {"id": "J07", "nom": "Penaud", "prenom": "Damian", "poste": "Ailier droit", "groupe": "Arrières"},
    {"id": "J08", "nom": "Ramos", "prenom": "Thomas", "poste": "Arrière", "groupe": "Arrières"},
    {"id": "J09", "nom": "Fickou", "prenom": "Gaël", "poste": "Premier centre", "groupe": "Arrières"},
    {"id": "J10", "nom": "Ollivon", "prenom": "Charles", "poste": "Troisième ligne aile", "groupe": "Avants"},
]


def get_gsheet_client() -> Optional[gspread.Client]:
    """Récupère les credentials depuis st.secrets et initialise le client gspread.

    Prend en charge :
    - st.secrets["gcp_service_account"]
    - st.secrets["connections"]["gsheets"]
    """
    creds_dict: Optional[Dict[str, Any]] = None

    # Format 1 : st.secrets["gcp_service_account"]
    if "gcp_service_account" in st.secrets:
        try:
            creds_dict = dict(st.secrets["gcp_service_account"])
        except Exception as e:
            st.error(f"Erreur de lecture de 'gcp_service_account' dans st.secrets : {e}")

    # Format 2 : st.secrets["connections"]["gsheets"]
    elif "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        conn_conf = st.secrets["connections"]["gsheets"]
        try:
            if isinstance(conn_conf, dict):
                if "project_id" in conn_conf and "private_key" in conn_conf:
                    creds_dict = dict(conn_conf)
                elif "service_account" in conn_conf:
                    sa = conn_conf["service_account"]
                    creds_dict = dict(sa) if isinstance(sa, dict) else json.loads(sa)
        except Exception as e:
            st.error(f"Erreur de lecture de 'connections.gsheets' dans st.secrets : {e}")

    if not creds_dict:
        return None

    try:
        credentials = Credentials.from_service_account_info(creds_dict, scopes=GSHEET_SCOPES)
        return gspread.authorize(credentials)
    except Exception as e:
        st.error(f"Échec de l'authentification Google Sheets : {e}")
        return None


def is_gsheet_configured() -> bool:
    """Vérifie si un compte de service est configuré dans st.secrets."""
    return ("gcp_service_account" in st.secrets) or (
        "connections" in st.secrets and "gsheets" in st.secrets["connections"]
    )


def get_connection() -> Optional[gspread.Client]:
    """Alias pour get_gsheet_client()."""
    return get_gsheet_client()


@st.cache_data(ttl=600)
def get_worksheet_data(spreadsheet_name: str, worksheet_name: str) -> pd.DataFrame:
    """Récupère les données d'un onglet Google Sheets sous forme de DataFrame pandas.

    Utilise le cache st.cache_data(ttl=600) pour respecter les quotas de requêtes Google (429 Rate Limit).
    L'argument spreadsheet_name peut être le titre, l'URL complète ou l'ID de la feuille de calcul.
    """
    client = get_gsheet_client()
    if not client:
        return pd.DataFrame()

    try:
        if spreadsheet_name.startswith("http"):
            sh = client.open_by_url(spreadsheet_name)
        elif len(spreadsheet_name) > 30 and "/" not in spreadsheet_name and " " not in spreadsheet_name:
            try:
                sh = client.open_by_key(spreadsheet_name)
            except Exception:
                sh = client.open(spreadsheet_name)
        else:
            sh = client.open(spreadsheet_name)

        worksheet = sh.worksheet(worksheet_name)
        records = worksheet.get_all_records()
        return pd.DataFrame(records)
    except Exception as e:
        st.warning(f"Impossible de lire l'onglet '{worksheet_name}' de '{spreadsheet_name}' : {e}")
        return pd.DataFrame()


def get_players() -> pd.DataFrame:
    """Retourne la liste des joueurs (effectif)."""
    # Si Google Sheets est configuré, tenter de lire l'onglet 'Effectif'
    target_sheet = st.secrets.get("spreadsheet_url") or st.secrets.get("spreadsheet_name")
    if target_sheet and is_gsheet_configured():
        df_sheet = get_worksheet_data(target_sheet, "Effectif")
        if not df_sheet.empty:
            return df_sheet

    if "mock_players" not in st.session_state:
        st.session_state["mock_players"] = pd.DataFrame(MOCK_PLAYERS)
    return st.session_state["mock_players"]


def get_wellness_data(days_back: int = 30) -> pd.DataFrame:
    """Récupère les données de Wellness (Hooper + sRPE)."""
    target_sheet = st.secrets.get("spreadsheet_url") or st.secrets.get("spreadsheet_name")
    if target_sheet and is_gsheet_configured():
        df_sheet = get_worksheet_data(target_sheet, "Wellness")
        if not df_sheet.empty:
            return df_sheet

    if "mock_wellness" not in st.session_state:
        today = datetime.date.today()
        records = []
        players = get_players()

        import numpy as np
        np.random.seed(42)

        for day_offset in range(14, -1, -1):
            current_date = today - datetime.timedelta(days=day_offset)
            if current_date.weekday() == 6:
                continue

            for _, player in players.iterrows():
                sommeil = int(np.random.choice([1, 2, 3, 4], p=[0.2, 0.5, 0.25, 0.05]))
                fatigue = int(np.random.choice([1, 2, 3, 4], p=[0.2, 0.45, 0.3, 0.05]))
                courbatures = int(np.random.choice([1, 2, 3, 4], p=[0.15, 0.5, 0.3, 0.05]))
                stress = int(np.random.choice([1, 2, 3, 4], p=[0.4, 0.45, 0.1, 0.05]))

                rpe = int(np.random.choice([3, 4, 5, 6, 7, 8], p=[0.1, 0.2, 0.3, 0.25, 0.1, 0.05]))
                duree = 75 if current_date.weekday() in [1, 3] else 60

                records.append({
                    "Date": current_date.isoformat(),
                    "Joueur_ID": player["id"],
                    "Nom": f"{player['prenom']} {player['nom']}",
                    "Poste": player["poste"],
                    "Groupe": player["groupe"],
                    "Sommeil": sommeil,
                    "Fatigue": fatigue,
                    "Courbatures": courbatures,
                    "Stress": stress,
                    "Score_Hooper": sommeil + fatigue + courbatures + stress,
                    "RPE": rpe,
                    "Duree_min": duree,
                    "Charge_sRPE": rpe * duree,
                    "Douleurs_Localisation": "Aucune" if np.random.rand() > 0.15 else "Ischios"
                })

        st.session_state["mock_wellness"] = pd.DataFrame(records)

    return st.session_state["mock_wellness"]


def save_wellness_entry(entry: Dict[str, Any]) -> bool:
    """Enregistre une nouvelle saisie quotidienne (Google Sheets ou Session State)."""
    client = get_gsheet_client()
    target_sheet = st.secrets.get("spreadsheet_url") or st.secrets.get("spreadsheet_name")

    if client and target_sheet:
        try:
            if target_sheet.startswith("http"):
                sh = client.open_by_url(target_sheet)
            else:
                sh = client.open(target_sheet)

            sheet = sh.worksheet("Wellness")
            row = [
                str(entry.get("Date", datetime.date.today().isoformat())),
                str(entry.get("Joueur_ID", "")),
                str(entry.get("Nom", "")),
                str(entry.get("Poste", "")),
                str(entry.get("Groupe", "")),
                entry.get("Sommeil", 3),
                entry.get("Fatigue", 3),
                entry.get("Courbatures", 3),
                entry.get("Stress", 3),
                entry.get("Score_Hooper", 12),
                entry.get("RPE", 5),
                entry.get("Duree_min", 60),
                entry.get("Charge_sRPE", 300),
                str(entry.get("Douleurs_Localisation", "Aucune")),
                str(entry.get("Commentaires", ""))
            ]
            sheet.append_row(row)
            # Invalider le cache de lecture pour recharger les données fraîches
            get_worksheet_data.clear()
            return True
        except Exception as e:
            st.error(f"Échec de l'enregistrement Google Sheets : {e}")

    # Fallback Session State
    df = get_wellness_data()
    new_df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
    st.session_state["mock_wellness"] = new_df
    return True
