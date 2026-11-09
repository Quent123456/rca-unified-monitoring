"""Module de synchronisation avec Google Sheets (avec mode démo / données simulées).

Ce module gère :
1. L'authentification via les secrets Streamlit (`st.secrets["gcp_service_account"]`).
2. La lecture et l'écriture des données (Effectif, Wellness, Séances GPS).
3. Un fallback automatique vers des données simulées (mock data) si aucun compte de service n'est configuré.
"""

from __future__ import annotations

import datetime
from typing import Any, Dict, List, Optional
import pandas as pd
import streamlit as st

# Liste standard des postes Rugby pour l'effectif RCA
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


def is_gsheet_configured() -> bool:
    """Vérifie si les identifiants Google Sheets sont configurés dans st.secrets."""
    return "gcp_service_account" in st.secrets and "spreadsheet_url" in st.secrets


def get_connection():
    """Initialise et retourne le client gspread si configuré, sinon None."""
    if not is_gsheet_configured():
        return None
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=scopes
        )
        return gspread.authorize(creds)
    except Exception as e:
        st.warning(f"Erreur de connexion Google Sheets: {e}. Mode démo activé.")
        return None


def get_players() -> pd.DataFrame:
    """Retourne la liste des joueurs (effectif)."""
    if "mock_players" not in st.session_state:
        st.session_state["mock_players"] = pd.DataFrame(MOCK_PLAYERS)
    return st.session_state["mock_players"]


def get_wellness_data(days_back: int = 30) -> pd.DataFrame:
    """Récupère les données d'évaluation quotidienne (Wellness Hooper).

    Colonnes : Date, Joueur_ID, Nom, Sommeil, Fatigue, Courbatures, Stress, RPE, Duree_min, sRPE.
    """
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
    client = get_connection()
    if client and "spreadsheet_url" in st.secrets:
        try:
            sheet = client.open_by_url(st.secrets["spreadsheet_url"]).worksheet("Wellness")
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
            return True
        except Exception as e:
            st.error(f"Échec de l'enregistrement Google Sheets : {e}")

    # Fallback Session State
    df = get_wellness_data()
    new_df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
    st.session_state["mock_wellness"] = new_df
    return True
