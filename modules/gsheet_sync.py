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
    {"id": "J01", "nom": "BAMIERE", "prenom": "Clément", "poste": "Demi de mêlée", "groupe": "Équipe A"},
    {"id": "J02", "nom": "VASELLI", "prenom": "Enzo", "poste": "3ème ligne", "groupe": "Équipe A"},
    {"id": "J03", "nom": "HARRY", "prenom": "Max", "poste": "Demi d'ouverture", "groupe": "Équipe A"},
    {"id": "J04", "nom": "COZIC", "prenom": "Adrien", "poste": "Pilier", "groupe": "Équipe A"},
    {"id": "J05", "nom": "LESUEUR", "prenom": "Jean-Étienne", "poste": "Talonneur", "groupe": "Équipe A"},
    {"id": "J06", "nom": "WAEYAERT", "prenom": "Thomas", "poste": "2ème ligne", "groupe": "Équipe B"},
    {"id": "J07", "nom": "SELTZER", "prenom": "Antoine", "poste": "Ailier", "groupe": "Équipe B"},
    {"id": "J08", "nom": "PRÉVOST", "prenom": "Thomas", "poste": "Arrière", "groupe": "Équipe B"},
    {"id": "J09", "nom": "VENIN", "prenom": "Benoît", "poste": "Centre", "groupe": "Équipe B"},
    {"id": "J10", "nom": "MONTI", "prenom": "Ciro", "poste": "3ème ligne", "groupe": "Équipe B"},
]

import unicodedata
import re

FICTIONAL_TEST_PLAYERS = {
    ("alldritt", "gregory"),
    ("atonio", "uini"),
    ("baille", "cyril"),
    ("bamba", "demba"),
    ("barassi", "pierre louis"),
    ("bielle biarrey", "louis"),
    ("dupont", "antoine"),
    ("fickou", "gael"),
    ("flament", "thibaud"),
    ("gros", "jean baptiste"),
    ("haouas", "mohamed"),
    ("jelonch", "anthony"),
    ("marchand", "julien"),
    ("mauvaka", "peato"),
    ("meafou", "emmanuel"),
    ("meafour", "emmanuel"),
    ("ntamack", "romain"),
    ("ollivon", "charles"),
    ("penaud", "damian"),
    ("ramos", "thomas"),
    ("woki", "cameron"),
}


def _norm_identity_token(val: str) -> str:
    txt = unicodedata.normalize("NFKD", str(val or "").strip())
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    txt = re.sub(r"[-_'’./,;]+", " ", txt.lower())
    return " ".join(txt.split())


FICTIONAL_FULLNAME_STRINGS = {
    f"{n} {p}" for n, p in FICTIONAL_TEST_PLAYERS
} | {
    f"{p} {n}" for n, p in FICTIONAL_TEST_PLAYERS
}


def is_fictional_test_player(full_name: str = "", nom: str = "", prenom: str = "") -> bool:
    """Retourne True si l'identité correspond à un joueur test fictif (XV de France / Top 14)."""
    n_norm = _norm_identity_token(nom)
    p_norm = _norm_identity_token(prenom)
    if n_norm and p_norm:
        if (n_norm, p_norm) in FICTIONAL_TEST_PLAYERS or (p_norm, n_norm) in FICTIONAL_TEST_PLAYERS:
            return True
    full_norm = _norm_identity_token(full_name or f"{n_norm} {p_norm}")
    if not full_norm:
        return False
    return full_norm in FICTIONAL_FULLNAME_STRINGS


def filter_out_test_players(df: pd.DataFrame) -> pd.DataFrame:
    """Filtre tout DataFrame d'effectif pour exclure les joueurs fictifs/tests."""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return df

    def _keep(row: pd.Series) -> bool:
        nom = row.get("nom", row.get("Nom", ""))
        prenom = row.get("prenom", row.get("Prénom", row.get("Prenom", "")))
        full = row.get("Nom ; Prénom", row.get("Nom / Prenom", row.get("Nom / Prénom", row.get("Athlete Name", row.get("Player Name", "")))))
        return not is_fictional_test_player(full_name=str(full or ""), nom=str(nom or ""), prenom=str(prenom or ""))

    mask = df.apply(_keep, axis=1)
    return df.loc[mask].copy()


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
    """Retourne la liste des joueurs (effectif).
    
    Tente de lire l'onglet 'Effectif' ou 'Joueurs' du Google Sheet,
    normalise les colonnes et effectue un tri alphabétique par NOM, Prénom.
    """
    target_sheet = st.secrets.get("spreadsheet_url") or st.secrets.get("spreadsheet_name")
    df_raw = pd.DataFrame()

    if target_sheet and is_gsheet_configured():
        for ws_name in ["Effectif", "effectif", "Joueurs", "joueurs"]:
            df_raw = get_worksheet_data(target_sheet, ws_name)
            if not df_raw.empty:
                break

    if df_raw.empty:
        if "mock_players" not in st.session_state:
            st.session_state["mock_players"] = pd.DataFrame(MOCK_PLAYERS)
        df_raw = st.session_state["mock_players"].copy()

    # Normalisation des colonnes
    col_mapping = {}
    for col in df_raw.columns:
        c_clean = str(col).strip().lower()
        if c_clean in ["nom", "last_name", "lastname"]:
            col_mapping[col] = "nom"
        elif c_clean in ["prénom", "prenom", "first_name", "firstname"]:
            col_mapping[col] = "prenom"
        elif c_clean in ["groupe", "équipe", "equipe", "group", "team"]:
            col_mapping[col] = "groupe"
        elif c_clean in ["poste", "position"]:
            col_mapping[col] = "poste"
        elif c_clean in ["id", "joueur_id", "id_joueur"]:
            col_mapping[col] = "id"

    df_clean = df_raw.rename(columns=col_mapping)
    for expected_col in ["nom", "prenom", "groupe", "poste", "id"]:
        if expected_col not in df_clean.columns:
            df_clean[expected_col] = "Inconnu" if expected_col != "id" else ""

    # Tri alphabétique NOM puis Prénom et filtrage strict des joueurs fictifs/tests
    df_clean = filter_out_test_players(df_clean)
    df_clean = df_clean.sort_values(
        by=["nom", "prenom"],
        key=lambda s: s.astype(str).str.upper()
    ).reset_index(drop=True)

    return df_clean


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
    """Enregistre une nouvelle saisie quotidienne dans Google Sheets (onglet Réponses/Wellness) ou Session State."""
    client = get_gsheet_client()
    target_sheet = st.secrets.get("spreadsheet_url") or st.secrets.get("spreadsheet_name")

    if client and target_sheet:
        try:
            if target_sheet.startswith("http"):
                sh = client.open_by_url(target_sheet)
            else:
                sh = client.open(target_sheet)

            available_titles = [w.title for w in sh.worksheets()]
            sheet = None
            for candidate in ["Réponses", "Reponses", "RPE", "Wellness", "Form Responses 1"]:
                if candidate in available_titles:
                    sheet = sh.worksheet(candidate)
                    break
            if not sheet:
                sheet = sh.worksheet(available_titles[0])

            row = [
                str(entry.get("Date", datetime.date.today().isoformat())),
                str(entry.get("Joueur_ID", "")),
                str(entry.get("Nom", "")),
                str(entry.get("Prenom", "")),
                str(entry.get("Poste", "")),
                str(entry.get("Groupe", "")),
                str(entry.get("Session_Type", "Rugby Collectif")),
                entry.get("Duree_min", 60),
                entry.get("RPE", 5),
                entry.get("Charge_Foster", entry.get("Charge_sRPE", 300)),
                entry.get("Sommeil", 3),
                entry.get("Fatigue", 3),
                entry.get("Courbatures", 3),
                entry.get("Stress", 3),
                entry.get("Score_Hooper", 12),
                str(entry.get("Douleurs_Localisation", "Aucune")),
                str(entry.get("Commentaires", ""))
            ]
            sheet.append_row(row)
            # Invalidation immédiate du cache de Streamlit
            st.cache_data.clear()
            return True
        except Exception as e:
            st.error(f"Échec de l'enregistrement Google Sheets : {e}")

    # Fallback Session State (Mode Démo / Hors ligne)
    df = get_wellness_data()
    entry_dict = dict(entry)
    if "Charge_sRPE" not in entry_dict and "Charge_Foster" in entry_dict:
        entry_dict["Charge_sRPE"] = entry_dict["Charge_Foster"]
    new_df = pd.concat([df, pd.DataFrame([entry_dict])], ignore_index=True)
    st.session_state["mock_wellness"] = new_df
    st.cache_data.clear()
    return True
