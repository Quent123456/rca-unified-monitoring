"""
modules/database.py
Gestion SQLite — joueurs et sessions d'entraînement.
"""

import sqlite3
import os
import pandas as pd
from datetime import date, timedelta

# Chemin de la base de données (même répertoire que app.py)
DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "training.db")


def _get_connection() -> sqlite3.Connection:
    """Retourne une connexion SQLite avec row_factory."""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    """Crée les tables si elles n'existent pas et insère des joueurs de démonstration."""
    conn = _get_connection()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS players (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            nom      TEXT NOT NULL,
            prenom   TEXT NOT NULL,
            poste    TEXT NOT NULL,
            numero   INTEGER,
            groupe   TEXT DEFAULT 'Groupe Élargi'
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            player_id   INTEGER NOT NULL,
            session_date TEXT NOT NULL,
            rpe         REAL NOT NULL,
            duration    REAL NOT NULL,
            foster_load REAL NOT NULL,
            fatigue     INTEGER NOT NULL,
            courbatures INTEGER NOT NULL,
            sommeil     INTEGER NOT NULL,
            notes       TEXT,
            groupe_entrainement TEXT DEFAULT 'Équipe A',
            FOREIGN KEY (player_id) REFERENCES players(id)
        )
    """)


    conn.commit()
    conn.close()


# ── JOUEURS ────────────────────────────────────────────────────────────────────

def get_players() -> pd.DataFrame:
    """Retourne tous les joueurs sous forme de DataFrame."""
    conn = _get_connection()
    df = pd.read_sql_query(
        "SELECT * FROM players ORDER BY nom, prenom", conn
    )
    conn.close()
    return df


def add_player(nom: str, prenom: str, poste: str, numero: int, groupe: str = "Groupe Élargi") -> None:
    """Ajoute un nouveau joueur."""
    conn = _get_connection()
    conn.execute(
        "INSERT INTO players (nom, prenom, poste, numero, groupe) VALUES (?,?,?,?,?)",
        (nom, prenom, poste, numero, groupe),
    )
    conn.commit()
    conn.close()


def update_player(player_id: int, nom: str, prenom: str, poste: str, groupe: str) -> None:
    """Met à jour les informations d'un joueur existant."""
    conn = _get_connection()
    conn.execute(
        "UPDATE players SET nom=?, prenom=?, poste=?, groupe=? WHERE id=?",
        (nom, prenom, poste, groupe, player_id)
    )
    conn.commit()
    conn.close()


def delete_player(player_id: int) -> None:
    """Supprime un joueur et toutes ses sessions."""
    conn = _get_connection()
    conn.execute("DELETE FROM sessions WHERE player_id = ?", (player_id,))
    conn.execute("DELETE FROM players WHERE id = ?", (player_id,))
    conn.commit()
    conn.close()


# ── SESSIONS ───────────────────────────────────────────────────────────────────

def add_session(
    player_id: int,
    session_date: str,
    rpe: float,
    duration: float,
    foster_load: float,
    fatigue: int,
    courbatures: int,
    sommeil: int,
    notes: str = "",
    groupe_entrainement: str = "Équipe A",
) -> None:
    """Enregistre une session d'entraînement."""
    conn = _get_connection()
    # Vérification doublon (même joueur, même date)
    existing = conn.execute(
        "SELECT id FROM sessions WHERE player_id=? AND session_date=?",
        (player_id, session_date),
    ).fetchone()
    if existing:
        conn.execute(
            """UPDATE sessions
               SET rpe=?, duration=?, foster_load=?, fatigue=?,
                   courbatures=?, sommeil=?, notes=?, groupe_entrainement=?
               WHERE player_id=? AND session_date=?""",
            (rpe, duration, foster_load, fatigue, courbatures, sommeil, notes, groupe_entrainement,
             player_id, session_date),
        )
    else:
        conn.execute(
            """INSERT INTO sessions
               (player_id, session_date, rpe, duration, foster_load,
                fatigue, courbatures, sommeil, notes, groupe_entrainement)
               VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (player_id, session_date, rpe, duration, foster_load,
             fatigue, courbatures, sommeil, notes, groupe_entrainement),
        )
    conn.commit()
    conn.close()


def get_sessions(player_id: int, days: int = 35) -> pd.DataFrame:
    """
    Retourne les sessions d'un joueur sur les N derniers jours.
    Inclut des lignes à 0 pour les jours sans séance (repos).
    """
    conn = _get_connection()
    df = pd.read_sql_query(
        """SELECT session_date, rpe, duration, foster_load,
                  fatigue, courbatures, sommeil, notes
           FROM sessions
           WHERE player_id = ?
           ORDER BY session_date""",
        conn,
        params=(player_id,),
    )
    conn.close()

    if df.empty:
        return df

    df["session_date"] = pd.to_datetime(df["session_date"])

    # Compléter avec tous les jours (repos = 0)
    end_date = date.today()
    start_date = end_date - timedelta(days=days - 1)
    all_dates = pd.date_range(start=start_date, end=end_date, freq="D")
    df_full = pd.DataFrame({"session_date": all_dates})
    df_full = df_full.merge(df, on="session_date", how="left")
    df_full["foster_load"] = df_full["foster_load"].fillna(0)
    df_full["rpe"] = df_full["rpe"].fillna(0)
    df_full["duration"] = df_full["duration"].fillna(0)
    # Bien-être : on garde NaN pour les jours sans saisie
    return df_full


def get_all_sessions_team(days: int = 35) -> pd.DataFrame:
    """Retourne toutes les sessions de l'équipe avec le nom des joueurs."""
    conn = _get_connection()
    df = pd.read_sql_query(
        """SELECT s.session_date, s.foster_load, s.rpe, s.duration,
                  s.fatigue, s.courbatures, s.sommeil, s.groupe_entrainement,
                  p.id as player_id,
                  p.nom || ' ' || p.prenom as player_name,
                  p.poste, p.groupe
           FROM sessions s
           JOIN players p ON s.player_id = p.id
           ORDER BY s.session_date""",
        conn,
    )
    conn.close()
    if not df.empty:
        df["session_date"] = pd.to_datetime(df["session_date"])
    return df
