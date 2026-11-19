"""
Module de Connexion & d'Intégration RPE (Google Sheets)
Club : Rugby Club Amiénois (RCA) - Catapult GPS Processor
Description :
    Connecte en lecture seule l'application de traitement GPS au Google Sheet
    de l'application de suivi de charge RPE (suivi-charge-rca).
    Permet d'enrichir automatiquement :
    1. Le Rapport Journalier (séance du jour) avec le RPE moyen par groupe et équipe.
    2. Le Rapport Hebdomadaire (cumul de la semaine et historique multi-semaines).

Règles de jointure & robustesse :
    - Clé de jointure : Nom + Prénom + Date (nettoyage, minuscules, suppression des accents et espaces).
    - Attribution des postes : Exclusivement le tableau GPS Catapult (les postes du sheet RPE sont ignorés).
    - Gestion des valeurs manquantes : 'NR' (Non Renseigné) pour l'individuel, et ignorées (skipna=True) pour les moyennes.
    - Mise en cache Streamlit : @st.cache_data(ttl=300) pour préserver les quotas Google Sheets API (429 Rate Limit).
"""

from __future__ import annotations

import datetime
import difflib
import os
import re
import unicodedata
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd
import streamlit as st

# Scopes Google Sheets en lecture seule (sécurité maximale)
GSHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Noms courants de la feuille de calcul et des onglets de réponses RPE
DEFAULT_SPREADSHEET_NAME = "suivi-charge-rca"
COMMON_WORKSHEET_NAMES = [
    "Saisies",
    "Form Responses 1",
    "Réponses au formulaire 1",
    "RPE",
    "Charge",
    "Données",
    "Donnees",
    "Sheet1",
]


# ==============================================================================
# 1. NORMALISATION DU TEXTE & DES DATES (CORRESPONDANCE ULTRA-ROBUSTE)
# ==============================================================================

JOURS_FR_STANDARDS: List[str] = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"]

JOURS_MAP_CLEAN: Dict[str, str] = {
    "lundi": "LUNDI", "lun": "LUNDI", "monday": "LUNDI", "mon": "LUNDI",
    "mardi": "MARDI", "mar": "MARDI", "tuesday": "MARDI", "tue": "MARDI",
    "mercredi": "MERCREDI", "mer": "MERCREDI", "wednesday": "MERCREDI", "wed": "MERCREDI",
    "jeudi": "JEUDI", "jeu": "JEUDI", "thursday": "JEUDI", "thu": "JEUDI",
    "vendredi": "VENDREDI", "ven": "VENDREDI", "friday": "VENDREDI", "fri": "VENDREDI",
    "samedi": "SAMEDI", "sam": "SAMEDI", "saturday": "SAMEDI", "sat": "SAMEDI",
    "dimanche": "DIMANCHE", "dim": "DIMANCHE", "sunday": "DIMANCHE", "sun": "DIMANCHE",
}


def normalize_player_identity(raw_name: Any) -> Tuple[str, List[str]]:
    r"""Nettoyage et normalisation automatique de toutes les chaînes d'identité :
    - Convertir en minuscules et supprimer les accents (unicodedata NFKD).
    - Supprimer toute annotation entre parenthèses ou crochets (ex: (TIJ), (U21), [A]) via regex r"\([^)]*\)|\[[^\]]*\]".
    - Supprimer la ponctuation (., -, ', _, ;, :, /) et normaliser les espaces multiples.
    - Découper en liste de tokens (split).
    """
    if raw_name is None or pd.isna(raw_name):
        return "", []
    s = str(raw_name).strip()
    # Supprimer parenthèses et crochets avec leur contenu (ex: (TIJ), [A])
    s = re.sub(r"\([^)]*\)|\[[^\]]*\]", " ", s)
    # Décomposer et supprimer les accents
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    # Supprimer la ponctuation
    s = re.sub(r"[._\-',;:/]+", " ", s)
    # Minuscules et espaces
    s = " ".join(s.lower().split())
    tokens = s.split()
    return s, tokens


def normalize_text_key(text: Any) -> str:
    """Normalise une chaîne de texte (wrapper rétrocompatible)."""
    norm, _ = normalize_player_identity(text)
    return norm


def normalize_player_name_key(name: Any) -> str:
    """Normalise un nom de joueur (wrapper rétrocompatible en majuscules)."""
    norm, _ = normalize_player_identity(name)
    return norm.upper()


def parse_week_number(val: Any) -> Optional[int]:
    """
    Extrait le numéro de semaine (entier entre 1 et 53).
    Gère: 45, 45.0, '45', 'S45', 'W43', '2026-W43', 'Semaine 45', 'Semaine n° 45', etc.
    """
    if val is None or pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        try:
            w = int(round(float(val)))
            if 1 <= w <= 53 and abs(float(val) - w) < 0.01:
                return w
        except Exception:
            pass
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "nat"):
        return None
    m_iso = re.search(r"(?:19|20)\d{2}\s*[-_ ]?[WwSs]\s*(\d{1,2})\b", s)
    if m_iso:
        w = int(m_iso.group(1))
        if 1 <= w <= 53:
            return w
    m_sw = re.search(r"\b(?:[WwSs]|semaine|sem\.?)\s*n?[°o]?\s*(\d{1,2})\b", s, flags=re.IGNORECASE)
    if m_sw:
        w = int(m_sw.group(1))
        if 1 <= w <= 53:
            return w
    m = re.search(r"\b(\d{1,2})\b", s) or re.search(r"(\d{1,2})", s)
    if m:
        try:
            w = int(m.group(1))
            if 1 <= w <= 53:
                return w
        except Exception:
            pass
    return None


def normalize_weekday_str(val: Any) -> Optional[str]:
    """
    Normalise une chaîne de jour de semaine en français majuscule ('MARDI', 'MERCREDI', etc.).
    Supporte les formats 'Mardi', 'mar', 'tuesday', '2026-11-06', datetime.date, etc.
    """
    if val is None or pd.isna(val):
        return None
    if isinstance(val, (datetime.date, datetime.datetime)):
        idx = val.weekday()
        if 0 <= idx < len(JOURS_FR_STANDARDS):
            return JOURS_FR_STANDARDS[idx]
    s = str(val).strip().lower()
    if not s or s in ("nan", "none", "nat"):
        return None
    s_clean = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s_clean = re.sub(r"[^a-z]+", "", s_clean)
    if s_clean in JOURS_MAP_CLEAN:
        return JOURS_MAP_CLEAN[s_clean]
    for k, v in JOURS_MAP_CLEAN.items():
        if k in s_clean:
            return v
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            dt = datetime.datetime.strptime(str(val).strip()[:10], fmt)
            return JOURS_FR_STANDARDS[dt.weekday()]
        except Exception:
            pass
    return None


def compare_identities(norm_g: str, tokens_g: List[str], norm_c: str, tokens_c: List[str]) -> Tuple[float, str]:
    """Compare deux identités normalisées selon la cascade générique multi-critères (zéro hardcoding) :
    - Niveau 1 : Match exact ou inversion Nom/Prénom
    - Niveau 2 : Match par initiale + prénom/nom
    - Niveau 3 : Match par token unique (prénom seul ou nom seul)
    - Niveau 4 : Similarité floue (Fuzzy matching)
    """
    if not norm_g or not norm_c:
        return 0.0, "None"

    # Niveau 1 — Match exact ou inversion Nom/Prénom
    if norm_g == norm_c or (len(tokens_g) >= 2 and len(tokens_c) >= 2 and set(tokens_g) == set(tokens_c)):
        return 1.0, "Exact / Inversion"

    # Niveau 2 — Match par initiale + prénom/nom (ex: 'd anthymme' vs 'dubus anthyme')
    def _check_initial_match(t_short: List[str], t_long: List[str]) -> Tuple[float, bool]:
        inits = [t for t in t_short if len(t) == 1]
        words = [t for t in t_short if len(t) >= 3]
        if len(inits) == 1 and len(words) >= 1 and len(t_long) >= 2:
            init = inits[0]
            matching_init_cand = [t for t in t_long if t.startswith(init)]
            if matching_init_cand:
                other_long = [t for t in t_long if t not in matching_init_cand]
                best_sim = 0.0
                for w in words:
                    for o in other_long:
                        sim = difflib.SequenceMatcher(None, w, o).ratio()
                        if sim > best_sim:
                            best_sim = sim
                if best_sim >= 0.80:
                    return best_sim, True
        return 0.0, False

    sim_init, ok_init = _check_initial_match(tokens_g, tokens_c)
    if not ok_init:
        sim_init, ok_init = _check_initial_match(tokens_c, tokens_g)
    if ok_init:
        return sim_init, f"Initiale + Nom/Prénom ({sim_init:.2f})"

    # Niveau 3 — Match par token unique (prénom seul ou nom seul)
    if len(tokens_g) == 1 and len(tokens_g[0]) >= 3:
        w = tokens_g[0]
        best_sim = 0.0
        for c_tok in tokens_c:
            sim = difflib.SequenceMatcher(None, w, c_tok).ratio()
            if sim > best_sim:
                best_sim = sim
        if best_sim >= 0.85:
            return best_sim, f"Token unique ({best_sim:.2f})"

    # Niveau 4 — Similarité floue (Fuzzy matching)
    ratio_full = difflib.SequenceMatcher(None, norm_g, norm_c).ratio()

    if len(tokens_g) >= 2 and len(tokens_c) >= 2:
        tok_scores = []
        for tg in tokens_g:
            best_t = max(difflib.SequenceMatcher(None, tg, tc).ratio() for tc in tokens_c)
            tok_scores.append(best_t)
        ratio_tokens = sum(tok_scores) / len(tok_scores)
    else:
        ratio_tokens = ratio_full

    score_combined = max(ratio_full, ratio_tokens)
    if score_combined >= 0.82:
        return score_combined, f"Fuzzy matching ({score_combined:.2f})"

    return score_combined, "None"


def are_player_names_matching(name_gps: str, name_rpe: str) -> bool:
    """Vérifie si deux noms de joueurs correspondent (100% générique, sans aucun alias en dur)."""
    norm_g, tokens_g = normalize_player_identity(name_gps)
    norm_r, tokens_r = normalize_player_identity(name_rpe)
    if not norm_g or not norm_r:
        return False
    score, method = compare_identities(norm_g, tokens_g, norm_r, tokens_r)
    return method != "None" and score >= 0.82


def match_player_rpe(
    gps_name: str,
    candidates_df: Any,
    effectif_df: Optional[Any] = None,
) -> Optional[Dict[str, Any]]:
    """Associe un joueur GPS à un enregistrement RPE de manière 100% automatique et générique (zéro hardcoding).
    
    Retourne un dict avec les détails du matching ou None si non trouvé.
    """
    norm_g, tokens_g = normalize_player_identity(gps_name)
    if not norm_g:
        return None

    # Conversion des candidats en liste standardisée de dicts
    candidates: List[Dict[str, Any]] = []
    if isinstance(candidates_df, pd.DataFrame):
        for _, row in candidates_df.iterrows():
            c_name = str(row.get("Nom / Prenom", row.get("nom_complet", row.get("Nom", "")))).strip()
            rpe_val = parse_rpe_value(row.get("RPE (Chiffre)", row.get("rpe", row.get("RPE"))))
            dur_val = parse_duration_minutes(row.get("Duree (min)", row.get("duree", row.get("Duree", 60))))
            s_rpe = round(float(rpe_val) * float(dur_val), 1) if rpe_val is not None else 0.0
            candidates.append({
                "nom_complet": c_name,
                "rpe": float(rpe_val) if rpe_val is not None else None,
                "duree": float(dur_val),
                "session_rpe": s_rpe,
            })
    elif isinstance(candidates_df, list):
        for item in candidates_df:
            if isinstance(item, dict):
                c_name = str(item.get("nom_complet", item.get("Nom / Prenom", item.get("Nom", "")))).strip()
                rpe_val = item.get("rpe", item.get("RPE"))
                dur_val = item.get("duree", item.get("Duree", 60))
                s_rpe = item.get("session_rpe", round(float(rpe_val or 0) * float(dur_val or 60), 1))
                candidates.append({
                    "nom_complet": c_name,
                    "rpe": float(rpe_val) if rpe_val is not None else None,
                    "duree": float(dur_val),
                    "session_rpe": float(s_rpe),
                })

    # Étape 1 : Rapprochement direct avec les candidats RPE du jour
    ranked_candidates = []
    for c in candidates:
        norm_c, tokens_c = normalize_player_identity(c.get("nom_complet", ""))
        score, method = compare_identities(norm_g, tokens_g, norm_c, tokens_c)
        if method != "None":
            ranked_candidates.append({
                "candidate": c,
                "score": score,
                "method": method,
                "norm_c": norm_c,
            })

    if ranked_candidates:
        ranked_candidates.sort(key=lambda x: x["score"], reverse=True)
        top = ranked_candidates[0]

        # Si Niveau 3 (Token unique) : rejet si ambiguïté entre 2 candidats du jour (ex: deux "Thomas")
        if "Token unique" in top["method"]:
            if len(ranked_candidates) > 1 and ranked_candidates[1]["score"] >= 0.85:
                return None
            return {
                "matched_name": top["candidate"]["nom_complet"],
                "score": top["score"],
                "method": top["method"],
                "rpe": top["candidate"].get("rpe"),
                "duree": top["candidate"].get("duree"),
                "session_rpe": top["candidate"].get("session_rpe"),
            }

        # Si Niveau 4 (Fuzzy matching) : rejet si trop proche du 2e candidat
        if "Fuzzy matching" in top["method"]:
            if len(ranked_candidates) > 1 and (top["score"] - ranked_candidates[1]["score"] < 0.08):
                return None
            return {
                "matched_name": top["candidate"]["nom_complet"],
                "score": top["score"],
                "method": top["method"],
                "rpe": top["candidate"].get("rpe"),
                "duree": top["candidate"].get("duree"),
                "session_rpe": top["candidate"].get("session_rpe"),
            }

        return {
            "matched_name": top["candidate"]["nom_complet"],
            "score": top["score"],
            "method": top["method"],
            "rpe": top["candidate"].get("rpe"),
            "duree": top["candidate"].get("duree"),
            "session_rpe": top["candidate"].get("session_rpe"),
        }

    # Étape 2 : Enrichissement dynamique via l'onglet Effectif
    eff_list: List[Dict[str, Any]] = []
    if effectif_df is not None:
        if isinstance(effectif_df, pd.DataFrame):
            eff_list = effectif_df.to_dict(orient="records")
        elif isinstance(effectif_df, list):
            eff_list = effectif_df

    if eff_list:
        matched_eff = []
        for p in eff_list:
            nom_p = f"{p.get('nom', p.get('Nom', ''))} {p.get('prenom', p.get('Prénom', p.get('Prenom', '')))}".strip()
            norm_e, tokens_e = normalize_player_identity(nom_p)
            score_e, method_e = compare_identities(norm_g, tokens_g, norm_e, tokens_e)
            if method_e != "None":
                matched_eff.append({"player": p, "score": score_e, "method": method_e, "full_name": nom_p})

        if matched_eff:
            matched_eff.sort(key=lambda x: x["score"], reverse=True)
            # Vérifier qu'il y a unicité dans Effectif
            if len(matched_eff) == 1 or (matched_eff[0]["score"] >= 0.90 and matched_eff[0]["score"] - matched_eff[1]["score"] >= 0.10):
                target_eff_p = matched_eff[0]
                norm_eff_p, tokens_eff_p = normalize_player_identity(target_eff_p["full_name"])
                for c in candidates:
                    norm_c, tokens_c = normalize_player_identity(c.get("nom_complet", ""))
                    sc_c, meth_c = compare_identities(norm_eff_p, tokens_eff_p, norm_c, tokens_c)
                    if meth_c != "None" and sc_c >= 0.85:
                        return {
                            "matched_name": c["nom_complet"],
                            "score": sc_c,
                            "method": f"Via Effectif ({target_eff_p['method']})",
                            "rpe": c.get("rpe"),
                            "duree": c.get("duree"),
                            "session_rpe": c.get("session_rpe"),
                        }

    return None


def parse_and_normalize_date(date_val: Any) -> Optional[datetime.date]:
    """
    Convertit n'importe quel format de date (ISO, français, datetime, timestamp, float Excel)
    en un objet datetime.date standardisé.
    """
    if date_val is None or pd.isna(date_val):
        return None

    if isinstance(date_val, datetime.datetime):
        return date_val.date()
    if isinstance(date_val, datetime.date):
        return date_val

    # Si c'est un entier ou float Excel (ex: 46332)
    if isinstance(date_val, (int, float)):
        try:
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(date_val)))
        except Exception:
            pass

    s = str(date_val).strip()
    if not s or s.lower() in ["none", "nan", "nat", "-"]:
        return None

    # Si c'est un numéro de série Excel sous forme de chaîne ('46332' ou '46332.0')
    try:
        f_val = float(s)
        if 40000 <= f_val <= 60000:
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(f_val)))
    except Exception:
        pass

    # Extraction de la partie date (supprime l'heure si '06/11/2026 19:30:00')
    s_part = s.split(" ")[0].strip()

    # 1. Format JJ/MM/AAAA ou JJ-MM-AAAA
    m_fr = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$", s_part)
    if m_fr:
        d, m, y = int(m_fr.group(1)), int(m_fr.group(2)), int(m_fr.group(3))
        if y < 100:
            y += 2000
        try:
            return datetime.date(y, m, d)
        except ValueError:
            pass

    # 2. Format ISO AAAA-MM-JJ
    m_iso = re.match(r"^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})$", s_part)
    if m_iso:
        y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        try:
            return datetime.date(y, m, d)
        except ValueError:
            pass

    # 3. Format textuel avec nom de jour en préfixe (ex: 'Vendredi 06/11/2026')
    m_day_pref = re.search(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})", s)
    if m_day_pref:
        d, m, y = int(m_day_pref.group(1)), int(m_day_pref.group(2)), int(m_day_pref.group(3))
        if y < 100:
            y += 2000
        try:
            return datetime.date(y, m, d)
        except ValueError:
            pass

    # Repli avec pd.to_datetime
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.date()
    except Exception:
        pass

    return None


def parse_rpe_value(val: Any) -> Optional[float]:
    """
    Extrait et caste impérativement en float la note numérique de difficulté de séance (RPE 1-10).
    Supporte: 7, 7.5, '7', '7.5', '7/10', '7 - Très difficile', 'Score: 8', etc.
    Retourne None si la valeur est absente ou invalide.
    """
    if val is None or pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        try:
            f_val = float(val)
            if 0.0 <= f_val <= 10.0:
                return round(f_val, 1)
        except Exception:
            pass
    s = str(val).strip()
    if not s or s.upper() in ["NR", "NON RENSEIGNÉ", "NON RENSEIGNE", "-", "NAN", "NULL"]:
        return None

    # Recherche directe d'un nombre (entier ou décimal avec virgule ou point)
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:/\s*10)?", s)
    if m:
        try:
            v = float(m.group(1).replace(",", "."))
            if 0.0 <= v <= 10.0:
                return round(v, 1)
        except Exception:
            pass
    return None


def parse_duration_minutes(val: Any, default: float = 80.0) -> float:
    """Extrait la durée de séance en minutes (par défaut 80 min)."""
    if val is None or pd.isna(val):
        return default
    s = str(val).strip()
    m = re.search(r"(\d+(?:[.,]\d+)?)", s)
    if m:
        try:
            v = float(m.group(1).replace(",", "."))
            if 10.0 <= v <= 240.0:
                return v
        except Exception:
            pass
    return default


# ==============================================================================
# 2. CONNEXION GOOGLE SHEETS & CACHE STREAMLIT (@st.cache_data(ttl=300))
# ==============================================================================

def _get_secret_safely(key: str, default: Any = None) -> Any:
    """Accède aux secrets Streamlit sans lever d'exception en environnement de test."""
    try:
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return default


def get_gcp_service_account_credentials() -> Optional[Any]:
    """
    Résout les identifiants du compte de service Google Cloud :
    1. Depuis st.secrets["gcp_service_account"] (dict ou table TOML)
    2. Depuis st.secrets["connections"]["gsheets"]
    3. Depuis la variable d'environnement GCP_SERVICE_ACCOUNT (chemin fichier ou JSON)
    """
    try:
        from google.oauth2.service_account import Credentials
    except ImportError:
        return None

    # 1. Vérification dans st.secrets["gcp_service_account"]
    gcp_dict = _get_secret_safely("gcp_service_account")
    if gcp_dict:
        try:
            sec_dict = dict(gcp_dict)
            if "private_key" in sec_dict and "\\n" in sec_dict["private_key"]:
                sec_dict["private_key"] = sec_dict["private_key"].replace("\\n", "\n")
            return Credentials.from_service_account_info(sec_dict, scopes=GSHEET_SCOPES)
        except Exception as e:
            print(f"Erreur parsing st.secrets['gcp_service_account'] : {e}")

    # 2. Vérification dans st.secrets["connections"]["gsheets"]
    conn_sec = _get_secret_safely("connections")
    if isinstance(conn_sec, dict) and "gsheets" in conn_sec:
        gsh_sec = conn_sec["gsheets"]
        if isinstance(gsh_sec, dict):
            # Certains schémas incluent les clés de service account directement
            if "type" in gsh_sec and "private_key" in gsh_sec:
                try:
                    s_copy = dict(gsh_sec)
                    if "\\n" in s_copy.get("private_key", ""):
                        s_copy["private_key"] = s_copy["private_key"].replace("\\n", "\n")
                    return Credentials.from_service_account_info(s_copy, scopes=GSHEET_SCOPES)
                except Exception:
                    pass

    # 3. Variable d'environnement GCP_SERVICE_ACCOUNT
    env_sa = os.environ.get("GCP_SERVICE_ACCOUNT", "").strip()
    if env_sa:
        if os.path.exists(env_sa):
            try:
                return Credentials.from_service_account_file(env_sa, scopes=GSHEET_SCOPES)
            except Exception:
                pass
        elif env_sa.startswith("{"):
            try:
                import json
                sa_data = json.loads(env_sa)
                return Credentials.from_service_account_info(sa_data, scopes=GSHEET_SCOPES)
            except Exception:
                pass

    return None


def get_spreadsheet_target() -> str:
    """Retourne l'URL, l'ID ou le nom du Google Sheet RPE."""
    # Secrets explicites pour l'app RPE
    for k in ["RPE_GSHEET_URL", "RPE_SPREADSHEET_URL", "RPE_SPREADSHEET", "rpe_spreadsheet"]:
        v = _get_secret_safely(k)
        if v:
            return str(v).strip()

    # Secret Streamlit Connections
    conn = _get_secret_safely("connections")
    if isinstance(conn, dict) and "gsheets" in conn:
        gsh = conn["gsheets"]
        if isinstance(gsh, dict) and "spreadsheet" in gsh:
            return str(gsh["spreadsheet"]).strip()

    gsheet_direct = _get_secret_safely("gsheets")
    if isinstance(gsheet_direct, dict) and "spreadsheet" in gsheet_direct:
        return str(gsheet_direct["spreadsheet"]).strip()

    return os.environ.get("RPE_SPREADSHEET_URL", DEFAULT_SPREADSHEET_NAME)


@st.cache_data(ttl=60)
def fetch_raw_rpe_data() -> pd.DataFrame:
    csv_url = "https://docs.google.com/spreadsheets/d/1dV4evNta1SetRqvOizR-1xLAxLYvqSW_SEc6-mLr8bs/gviz/tq?tqx=out:csv&sheet=Saisies"
    try:
        df = pd.read_csv(csv_url)
        if df is not None and not df.empty:
            if "Cle de recherche" in df.columns:
                df = df.drop_duplicates(subset=["Cle de recherche"], keep="last")
            return df
    except Exception as e:
        print(f"Erreur lecture CSV direct : {e}")
    try:
        from modules.gsheets import read_saisies_as_df
        df_sec = read_saisies_as_df()
        if df_sec is not None and not df_sec.empty:
            return df_sec
    except Exception:
        pass
    return pd.DataFrame()


@st.cache_data(ttl=300)
def fetch_effectif_data() -> pd.DataFrame:
    """Charge dynamiquement l'onglet Effectif depuis Google Sheets comme table de référence vivante."""
    try:
        from modules.gsheet_sync import get_players
        df = get_players()
        if df is not None and not df.empty:
            return df
    except Exception:
        pass
    try:
        from modules.gsheets import get_spreadsheet, execute_with_retry
        sh = get_spreadsheet()
        if sh:
            ws = execute_with_retry(lambda: sh.worksheet("Effectif"))
            records = execute_with_retry(lambda: ws.get_all_records())
            return pd.DataFrame(records)
    except Exception:
        pass
    return pd.DataFrame()


def get_rpe_sync_status() -> Tuple[bool, str]:
    """Retourne l'état de connexion de l'intégration Google Sheets RPE."""
    return True, "Connecté au flux Google Sheets RPE"


# ==============================================================================
# 3. IDENTIFICATION DES COLONNES DU SHEET RPE
# ==============================================================================

def find_column_by_aliases(
    df: pd.DataFrame,
    aliases: List[str],
    exclude: Optional[List[str]] = None,
) -> Optional[str]:
    """
    Trouve la colonne correspondante dans un DataFrame selon une liste d'alias.
    Priorité absolue à la correspondance exacte normalisée pour éviter les faux positifs (ex: 'Nom' dans 'Nom et Prénom').
    """
    if df is None or df.empty:
        return None
    excl_set = {normalize_text_key(e) for e in exclude if e} if exclude else set()

    # Passe 1 : Correspondance EXACTE prioritaire
    for a in aliases:
        a_clean = normalize_text_key(a)
        for col in df.columns:
            col_clean = normalize_text_key(col)
            if col_clean in excl_set:
                continue
            if a_clean == col_clean:
                return col

    # Passe 2 : Correspondance par mot complet délimité
    for a in aliases:
        a_clean = normalize_text_key(a)
        pattern = r"\b" + re.escape(a_clean) + r"\b"
        for col in df.columns:
            col_clean = normalize_text_key(col)
            if col_clean in excl_set:
                continue
            if re.search(pattern, col_clean):
                return col

    return None


def identify_rpe_columns(df_rpe: pd.DataFrame) -> Dict[str, Optional[str]]:
    """Détecte les colonnes de semaine, jour, date, nom, prénom, RPE et durée dans le DataFrame RPE."""
    # 1. Colonne Numéro de Semaine
    col_week = find_column_by_aliases(df_rpe, [
        "Numero Semaine", "Numéro de Semaine", "Numero de Semaine", "Numéro Semaine",
        "Semaine n°", "Semaine N°", "N° Semaine", "Semaine", "Week", "Week Number", "Week_Number", "Semaine ISO", "S"
    ])

    # 2. Colonne Jour de la séance
    col_day = find_column_by_aliases(df_rpe, [
        "Jour seance", "Jour de la séance", "Jour de la seance", "Jour de séance", "Jour de seance",
        "Jour d'entraînement", "Jour d'entrainement", "Jour entraînement", "Jour entrainement",
        "Jour de la semaine", "Jour", "Weekday", "Day"
    ], exclude=[col_week] if col_week else None)

    # 3. Colonne Date / Horodateur
    col_date = find_column_by_aliases(df_rpe, [
        "Horodateur", "Date de la séance", "Date de seance", "Date séance", "Date seance",
        "Date de l'entraînement", "Date entraînement", "Date entrainement", "Date",
        "Timestamp", "Date & Heure"
    ], exclude=[c for c in [col_week, col_day] if c])

    # 4. Colonne RPE (Priorité stricte à 'RPE (Chiffre)' et 'RPE')
    excl_meta = [c for c in [col_week, col_day, col_date] if c]
    col_rpe = find_column_by_aliases(df_rpe, [
        "RPE (Chiffre)", "RPE (chiffre)", "RPE", "Score RPE", "RPE (1-10)", "Note RPE",
        "Difficulté de la séance", "Difficulte de la seance", "Difficulté", "Difficulte",
        "Intensité", "Intensite", "Charge RPE", "Note"
    ], exclude=excl_meta)

    # 5. Colonne Durée
    excl_meta_rpe = excl_meta + ([col_rpe] if col_rpe else [])
    col_duration = find_column_by_aliases(df_rpe, [
        "Duree (min)", "Durée de la séance (min)", "Durée (min)",
        "Durée de la séance", "Duree de la seance", "Durée", "Duree", "Duration", "Temps (min)", "Temps"
    ], exclude=excl_meta_rpe)

    # 6. Colonne Joueur unique (Nom complet ou Nom Prénom)
    excl_all = excl_meta_rpe + ([col_duration] if col_duration else [])
    col_joueur_unique = find_column_by_aliases(df_rpe, [
        "Nom / Prenom", "Nom / Prénom", "Nom et Prénom", "Nom & Prénom", "Nom Prénom", "Nom ; Prénom", "Nom complet",
        "Nom du Joueur", "Nom Joueur", "Joueur", "Athlète", "Athlete", "Player Name", "Player"
    ], exclude=excl_all)

    # 7. Colonnes Nom et Prénom séparées
    excl_names = excl_all + ([col_joueur_unique] if col_joueur_unique else [])
    col_nom = find_column_by_aliases(df_rpe, ["Nom de famille", "Nom"], exclude=excl_names)
    excl_prenom = excl_names + ([col_nom] if col_nom else [])
    col_prenom = find_column_by_aliases(df_rpe, ["Prénom", "Prenom"], exclude=excl_prenom)

    # 8. Colonne Poste
    excl_poste = excl_prenom + ([col_prenom] if col_prenom else [])
    col_poste = find_column_by_aliases(df_rpe, ["Poste", "Position", "Poste_Detaille", "Ligne"], exclude=excl_poste)

    return {
        "col_week": col_week,
        "col_day": col_day,
        "col_date": col_date,
        "col_nom": col_nom,
        "col_prenom": col_prenom,
        "col_joueur_unique": col_joueur_unique,
        "col_rpe": col_rpe,
        "col_duration": col_duration,
        "col_poste": col_poste,
    }


# ==============================================================================
# 4. FUSION GPS & RPE (JOINTURE SUR NOM DU JOUEUR + SEMAINE + JOUR)
# ==============================================================================

def enrich_gps_dataframe_with_rpe(
    df_gps: pd.DataFrame,
    session_date: Any,
    df_rpe: Optional[pd.DataFrame] = None,
    mapping_gps: Optional[Dict[Any, str]] = None,
    week_number: Optional[int] = None,
    weekday: Optional[str] = None,
    effectif_df: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """
    Enrichit le DataFrame GPS Catapult (df_indiv) avec les valeurs RPE individuelles :
    - Rapprochement 100% automatique et générique via match_player_rpe() (zéro hardcoding).
    - Note de difficulté RPE castée impérativement en float numérique.
    - Si un joueur n'a pas répondu : RPE = np.nan, RPE_Display = 'NR'.
    - Charge Foster calculée : Session_RPE = RPE × Durée (min).
    - Stocke les détails du matching dans df_out et dans st.session_state["rpe_matching_report"].
    """
    if df_gps is None or not isinstance(df_gps, pd.DataFrame) or df_gps.empty:
        return df_gps

    df_out = df_gps.copy()
    target_dt = parse_and_normalize_date(session_date)
    target_week = week_number
    target_weekday = normalize_weekday_str(weekday) if weekday else None

    if target_dt is not None:
        if target_week is None:
            target_week = target_dt.isocalendar()[1]
        if target_weekday is None:
            target_weekday = JOURS_FR_STANDARDS[target_dt.weekday()]
    elif isinstance(session_date, str):
        if target_weekday is None:
            target_weekday = normalize_weekday_str(session_date)
        if target_week is None:
            target_week = parse_week_number(session_date)

    c_gps_player = find_column_by_aliases(df_out, ["Nom ; Prénom", "Player Name", "Athlete Name", "Joueur", "Nom"])

    # 1. Résolution stricte des numéros de capteurs GPS vers les noms de joueurs de la séance
    if mapping_gps:
        map_lookup: Dict[Any, str] = {}
        for k, v in mapping_gps.items():
            if v and str(v).strip() and str(v).strip() not in ["— Sélectionner un joueur —", "-", "None"]:
                v_clean = str(v).strip()
                k_str = str(k).strip()
                map_lookup[k] = v_clean
                map_lookup[k_str] = v_clean
                m_dig = re.search(r"\b(\d+)\b", k_str)
                if m_dig:
                    n_i = int(m_dig.group(1))
                    map_lookup[n_i] = v_clean
                    map_lookup[str(n_i)] = v_clean
                    map_lookup[float(n_i)] = v_clean
                    map_lookup[f"{n_i}.0"] = v_clean
                    map_lookup[f"GPS {n_i}"] = v_clean
                    map_lookup[f"GPS_{n_i}"] = v_clean
                    map_lookup[f"Capteur {n_i}"] = v_clean

        def _resolve_mapped_player(val: Any) -> Any:
            if val is None or pd.isna(val):
                return val
            if val in map_lookup:
                return map_lookup[val]
            s_val = str(val).strip()
            if s_val in map_lookup:
                return map_lookup[s_val]
            m_d = re.search(r"\b(\d+)\b", s_val)
            if m_d:
                n = int(m_d.group(1))
                if n in map_lookup:
                    return map_lookup[n]
                if str(n) in map_lookup:
                    return map_lookup[str(n)]
            try:
                f_v = float(s_val)
                if f_v in map_lookup:
                    return map_lookup[f_v]
                if f_v.is_integer() and int(f_v) in map_lookup:
                    return map_lookup[int(f_v)]
            except Exception:
                pass
            return val

        # Appliquer la résolution sur toutes les colonnes candidates de joueurs
        player_candidate_cols = [c for c in ["Nom ; Prénom", "Player Name", "Athlete Name", "Joueur", "Nom", "athlete_name"] if c in df_out.columns]
        for p_col in player_candidate_cols:
            df_out[p_col] = df_out[p_col].apply(_resolve_mapped_player)

    if df_rpe is None or df_rpe.empty:
        df_rpe = fetch_raw_rpe_data()

    if effectif_df is None or (isinstance(effectif_df, pd.DataFrame) and effectif_df.empty):
        effectif_df = fetch_effectif_data()

    # Si aucune donnée RPE n'est disponible
    if df_rpe is None or df_rpe.empty:
        df_out["RPE"] = np.nan
        df_out["Session_RPE"] = np.nan
        df_out["RPE_Display"] = "NR"
        df_out["RPE_Duration"] = np.nan
        df_out["RPE_Matched_Name"] = ""
        df_out["RPE_Match_Method"] = "NR"
        df_out["RPE_Match_Score"] = 0.0
        return df_out

    cols_map = identify_rpe_columns(df_rpe)
    col_week = cols_map["col_week"]
    col_day = cols_map["col_day"]
    col_date = cols_map["col_date"]
    col_nom = cols_map["col_nom"]
    col_prenom = cols_map["col_prenom"]
    col_player = cols_map["col_joueur_unique"]
    col_rpe = cols_map["col_rpe"]
    col_dur = cols_map["col_duration"]

    # Construction des candidats RPE prioritaires (exact date/jour) et secondaires (même semaine)
    candidates_exact: List[Dict[str, Any]] = []
    candidates_week: List[Dict[str, Any]] = []

    for _, r in df_rpe.iterrows():
        r_week = parse_week_number(r[col_week]) if col_week and col_week in r else None
        r_day = normalize_weekday_str(r[col_day]) if col_day and col_day in r else None
        r_date = parse_and_normalize_date(r[col_date]) if col_date and col_date in r else None

        if r_date is not None and r_week is None:
            r_week = r_date.isocalendar()[1]

        # Extraire le nom du joueur dans Saisies
        full_name_str = ""
        if col_player and col_player in r and pd.notna(r[col_player]):
            full_name_str = str(r[col_player]).strip()
        elif col_nom and col_prenom and col_nom in r and col_prenom in r and col_nom != col_prenom:
            val_n = str(r[col_nom]).strip() if pd.notna(r[col_nom]) else ""
            val_p = str(r[col_prenom]).strip() if pd.notna(r[col_prenom]) else ""
            full_name_str = f"{val_n} {val_p}".strip()
        elif col_nom and col_nom in r and pd.notna(r[col_nom]):
            full_name_str = str(r[col_nom]).strip()

        if not full_name_str:
            continue

        try:
            from modules.gsheet_sync import is_fictional_test_player
            if is_fictional_test_player(full_name_str):
                continue
        except Exception:
            pass

        rpe_val = parse_rpe_value(r[col_rpe]) if col_rpe and col_rpe in r else None
        if rpe_val is None:
            continue

        dur_val = parse_duration_minutes(r[col_dur]) if col_dur and col_dur in r else 60.0
        s_rpe = round(float(rpe_val) * float(dur_val), 1)
        payload = {
            "nom_complet": full_name_str,
            "rpe": float(rpe_val),
            "duree": float(dur_val),
            "session_rpe": s_rpe,
        }

        # Filtrage par date/séance
        is_exact_day = (
            (target_dt is not None and r_date is not None and r_date == target_dt)
            or (target_weekday is not None and r_day is not None and r_day == target_weekday and target_week is not None and r_week == target_week)
        )
        is_same_week = (target_week is not None and r_week is not None and int(r_week) == int(target_week))

        if is_exact_day:
            candidates_exact.append(payload)
        elif is_same_week:
            candidates_week.append(payload)

    # Si des candidats pour le jour exact existent, on s'y restreint strictement.
    # Sinon (ex: date inconnue ou aucun horodateur précis), repli sur la semaine.
    pool_primary = candidates_exact if candidates_exact else candidates_week
    pool_secondary = []

    rpe_col_vals: List[Optional[float]] = []
    srpe_col_vals: List[Optional[float]] = []
    disp_col_vals: List[str] = []
    dur_col_vals: List[Optional[float]] = []
    matched_names: List[str] = []
    match_methods: List[str] = []
    match_scores: List[float] = []

    matched_report_list: List[Dict[str, Any]] = []
    unmatched_report_list: List[Dict[str, Any]] = []

    for _, row in df_out.iterrows():
        p_name = str(row[c_gps_player]).strip() if c_gps_player and c_gps_player in row else ""
        if not p_name:
            rpe_col_vals.append(np.nan)
            srpe_col_vals.append(np.nan)
            disp_col_vals.append("NR")
            dur_col_vals.append(np.nan)
            matched_names.append("")
            match_methods.append("NR")
            match_scores.append(0.0)
            continue

        match_res = match_player_rpe(p_name, pool_primary, effectif_df)
        if not match_res and pool_secondary:
            match_res = match_player_rpe(p_name, pool_secondary, effectif_df)

        if match_res:
            r_val = match_res["rpe"]
            s_rpe = match_res["session_rpe"]
            dur = match_res["duree"]
            rpe_col_vals.append(r_val)
            srpe_col_vals.append(s_rpe)
            disp_col_vals.append(f"{r_val:.1f}" if abs(r_val - round(r_val)) > 0.05 else f"{int(round(r_val))}")
            dur_col_vals.append(dur)
            matched_names.append(match_res["matched_name"])
            match_methods.append(match_res["method"])
            match_scores.append(match_res["score"])

            matched_report_list.append({
                "gps_name": p_name,
                "sheet_name": match_res["matched_name"],
                "score": match_res["score"],
                "method": match_res["method"],
                "rpe": r_val,
                "duree": dur,
                "session_rpe": s_rpe,
            })
        else:
            rpe_col_vals.append(np.nan)
            srpe_col_vals.append(np.nan)
            disp_col_vals.append("NR")
            dur_col_vals.append(np.nan)
            matched_names.append("")
            match_methods.append("NR")
            match_scores.append(0.0)

            unmatched_report_list.append({
                "gps_name": p_name,
                "reason": "Aucune saisie RPE trouvée pour cette séance",
            })

    df_out["RPE"] = rpe_col_vals
    df_out["Session_RPE"] = srpe_col_vals
    df_out["RPE_Display"] = disp_col_vals
    df_out["RPE_Duration"] = dur_col_vals
    df_out["RPE_Matched_Name"] = matched_names
    df_out["RPE_Match_Method"] = match_methods
    df_out["RPE_Match_Score"] = match_scores

    # Enregistrer le rapport de matching dans st.session_state
    try:
        st.session_state["rpe_matching_report"] = {
            "session_date": str(session_date),
            "matched": matched_report_list,
            "unmatched": unmatched_report_list,
            "total_gps": len(df_out),
            "total_matched": len(matched_report_list),
            "total_nr": len(unmatched_report_list),
        }
    except Exception:
        pass

    return df_out


# ==============================================================================
# 5. CALCUL DES MOYENNES DE GROUPES & ÉQUIPE (AVEC EXCLUSION STRICTE DES 'NR')
# ==============================================================================

def calculate_rpe_group_metrics(df_enriched: pd.DataFrame) -> Dict[str, Any]:
    """
    Calcule les moyennes RPE et de charge Session-RPE (Foster) par ligne de poste et pour l'équipe :
    - 5 de devant (1ère & 2ème lignes)
    - 3ème Ligne
    - Trois-quarts
    - Équipe A (Effectif global)
    Règle cruciale : les valeurs 'NR' (NaN) sont strictement ignorées (skipna=True)
    afin de ne pas biaiser mathématiquement les moyennes.
    """
    default_stats: Dict[str, Any] = {
        "equipe": None,
        "5 de devant": None,
        "3ème Ligne": None,
        "Trois-quarts": None,
        "foster_equipe": None,
        "foster_5_devant": None,
        "foster_3eme_ligne": None,
        "foster_trois_quarts": None,
        "rpe_moyen_equipe": None,
        "rpe_moyen_5_devant": None,
        "rpe_moyen_3eme_ligne": None,
        "rpe_moyen_trois_quarts": None,
        "nb_repondants_tot": 0,
        "nb_joueurs_tot": len(df_enriched) if df_enriched is not None else 0,
        "taux_reponse_pct": 0.0,
    }

    if df_enriched is None or not isinstance(df_enriched, pd.DataFrame) or df_enriched.empty:
        return default_stats

    if "RPE" not in df_enriched.columns:
        return default_stats

    # Conversion propre en valeurs numériques avec exclusion des NaN
    rpe_num = pd.to_numeric(df_enriched["RPE"], errors="coerce")
    srpe_num = pd.to_numeric(df_enriched.get("Session_RPE", pd.Series(dtype=float)), errors="coerce")

    # Équipe globale
    valid_mask = rpe_num.notna()
    nb_rep = int(valid_mask.sum())
    nb_tot = len(df_enriched)

    default_stats["nb_repondants_tot"] = nb_rep
    default_stats["nb_joueurs_tot"] = nb_tot
    default_stats["taux_reponse_pct"] = round((nb_rep / nb_tot * 100.0), 1) if nb_tot > 0 else 0.0

    if nb_rep > 0:
        default_stats["equipe"] = round(float(rpe_num[valid_mask].mean()), 1)
        default_stats["rpe_moyen_equipe"] = default_stats["equipe"]
        if srpe_num.notna().any():
            default_stats["foster_equipe"] = round(float(srpe_num[valid_mask].mean()), 1)

    # Lignes de postes officielles RCA (dérivées exclusivement du tableau GPS)
    col_grp = "Groupe" if "Groupe" in df_enriched.columns else None

    group_mapping = {
        "5 de devant": ["5 de devant", "devant", "1ere ligne", "2eme ligne"],
        "3ème Ligne": ["3ème ligne", "3eme ligne", "troisieme ligne"],
        "Trois-quarts": ["trois-quarts", "trois quarts", "arrieres", "arrières", "3/4"],
    }

    for grp_std, aliases in group_mapping.items():
        if col_grp:
            mask_grp = df_enriched[col_grp].astype(str).str.strip().str.lower().apply(
                lambda s: any(a in s for a in aliases)
            )
            df_g = df_enriched[mask_grp]
        else:
            df_g = pd.DataFrame()

        if not df_g.empty:
            g_rpe = pd.to_numeric(df_g["RPE"], errors="coerce").dropna()
            g_srpe = pd.to_numeric(df_g.get("Session_RPE", pd.Series(dtype=float)), errors="coerce").dropna()

            if not g_rpe.empty:
                val_grp_rpe = round(float(g_rpe.mean()), 1)
                default_stats[grp_std] = val_grp_rpe
                if grp_std == "5 de devant":
                    default_stats["rpe_moyen_5_devant"] = val_grp_rpe
                elif grp_std == "3ème Ligne":
                    default_stats["rpe_moyen_3eme_ligne"] = val_grp_rpe
                elif grp_std == "Trois-quarts":
                    default_stats["rpe_moyen_trois_quarts"] = val_grp_rpe
            if not g_srpe.empty:
                foster_key = "foster_5_devant" if grp_std == "5 de devant" else ("foster_3eme_ligne" if grp_std == "3ème Ligne" else "foster_trois_quarts")
                default_stats[foster_key] = round(float(g_srpe.mean()), 1)

    # Agrégation additionnelle par poste spécifique (9 postes canoniques)
    by_poste: Dict[str, Optional[float]] = {}
    col_pd = "Poste_Detaille" if "Poste_Detaille" in df_enriched.columns else ("Poste" if "Poste" in df_enriched.columns else None)
    if col_pd:
        tmp_postes = df_enriched[col_pd].apply(_normalize_specific_poste_rpe)
        for pos_canon in CANONICAL_9_POSTES_RPE:
            sub_p = df_enriched[tmp_postes == pos_canon]
            if not sub_p.empty:
                p_rpe = pd.to_numeric(sub_p["RPE"], errors="coerce").dropna()
                if not p_rpe.empty:
                    by_poste[pos_canon] = round(float(p_rpe.mean()), 1)
    default_stats["by_poste"] = by_poste

    return default_stats


CANONICAL_9_POSTES_RPE: List[str] = [
    "Pilier", "Talonneur", "2ème ligne", "3ème ligne",
    "Demi de mêlée", "Demi d'ouverture", "Centre", "Ailier", "Arrière",
]

POSTE_TO_GROUPE_RPE: Dict[str, str] = {
    "Pilier": "5 de devant",
    "Talonneur": "5 de devant",
    "2ème ligne": "5 de devant",
    "3ème ligne": "3ème Ligne",
    "Demi de mêlée": "Trois-quarts",
    "Demi d'ouverture": "Trois-quarts",
    "Centre": "Trois-quarts",
    "Ailier": "Trois-quarts",
    "Arrière": "Trois-quarts",
}


def _normalize_specific_poste_rpe(val: Any) -> str:
    k = normalize_text_key(val)
    if not k:
        return ""
    if any(w in k for w in ("pilier", "prop", "1ere ligne", "premiere ligne")):
        return "Pilier"
    if any(w in k for w in ("talonneur", "hooker")):
        return "Talonneur"
    if any(w in k for w in ("2eme ligne", "2e ligne", "deuxieme ligne", "seconde ligne", "lock")):
        return "2ème ligne"
    if any(w in k for w in ("3eme ligne", "3e ligne", "troisieme ligne", "flanker", "numero 8", "n8")):
        return "3ème ligne"
    if any(w in k for w in ("demi de melee", "demi de mele", "demi melee", "melee", "scrum half")):
        return "Demi de mêlée"
    if any(w in k for w in ("ouverture", "ouvreur", "fly half")):
        return "Demi d'ouverture"
    if "centre" in k:
        return "Centre"
    if any(w in k for w in ("ailier", "allier", "wing")):
        return "Ailier"
    if any(w in k for w in ("arriere", "fullback")):
        return "Arrière"
    return ""


def _resolve_rpe_row_player_and_poste(
    row: pd.Series,
    cols_map: Dict[str, Optional[str]],
    joueurs_lookup: Optional[Dict[str, str]] = None,
) -> Tuple[str, str, str]:
    """Retourne (nom_complet, poste_canonique, groupe) pour une ligne RPE brute."""
    col_player = cols_map.get("col_joueur_unique")
    col_nom = cols_map.get("col_nom")
    col_prenom = cols_map.get("col_prenom")
    col_poste = cols_map.get("col_poste")

    full_name = ""
    if col_player and col_player in row and pd.notna(row[col_player]):
        full_name = str(row[col_player]).strip()
    elif col_nom and col_prenom and col_nom in row and col_prenom in row and col_nom != col_prenom:
        vn = str(row[col_nom]).strip() if pd.notna(row[col_nom]) else ""
        vp = str(row[col_prenom]).strip() if pd.notna(row[col_prenom]) else ""
        full_name = f"{vn} {vp}".strip()
    elif col_nom and col_nom in row and pd.notna(row[col_nom]):
        full_name = str(row[col_nom]).strip()

    poste_raw = str(row[col_poste]).strip() if col_poste and col_poste in row and pd.notna(row[col_poste]) else ""
    poste_canon = _normalize_specific_poste_rpe(poste_raw)

    # Résolution via joueurs_lookup si le poste Google Sheets est manquant ou trop générique
    if not poste_canon and joueurs_lookup and full_name:
        for k_ref, v_ref in joueurs_lookup.items():
            if are_player_names_matching(full_name, k_ref):
                poste_canon = _normalize_specific_poste_rpe(v_ref)
                if poste_canon:
                    break

    grp = POSTE_TO_GROUPE_RPE.get(poste_canon, "")
    if not grp and poste_raw:
        rk = normalize_text_key(poste_raw)
        if any(t in rk for t in ("5 de devant", "devant", "avant", "cinq")):
            grp = "5 de devant"
        elif any(t in rk for t in ("3eme", "3e", "troisieme")):
            grp = "3ème Ligne"
        elif any(t in rk for t in ("trois quarts", "3 4", "arriere")):
            grp = "Trois-quarts"

    return full_name, poste_canon, grp


def filter_rpe_rows_for_week(
    df_rpe: pd.DataFrame,
    year: int,
    week_num: int,
) -> pd.DataFrame:
    """Filtre les lignes du Google Sheet Saisies correspondant à (year, week_num)."""
    if df_rpe is None or not isinstance(df_rpe, pd.DataFrame) or df_rpe.empty:
        return pd.DataFrame()

    cols_map = identify_rpe_columns(df_rpe)
    col_week = cols_map["col_week"]
    col_date = cols_map["col_date"]
    col_rpe = cols_map["col_rpe"]

    kept_idx: List[Any] = []
    for idx, row in df_rpe.iterrows():
        r_rpe = parse_rpe_value(row[col_rpe]) if col_rpe and col_rpe in row else None
        if r_rpe is None:
            continue

        r_week = parse_week_number(row[col_week]) if col_week and col_week in row else None
        r_date = parse_and_normalize_date(row[col_date]) if col_date and col_date in row else None

        match_week = False
        if r_week is not None and int(r_week) == int(week_num):
            if r_date is not None:
                iso_y, _, _ = r_date.isocalendar()
                if abs(int(iso_y) - int(year)) <= 0 or abs(int(r_date.year) - int(year)) <= 0:
                    match_week = True
            else:
                match_week = True
        elif r_week is None and r_date is not None:
            iso_y, iso_w, _ = r_date.isocalendar()
            if int(iso_y) == int(year) and int(iso_w) == int(week_num):
                match_week = True

        if match_week:
            kept_idx.append(idx)

    res = df_rpe.loc[kept_idx].copy() if kept_idx else pd.DataFrame()
    if not res.empty and "Cle de recherche" in res.columns:
        res = res.drop_duplicates(subset=["Cle de recherche"], keep="last")
    return res


def compute_direct_week_rpe_by_category(
    df_rpe: pd.DataFrame,
    year: int,
    week_num: int,
    joueurs_lookup: Optional[Dict[str, str]] = None,
) -> Dict[str, Optional[float]]:
    """
    Calcule directement les moyennes RPE de la semaine depuis Google Sheets Saisies
    pour 'Total', les 3 groupes (5 de devant, 3ème Ligne, Trois-quarts) et les 9 postes.
    """
    out: Dict[str, Optional[float]] = {
        "Total": None,
        "5 de devant": None,
        "3ème Ligne": None,
        "Trois-quarts": None,
    }
    for p in CANONICAL_9_POSTES_RPE:
        out[p] = None

    df_w = filter_rpe_rows_for_week(df_rpe, year=year, week_num=week_num)
    if df_w.empty:
        return out

    cols_map = identify_rpe_columns(df_w)
    col_rpe = cols_map["col_rpe"]

    all_vals: List[float] = []
    buckets: Dict[str, List[float]] = {k: [] for k in out if k != "Total"}

    for _, row in df_w.iterrows():
        full_name, poste_canon, grp = _resolve_rpe_row_player_and_poste(row, cols_map, joueurs_lookup=joueurs_lookup)
        if not full_name:
            continue
        try:
            from modules.gsheet_sync import is_fictional_test_player
            if is_fictional_test_player(full_name):
                continue
        except Exception:
            pass

        rpe_val = parse_rpe_value(row[col_rpe]) if col_rpe and col_rpe in row else None
        if rpe_val is None:
            continue
        f_val = float(rpe_val)
        all_vals.append(f_val)
        if grp in buckets:
            buckets[grp].append(f_val)
        if poste_canon in buckets:
            buckets[poste_canon].append(f_val)

    if all_vals:
        out["Total"] = round(float(np.mean(all_vals)), 1)
    for k, vals in buckets.items():
        if vals:
            out[k] = round(float(np.mean(vals)), 1)
    return out


def enrich_historical_week_data_with_rpe(
    week_data: Dict[str, Any],
    df_rpe: Optional[pd.DataFrame] = None,
    joueurs_lookup: Optional[Dict[str, str]] = None,
) -> Tuple[Dict[str, Any], bool]:
    """
    Rattrape dynamiquement les valeurs RPE manquantes (ou 'NR') dans un JSON hebdomadaire archivé
    (ex: S43) en interrogeant les saisies Google Sheets de la semaine ISO correspondante.
    Retourne (week_data_enrichi, modifie_bool).
    """
    if not isinstance(week_data, dict):
        return week_data, False

    year = int(week_data.get("year") or datetime.date.today().year)
    week_num = parse_week_number(week_data.get("week_num") or week_data.get("week"))
    if week_num is None:
        return week_data, False

    sessions = week_data.get("sessions", {})
    if not isinstance(sessions, dict) or not sessions:
        return week_data, False

    # Vérifie s'il manque au moins un RPE de séance ou de joueur
    needs_backfill = False
    for s_info in sessions.values():
        if not isinstance(s_info, dict):
            continue
        r_st = s_info.get("rpe_stats")
        if not isinstance(r_st, dict) or r_st.get("equipe") is None:
            needs_backfill = True
            break
        p_list = s_info.get("players_data", [])
        if isinstance(p_list, list) and p_list and not any(pd.notna(p.get("RPE")) for p in p_list if isinstance(p, dict)):
            needs_backfill = True
            break

    if not needs_backfill and isinstance(week_data.get("rpe_summary"), dict) and week_data["rpe_summary"].get("Total") is not None:
        return week_data, False

    if df_rpe is None:
        try:
            df_rpe = fetch_raw_rpe_data()
        except Exception:
            df_rpe = pd.DataFrame()

    if df_rpe is None or df_rpe.empty:
        return week_data, False

    df_w_rpe = filter_rpe_rows_for_week(df_rpe, year=year, week_num=week_num)
    if df_w_rpe.empty:
        return week_data, False

    modified = False
    for sid, s_info in sessions.items():
        if not isinstance(s_info, dict):
            continue
        p_list = s_info.get("players_data", [])
        if not isinstance(p_list, list) or not p_list:
            continue

        df_p = pd.DataFrame(p_list)
        # S'assurer de la présence du poste canonique / groupe si lookup disponible
        if joueurs_lookup and "Nom ; Prénom" in df_p.columns:
            def _enrich_poste_col(row: pd.Series) -> str:
                cur = _normalize_specific_poste_rpe(row.get("Poste_Detaille", ""))
                if cur:
                    return cur
                nm = str(row.get("Nom ; Prénom", "")).strip()
                for k_ref, v_ref in joueurs_lookup.items():
                    if are_player_names_matching(nm, k_ref):
                        return _normalize_specific_poste_rpe(v_ref) or str(row.get("Poste_Detaille", ""))
                return str(row.get("Poste_Detaille", ""))
            df_p["Poste_Detaille"] = df_p.apply(_enrich_poste_col, axis=1)

        has_existing_rpe = "RPE" in df_p.columns and pd.to_numeric(df_p["RPE"], errors="coerce").notna().any()
        if not has_existing_rpe:
            s_date = s_info.get("date") or s_info.get("date_str") or s_info.get("day_type") or ""
            s_day = s_info.get("day_type") or ""
            df_enriched = enrich_gps_dataframe_with_rpe(
                df_gps=df_p,
                session_date=s_date,
                df_rpe=df_w_rpe,
                week_number=week_num,
                weekday=normalize_weekday_str(s_day),
            )
            if "RPE" in df_enriched.columns and pd.to_numeric(df_enriched["RPE"], errors="coerce").notna().any():
                df_clean_records = df_enriched.where(pd.notna(df_enriched), None).to_dict(orient="records")
                s_info["players_data"] = df_clean_records
                s_info["rpe_stats"] = calculate_rpe_group_metrics(df_enriched)
                modified = True
        elif not isinstance(s_info.get("rpe_stats"), dict) or s_info["rpe_stats"].get("equipe") is None:
            s_info["rpe_stats"] = calculate_rpe_group_metrics(df_p)
            modified = True

    direct_summary = compute_direct_week_rpe_by_category(df_w_rpe, year=year, week_num=week_num, joueurs_lookup=joueurs_lookup)
    if direct_summary.get("Total") is not None:
        if week_data.get("rpe_summary") != direct_summary:
            week_data["rpe_summary"] = direct_summary
            modified = True

    return week_data, modified


# ==============================================================================
# 6. CALCUL RPE POUR LA SYNTHÈSE HEBDOMADAIRE (CUMUL & MOYENNES MULTI-SÉANCES)
# ==============================================================================

def calculate_weekly_rpe_from_sessions(synthese_dict: Dict[str, Any]) -> Dict[str, Any]:
    """
    Agrège les métriques RPE sur l'ensemble des séances enregistrées de la semaine :
    - Moyenne RPE hebdomadaire (Équipe, 5 de devant, 3ème Ligne, Trois-quarts)
    - Charge Foster cumulée (Session-RPE cumulé par joueur puis moyenne par groupe)
    """
    res: Dict[str, Any] = {
        "rpe_moyen_equipe": None,
        "foster_cumul_equipe": None,
        "rpe_5_devant": None,
        "rpe_3eme_ligne": None,
        "rpe_trois_quarts": None,
        "foster_5_devant": None,
        "foster_3eme_ligne": None,
        "foster_trois_quarts": None,
    }

    if not synthese_dict:
        return res

    all_players: List[Dict[str, Any]] = []
    for s_info in synthese_dict.values():
        if isinstance(s_info, dict):
            if "players_data" in s_info:
                p_data = s_info.get("players_data", [])
                if isinstance(p_data, list):
                    all_players.extend(p_data)
            elif "df_indiv" in s_info and isinstance(s_info["df_indiv"], pd.DataFrame):
                all_players.extend(s_info["df_indiv"].to_dict(orient="records"))

    if not all_players or not any("RPE" in p for p in all_players):
        # Fallback depuis les rpe_stats stockées au niveau séance
        team_rpes = []
        g_rpes = {"5 de devant": [], "3ème Ligne": [], "Trois-quarts": []}
        for s_info in synthese_dict.values():
            if isinstance(s_info, dict) and "rpe_stats" in s_info and isinstance(s_info["rpe_stats"], dict):
                r_st = s_info["rpe_stats"]
                if r_st.get("equipe") is not None:
                    team_rpes.append(r_st["equipe"])
                elif r_st.get("rpe_moyen_equipe") is not None:
                    team_rpes.append(r_st["rpe_moyen_equipe"])
                for g_key in g_rpes:
                    if r_st.get(g_key) is not None:
                        g_rpes[g_key].append(r_st[g_key])
        if team_rpes:
            res["rpe_moyen_equipe"] = round(float(np.mean(team_rpes)), 1)
            for g_key, vals in g_rpes.items():
                if vals:
                    k_out = "rpe_5_devant" if g_key == "5 de devant" else ("rpe_3eme_ligne" if g_key == "3ème Ligne" else "rpe_trois_quarts")
                    res[k_out] = round(float(np.mean(vals)), 1)
            return res
        return res

    df_all = pd.DataFrame(all_players)
    if df_all.empty or "RPE" not in df_all.columns:
        return res

    stats = calculate_rpe_group_metrics(df_all)

    res["rpe_moyen_equipe"] = stats.get("equipe")
    res["rpe_5_devant"] = stats.get("5 de devant")
    res["rpe_3eme_ligne"] = stats.get("3ème Ligne")
    res["rpe_trois_quarts"] = stats.get("Trois-quarts")

    # Calcul du cumul Foster hebdomadaire par joueur
    c_p = find_column_by_aliases(df_all, ["Nom ; Prénom", "Player Name", "Nom"])
    if c_p and "Session_RPE" in df_all.columns:
        df_all["Session_RPE_num"] = pd.to_numeric(df_all["Session_RPE"], errors="coerce")
        df_foster_p = df_all.groupby(c_p).agg({
            "Session_RPE_num": "sum",
            "Groupe": "first" if "Groupe" in df_all.columns else lambda x: "Tous",
        }).reset_index()

        valid_f = df_foster_p[df_foster_p["Session_RPE_num"] > 0]
        if not valid_f.empty:
            res["foster_cumul_equipe"] = round(float(valid_f["Session_RPE_num"].mean()), 1)

            for grp_name, foster_k in [
                ("5 de devant", "foster_5_devant"),
                ("3ème Ligne", "foster_3eme_ligne"),
                ("Trois-quarts", "foster_trois_quarts"),
            ]:
                if "Groupe" in valid_f.columns:
                    sub = valid_f[valid_f["Groupe"].str.lower().str.contains(grp_name[:4].lower(), na=False)]
                    if not sub.empty:
                        res[foster_k] = round(float(sub["Session_RPE_num"].mean()), 1)

    return res
