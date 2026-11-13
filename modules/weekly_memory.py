"""
==============================================================================
MODULE DE MÉMOIRE SÉLECTIVE & RAPPORTS HEBDOMADAIRES — RUGBY CLUB AMIÉNOIS (RCA)
==============================================================================
Module 100% indépendant gérant :
1. L'archivage manuel et ciblé de séances validées par semaine ISO.
2. Le stockage persistant des séances dans data/weekly_memory/.
3. La fusion et l'agrégation statistique multi-séances par joueur :
   - Somme pour le volume (Distance totale, Haute intensité, VHSR/Sprint, Accel/Decel, Temps)
   - Maximum pour les pics (Top speed km/h)
   - Moyenne / ratio pour les intensités relatives (m/min, % marche)
   - Comptage des séances participées
4. Le comparatif de charge journalière (Mardi vs Mercredi vs Vendredi vs Dimanche).
5. L'export PDF multi-pages officiel du Rapport Hebdomadaire pour les coachs.
==============================================================================
"""

from __future__ import annotations

import datetime
import io
import json
import os
import pathlib
import re
from typing import Any, Dict, List, Optional, Tuple, Union
import zipfile

from fpdf import FPDF
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import numpy as np
import pandas as pd

try:
    from modules import photo_manager
except ImportError:
    import photo_manager

try:
    from modules import rpe_connector
except ImportError:
    try:
        import rpe_connector
    except Exception:
        rpe_connector = None

try:
    from modules import gsheet_sync
except ImportError:
    try:
        import gsheet_sync
    except Exception:
        gsheet_sync = None

PROJECT_ROOT = pathlib.Path(__file__).parent.parent.resolve()
DATA_DIR = PROJECT_ROOT / "data"
WEEKLY_MEMORY_DIR = str(DATA_DIR / "weekly_memory")
WEEKLY_PDFS_DIR = str(DATA_DIR / "weekly_pdfs")
TARGETS_CONFIG_FILE = str(DATA_DIR / "targets_config.json")
os.makedirs(WEEKLY_MEMORY_DIR, exist_ok=True)
os.makedirs(WEEKLY_PDFS_DIR, exist_ok=True)


def format_readable_date(val: Any) -> str:
    """
    Convertit n'importe quelle valeur de date (numéro de série Excel ex: 46315, 46316, 46318,
    timestamp, chaîne ISO, ou datetime) en date lisible au format français JJ/MM/AAAA.
    """
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
        return val.strftime("%d/%m/%Y")
    
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ["nan", "none", "nat"]:
        return ""

    # Test si c'est un numéro de série Excel (ex: 46315 ou "46315.0")
    try:
        val_clean_num = val_str.replace(",", ".")
        val_float = float(val_clean_num)
        if 30000 <= val_float <= 80000:
            dt = pd.to_datetime(val_float, unit="D", origin="1899-12-30")
            return dt.strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        pass

    # Si c'est déjà JJ/MM/AAAA
    if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", val_str):
        parts = val_str.split("/")
        return f"{int(parts[0]):02d}/{int(parts[1]):02d}/{parts[2]}"

    # Si format ISO YYYY-MM-DD
    if re.match(r"^\d{4}-\d{2}-\d{2}", val_str):
        try:
            dt = pd.to_datetime(val_str)
            return dt.strftime("%d/%m/%Y")
        except Exception:
            pass

    try:
        dt = pd.to_datetime(val_str, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass

    return val_str


JOURS_SEANCE_STANDARDS = [
    "Mardi (Entraînement)",
    "Mercredi (Entraînement)",
    "Vendredi (Entraînement)",
    "Dimanche (Match)",
    "Lundi (Récup / Séance)",
    "Jeudi (Entraînement)",
    "Samedi (Match / Mise en place)",
    "Autre séance",
]


def get_current_iso_week() -> Tuple[int, int]:
    """Retourne l'année et le numéro de semaine ISO actuels (ex: (2026, 43))."""
    today = datetime.date.today()
    iso = today.isocalendar()
    return int(iso[0]), int(iso[1])


def get_week_file_path(year: int, week_num: int) -> str:
    """Retourne le chemin absolu du fichier JSON pour une semaine donnée."""
    os.makedirs(WEEKLY_MEMORY_DIR, exist_ok=True)
    return os.path.join(WEEKLY_MEMORY_DIR, f"week_{year}_W{week_num:02d}.json")


def list_saved_weeks() -> List[Dict[str, Any]]:
    """
    Scanne data/weekly_memory/ et retourne la liste triée des semaines archivées.
    Chaque élément : { 'key': '2026-W43', 'year': 2026, 'week_num': 43, 'label': 'Semaine 43 - 2026', 'nb_sessions': 3 }
    """
    os.makedirs(WEEKLY_MEMORY_DIR, exist_ok=True)
    res = []
    pattern = re.compile(r"^week_(\d{4})_W(\d{2})\.json$")
    for f in os.listdir(WEEKLY_MEMORY_DIR):
        m = pattern.match(f)
        if m:
            y = int(m.group(1))
            w = int(m.group(2))
            fp = os.path.join(WEEKLY_MEMORY_DIR, f)
            try:
                with open(fp, "r", encoding="utf-8") as jf:
                    data = json.load(jf)
                    sessions = data.get("sessions", {})
                    nb_s = len(sessions)
            except Exception:
                nb_s = 0
            res.append({
                "key": f"{y}-W{w:02d}",
                "year": y,
                "week_num": w,
                "label": f"Semaine {w:02d} — {y} ({nb_s} séance{'s' if nb_s > 1 else ''})",
                "nb_sessions": nb_s,
                "file_path": fp,
            })
    res.sort(key=lambda x: (x["year"], x["week_num"]), reverse=True)
    return res


def load_week_data(year: int, week_num: int) -> Dict[str, Any]:
    """Charge les données d'une semaine enregistrée."""
    fp = get_week_file_path(year, week_num)
    if not os.path.exists(fp):
        return {
            "key": f"{year}-W{week_num:02d}",
            "year": year,
            "week_num": week_num,
            "sessions": {},
        }
    try:
        with open(fp, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {
            "key": f"{year}-W{week_num:02d}",
            "year": year,
            "week_num": week_num,
            "sessions": {},
        }


def get_session_pdf_path(year: int, week_num: int, session_id: str) -> str:
    """Retourne le chemin absolu du fichier PDF archivé pour une séance spécifique."""
    os.makedirs(WEEKLY_PDFS_DIR, exist_ok=True)
    return os.path.join(WEEKLY_PDFS_DIR, f"week_{year}_W{week_num:02d}_{session_id}.pdf")


def get_session_pdf_bytes(
    year: int,
    week_num: int,
    session_id: str,
    generate_if_missing: bool = True,
) -> Optional[bytes]:
    """
    Récupère les octets (bytes) du PDF archivé de la séance.
    Si le fichier physique n'est pas encore présent sur disque mais que la séance existe en mémoire,
    il régénère le PDF à la volée à partir des données de joueurs sauvegardées.
    """
    pdf_path = get_session_pdf_path(year, week_num, session_id)
    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
        try:
            with open(pdf_path, "rb") as f:
                return f.read()
        except Exception:
            pass

    if not generate_if_missing:
        return None

    # Régénération dynamique à la volée
    week_data = load_week_data(year, week_num)
    sessions = week_data.get("sessions", {})
    if session_id not in sessions:
        return None

    s_info = sessions[session_id]
    p_data = s_info.get("players_data", [])
    if not p_data:
        return None

    df_indiv = pd.DataFrame(p_data)
    date_val = format_readable_date(s_info.get("date", ""))
    day_val = s_info.get("day_type", "")

    try:
        import app
        pdf_bytes = app.generer_rapport_pdf_coachs(
            df_indiv=df_indiv,
            date_seance=date_val or day_val,
            synthese_note=s_info.get("notes", ""),
            moyennes_groupes=s_info.get("group_stats"),
        )
        if pdf_bytes and len(pdf_bytes) > 1000:
            try:
                os.makedirs(WEEKLY_PDFS_DIR, exist_ok=True)
                with open(pdf_path, "wb") as f_out:
                    f_out.write(pdf_bytes)
                s_info["has_pdf"] = True
                s_info["pdf_filename"] = os.path.basename(pdf_path)
                fp = get_week_file_path(year, week_num)
                with open(fp, "w", encoding="utf-8") as f_w:
                    json.dump(week_data, f_w, ensure_ascii=False, indent=2)
            except Exception:
                pass
            return pdf_bytes
    except Exception as e:
        print(f"Erreur lors de la régénération du PDF journalier pour {session_id} : {e}")

    return None


def save_session_to_memory(
    df_indiv: pd.DataFrame,
    year: int,
    week_num: int,
    day_type: str,
    date_str: Optional[str] = None,
    notes: str = "",
    overwrite: bool = True,
    pdf_bytes: Optional[bytes] = None,
    group_stats: Optional[Dict[str, Any]] = None,
    adversaire: str = "",
    rpe_stats: Optional[Dict[str, Any]] = None,
) -> Tuple[bool, str]:
    """
    Sauvegarde manuellement une séance validée dans la mémoire hebdomadaire.
    Archivage simultané des données structurées et du fichier PDF journalier dans data/weekly_pdfs/.
    """
    if df_indiv is None or df_indiv.empty:
        return False, "Le tableau des données joueurs est vide."

    clean_day = day_type.strip()
    session_id = clean_day.lower().replace(" ", "_").replace("(", "").replace(")", "").replace("'", "")
    clean_date = format_readable_date(date_str) if date_str else datetime.date.today().strftime("%d/%m/%Y")
    if not clean_date:
        clean_date = datetime.date.today().strftime("%d/%m/%Y")

    week_data = load_week_data(year, week_num)
    existing_sessions = week_data.setdefault("sessions", {})

    if session_id in existing_sessions and not overwrite:
        return False, f"Une séance pour « {clean_day} » existe déjà pour la semaine {week_num}."

    # Sérialisation propre des enregistrements joueurs
    serialized_players: List[Dict[str, Any]] = []
    for _, row in df_indiv.iterrows():
        p_dict: Dict[str, Any] = {}
        for col in df_indiv.columns:
            val = row[col]
            if pd.isna(val) or val is None:
                p_dict[col] = None
            elif isinstance(val, (np.floating, float)):
                p_dict[col] = round(float(val), 2)
            elif isinstance(val, (np.integer, int)):
                p_dict[col] = int(val)
            else:
                p_dict[col] = str(val).strip()
        serialized_players.append(p_dict)

    # Métadonnées calculées de la séance
    total_dist_km = 0.0
    if "Distance (km)" in df_indiv.columns:
        total_dist_km = float(pd.to_numeric(df_indiv["Distance (km)"], errors="coerce").fillna(0).sum())

    total_hi_m = 0.0
    hi_col = "Haute intensité Podium (m)" if "Haute intensité Podium (m)" in df_indiv.columns else " Distance haute intensité (m)"
    if hi_col in df_indiv.columns:
        total_hi_m = float(pd.to_numeric(df_indiv[hi_col], errors="coerce").fillna(0).sum())

    avg_dpm = 0.0
    if "Distance Per Min (m/min)" in df_indiv.columns:
        avg_dpm = float(pd.to_numeric(df_indiv["Distance Per Min (m/min)"], errors="coerce").fillna(0).mean())

    # Archivage automatique du fichier PDF si fourni
    pdf_filename = f"week_{year}_W{week_num:02d}_{session_id}.pdf"
    pdf_path = get_session_pdf_path(year, week_num, session_id)
    has_pdf = False
    if pdf_bytes and len(pdf_bytes) > 1000:
        try:
            os.makedirs(WEEKLY_PDFS_DIR, exist_ok=True)
            with open(pdf_path, "wb") as f_pdf:
                f_pdf.write(pdf_bytes)
            has_pdf = True
        except Exception as e:
            print(f"Avertissement : Impossible d'écrire le PDF journalier sur disque : {e}")
    elif os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
        has_pdf = True

    adv_clean = str(adversaire or "").strip().upper()

    existing_sessions[session_id] = {
        "session_id": session_id,
        "day_type": clean_day,
        "date": clean_date,
        "recorded_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "notes": notes,
        "adversaire": adv_clean,
        "num_players": len(df_indiv),
        "total_distance_km": round(total_dist_km, 2),
        "total_hi_m": round(total_hi_m, 1),
        "avg_dpm": round(avg_dpm, 1),
        "has_pdf": has_pdf,
        "pdf_filename": pdf_filename,
        "group_stats": group_stats or {},
        "rpe_stats": rpe_stats or {},
        "players_data": serialized_players,
    }

    week_data["last_updated"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    fp = get_week_file_path(year, week_num)
    try:
        with open(fp, "w", encoding="utf-8") as f:
            json.dump(week_data, f, ensure_ascii=False, indent=2)

        try:
            try:

                import github_sync

            except (ImportError, ModuleNotFoundError):

                from modules import github_sync
            github_sync.push_file_to_github(
                fp,
                commit_message=f"data: enregistrement séance {clean_day} (Semaine {week_num:02d} — {year})"
            )
        except Exception:
            pass

        return True, f"Séance « {clean_day} » archivée avec succès en Semaine {week_num} ({len(df_indiv)} joueurs)."
    except Exception as e:
        return False, f"Erreur lors de la sauvegarde : {str(e)}"


def delete_session_from_memory(year: int, week_num: int, session_id: str) -> Tuple[bool, str]:
    """Supprime une séance spécifique d'une semaine archivée ainsi que son PDF associé."""
    week_data = load_week_data(year, week_num)
    sessions = week_data.get("sessions", {})
    if session_id in sessions:
        day_label = sessions[session_id].get("day_type", session_id)
        del sessions[session_id]

        # Suppression du fichier PDF archivé s'il existe
        pdf_path = get_session_pdf_path(year, week_num, session_id)
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except Exception:
                pass

        fp = get_week_file_path(year, week_num)
        try:
            if not sessions:
                if os.path.exists(fp):
                    os.remove(fp)
                try:
                    try:

                        import github_sync

                    except (ImportError, ModuleNotFoundError):

                        from modules import github_sync
                    github_sync.delete_file_from_github(
                        fp,
                        commit_message=f"data: suppression de la semaine {week_num:02d} — {year} (vide)"
                    )
                except Exception:
                    pass
                return True, f"Séance « {day_label} » supprimée. Semaine {week_num} vide nettoyée."
            else:
                week_data["last_updated"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                with open(fp, "w", encoding="utf-8") as f:
                    json.dump(week_data, f, ensure_ascii=False, indent=2)
                try:
                    try:

                        import github_sync

                    except (ImportError, ModuleNotFoundError):

                        from modules import github_sync
                    github_sync.push_file_to_github(
                        fp,
                        commit_message=f"data: suppression séance {day_label} (Semaine {week_num:02d} — {year})"
                    )
                except Exception:
                    pass
                return True, f"Séance « {day_label} » supprimée de la Semaine {week_num}."
        except Exception as e:
            return False, f"Erreur lors de la suppression : {str(e)}"
    return False, "Séance introuvable dans cette semaine."


def delete_player_records(player_name: str) -> int:
    """
    Supprime toutes les traces d'un joueur personnalisé dans la mémoire hebdomadaire
    lorsque son profil est supprimé de l'application (respect de la règle sans orphelin).
    """
    p_clean = player_name.strip().lower()
    total_cleaned = 0
    weeks = list_saved_weeks()
    for w_info in weeks:
        fp = w_info["file_path"]
        try:
            with open(fp, "r", encoding="utf-8") as f:
                data = json.load(f)
            changed = False
            for s_id, s_data in data.get("sessions", {}).items():
                p_list = s_data.get("players_data", [])
                new_list = [p for p in p_list if str(p.get("Nom ; Prénom", "")).strip().lower() != p_clean]
                if len(new_list) != len(p_list):
                    s_data["players_data"] = new_list
                    s_data["num_players"] = len(new_list)
                    changed = True
                    total_cleaned += 1
            if changed:
                with open(fp, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                try:
                    try:

                        import github_sync

                    except (ImportError, ModuleNotFoundError):

                        from modules import github_sync
                    github_sync.push_file_to_github(
                        fp,
                        commit_message=f"data: nettoyage données joueur {player_name} (Semaine {w_info.get('week_num', '')})"
                    )
                except Exception:
                    pass
        except Exception:
            pass
    return total_cleaned


# ==============================================================================
# MOTEUR D'AGRÉGATION HEBDOMADAIRE MULTI-SÉANCES
# ==============================================================================

def aggregate_weekly_sessions(
    week_data: Dict[str, Any],
    selected_session_ids: Optional[List[str]] = None,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Agrège les séances sélectionnées sur la semaine selon les règles athlétiques du rugby :
    - Somme (Cumul) : Distance totale, Haute intensité, VHSR/Sprint, Accélérations/Décélérations, Temps de jeu.
    - Maximum (Pic) : Top Speed (km/h).
    - Moyenne pondérée / ratio : Distance/min (m/min), % de marche.
    - Métrique de présence : Nombre de séances participées.
    Retourne (df_hebdo, synthese_jours).
    """
    sessions = week_data.get("sessions", {})
    if not sessions:
        return pd.DataFrame(), {}

    if selected_session_ids is None:
        target_sessions = list(sessions.values())
    else:
        target_sessions = [s for s_id, s in sessions.items() if s_id in selected_session_ids]

    if not target_sessions:
        return pd.DataFrame(), {}

    # 1. Comparatif journalier de la charge
    synthese_jours: Dict[str, Dict[str, Any]] = {}
    for s in target_sessions:
        s_id = s.get("session_id", "")
        synthese_jours[s_id] = {
            "day_type": s.get("day_type", ""),
            "date": s.get("date", ""),
            "num_players": s.get("num_players", 0),
            "total_distance_km": s.get("total_distance_km", 0.0),
            "total_hi_m": s.get("total_hi_m", 0.0),
            "avg_dpm": s.get("avg_dpm", 0.0),
            "adversaire": s.get("adversaire", ""),
            "rpe_stats": s.get("rpe_stats", {}),
            "players_data": s.get("players_data", []),
        }

    # 2. Consolidation individuelle par joueur
    all_player_rows: List[Dict[str, Any]] = []
    for s in target_sessions:
        day_name = s.get("day_type", "")
        p_list = s.get("players_data", [])
        for p in p_list:
            row_copy = dict(p)
            row_copy["_session_day"] = day_name
            all_player_rows.append(row_copy)

    if not all_player_rows:
        return pd.DataFrame(), synthese_jours

    df_all = pd.DataFrame(all_player_rows)
    p_col = "Nom ; Prénom"

    # Conversion numérique robuste
    numeric_cols = [
        "Temps de jeu ", "Distance (km)", "Distance_m",
        "Haute intensité Podium (m)", " Distance haute intensité (m)",
        "HSR Z3 (m)", "HSR_m", "VHSR (m)", "Top Speed (km/h)",
        "Distance Per Min (m/min)", "%Marche ", "%course", "%HSR ", "%VHSR",
        "Accel/Dec", "Accelerations Zone Count: 3 - 4 m/s/s", "Accelerations Zone Count: > 4 m/s/s",
        "Deceleration Zone Count: 3 - 4 m/s/s", "Deceleration Zone Count: > 4 m/s/s",
    ]
    for c in numeric_cols:
        if c in df_all.columns:
            df_all[c] = pd.to_numeric(df_all[c], errors="coerce").fillna(0.0)

    # Colonne standardisée pour la distance en mètres et haute intensité
    if "Distance_m" not in df_all.columns and "Distance (km)" in df_all.columns:
        df_all["Distance_m"] = df_all["Distance (km)"] * 1000.0
    if "Distance (km)" not in df_all.columns and "Distance_m" in df_all.columns:
        df_all["Distance (km)"] = df_all["Distance_m"] / 1000.0

    hi_col_target = "Haute intensité Podium (m)"
    if hi_col_target not in df_all.columns and " Distance haute intensité (m)" in df_all.columns:
        df_all[hi_col_target] = df_all[" Distance haute intensité (m)"]

    # Normalisation du nom de joueur
    if "Nom ; Prénom" not in df_all.columns:
        if "Nom" in df_all.columns and "Prénom" in df_all.columns:
            df_all["Nom ; Prénom"] = df_all["Nom"].astype(str).str.strip() + " " + df_all["Prénom"].astype(str).str.strip()
        elif "Nom" in df_all.columns:
            df_all["Nom ; Prénom"] = df_all["Nom"].astype(str).str.strip()
        elif "Player Name" in df_all.columns:
            df_all["Nom ; Prénom"] = df_all["Player Name"].astype(str).str.strip()
        elif "Athlete Name" in df_all.columns:
            df_all["Nom ; Prénom"] = df_all["Athlete Name"].astype(str).str.strip()
        elif "Joueur" in df_all.columns:
            df_all["Nom ; Prénom"] = df_all["Joueur"].astype(str).str.strip()
        else:
            df_all["Nom ; Prénom"] = "Joueur Inconnu"

    grouped = df_all.groupby(p_col)

    agg_records = []
    try:
        from modules.gsheet_sync import is_fictional_test_player
    except Exception:
        is_fictional_test_player = None

    for name, group in grouped:
        if is_fictional_test_player and is_fictional_test_player(str(name)):
            continue
        nb_seances = len(group)
        jours_presence = ", ".join(group["_session_day"].unique().tolist())

        # Groupe et poste de référence (normalisés avec lookup effectif si l'archive ne contenait que le groupe)
        raw_poste_det = group["Poste_Detaille"].iloc[-1] if "Poste_Detaille" in group.columns else ""
        poste_det = normaliser_poste_specifique(raw_poste_det, player_name=str(name)) or "Inconnu"
        groupe_rca = POSTE_TO_GROUPE_MAP.get(poste_det) or (group["Groupe"].iloc[-1] if "Groupe" in group.columns else "Trois-quarts")

        # CUMULS (SOMMES HEBDOMADAIRES)
        tot_time = float(group["Temps de jeu "].sum()) if "Temps de jeu " in group.columns else 0.0
        tot_dist_km = float(group["Distance (km)"].sum()) if "Distance (km)" in group.columns else 0.0
        tot_dist_m = tot_dist_km * 1000.0

        tot_hi_m = float(group[hi_col_target].sum()) if hi_col_target in group.columns else 0.0
        tot_hsr_z3 = float(group["HSR Z3 (m)"].sum()) if "HSR Z3 (m)" in group.columns else (float(group["HSR_m"].sum()) if "HSR_m" in group.columns else 0.0)
        tot_vhsr_m = float(group["VHSR (m)"].sum()) if "VHSR (m)" in group.columns else 0.0
        tot_acc_dec = float(group["Accel/Dec"].sum()) if "Accel/Dec" in group.columns else 0.0

        tot_acc_34 = float(group["Accelerations Zone Count: 3 - 4 m/s/s"].sum()) if "Accelerations Zone Count: 3 - 4 m/s/s" in group.columns else 0.0
        tot_acc_gt4 = float(group["Accelerations Zone Count: > 4 m/s/s"].sum()) if "Accelerations Zone Count: > 4 m/s/s" in group.columns else 0.0
        tot_dec_34 = float(group["Deceleration Zone Count: 3 - 4 m/s/s"].sum()) if "Deceleration Zone Count: 3 - 4 m/s/s" in group.columns else 0.0
        tot_dec_gt4 = float(group["Deceleration Zone Count: > 4 m/s/s"].sum()) if "Deceleration Zone Count: > 4 m/s/s" in group.columns else 0.0

        # PICS (MAXIMUMS HEBDOMADAIRES)
        top_speed = float(group["Top Speed (km/h)"].max()) if "Top Speed (km/h)" in group.columns else 0.0

        # INTENSITÉS RELATIVES & MOYENNES
        dpm = (tot_dist_m / tot_time) if tot_time > 0 else 0.0
        hsr_pm = (tot_hi_m / tot_time) if tot_time > 0 else 0.0
        pct_marche = float(group["%Marche "].mean()) if "%Marche " in group.columns else 0.0

        # RPE INDIVIDUEL (MOYENNE) & SESSION-RPE / FOSTER (CUMUL HEBDO)
        rpe_mean = float(pd.to_numeric(group["RPE"], errors="coerce").dropna().mean()) if "RPE" in group.columns and pd.to_numeric(group["RPE"], errors="coerce").notna().any() else np.nan
        foster_tot = float(pd.to_numeric(group["Session_RPE"], errors="coerce").dropna().sum()) if "Session_RPE" in group.columns and pd.to_numeric(group["Session_RPE"], errors="coerce").notna().any() else np.nan

        agg_records.append({
            "Nom ; Prénom": name,
            "Poste_Detaille": poste_det,
            "Poste": poste_det,
            "Groupe": groupe_rca,
            "Nombre de séances": nb_seances,
            "Séances": jours_presence,
            "Temps total (min)": round(tot_time, 1),
            "Temps de jeu ": round(tot_time, 1),
            "Distance Totale (km)": round(tot_dist_km, 2),
            "Distance (km)": round(tot_dist_km, 2),
            "Distance Totale (m)": round(tot_dist_m, 1),
            "Distance_m": round(tot_dist_m, 1),
            "Haute intensité (m)": round(tot_hi_m, 0),
            "Haute intensité Podium (m)": round(tot_hi_m, 0),
            " Distance haute intensité (m)": round(tot_hi_m, 0),
            "HSR Z3 (m)": round(tot_hsr_z3, 0),
            "Distance in Speed Zone 3  (km)": round(tot_hsr_z3 / 1000.0, 3),
            "Distance Sprint VHSR (m)": round(tot_vhsr_m, 0),
            "VHSR (m)": round(tot_vhsr_m, 0),
            "Top Speed (km/h)": round(top_speed, 1),
            "Distance/min (m/min)": round(dpm, 1),
            "Distance Per Min (m/min)": round(dpm, 1),
            "HSR m/min": round(hsr_pm, 1),
            "Accel/Dec total": int(round(tot_acc_dec)),
            "Accel/Dec": round(tot_acc_dec, 0),
            "Accelerations Zone Count: 3 - 4 m/s/s": int(round(tot_acc_34)),
            "Accelerations Zone Count: > 4 m/s/s": int(round(tot_acc_gt4)),
            "Deceleration Zone Count: 3 - 4 m/s/s": int(round(tot_dec_34)),
            "Deceleration Zone Count: > 4 m/s/s": int(round(tot_dec_gt4)),
            "Accélérations >3m/s²": int(round(tot_acc_34 + tot_acc_gt4)),
            "Décélérations >3m/s²": int(round(tot_dec_34 + tot_dec_gt4)),
            "% de marche moyen": round(pct_marche, 1),
            "%Marche ": round(pct_marche, 1),
            "RPE": round(rpe_mean, 1) if pd.notna(rpe_mean) else np.nan,
            "Session_RPE": round(foster_tot, 1) if pd.notna(foster_tot) else np.nan,
        })

    df_hebdo = pd.DataFrame(agg_records)
    df_hebdo.sort_values(by="Distance Totale (km)", ascending=False, inplace=True)
    return df_hebdo, synthese_jours


# ==============================================================================
# GÉNÉRATION DES PODIUMS HEBDOMADAIRES (AVEC PHOTOS & HTML NETTOYÉ)
# ==============================================================================

def _compute_top_and_vigilance_players_list(
    df_src: pd.DataFrame,
    metric_col: Optional[str],
    ascending_best: bool,
    fmt_fn: Any,
    require_positive: bool = False,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], int]:
    """
    Moteur unique (3 Premiers & 3 Derniers Vigilance) partagé entre les rapports Journalier et Hebdomadaire.
    - ascending_best=False -> 3 Premiers = plus grandes valeurs, 3 Derniers = plus petites valeurs.
    - ascending_best=True  -> 3 Premiers = plus petites valeurs (>0 si require_positive), 3 Derniers = plus grandes valeurs.
    - Si effectif valide < 6 : évite les doublons avec les 3 Premiers (affiche les joueurs restants
      si 4 ou 5 joueurs, et liste vide si <= 3 joueurs).
    """
    def _make_entry(rank: int, name: Any, stat_str: Any, poste: Any = "") -> Dict[str, Any]:
        p_name = str(name).strip() if pd.notna(name) and name is not None else "-"
        if p_name in ["nan", "None", ""]:
            p_name = "-"
        p_poste = normaliser_poste_specifique(poste, player_name=p_name) if p_name != "-" else ""
        if p_poste in ["nan", "None"]:
            p_poste = ""
        return {
            "rank": rank,
            "name": p_name,
            "poste": p_poste if p_poste else "RCA",
            "stat": str(stat_str) if stat_str is not None else "-",
            "avatar_src": photo_manager.get_player_image_b64(p_name if p_name != "-" else "", size=80 if rank == 1 else 65),
        }

    def _placeholder(rank: int) -> Dict[str, Any]:
        return {
            "rank": rank,
            "name": "-",
            "poste": "",
            "stat": "-",
            "avatar_src": photo_manager.get_player_image_b64("", size=80 if rank == 1 else 65),
        }

    if df_src is None or not isinstance(df_src, pd.DataFrame) or df_src.empty or not metric_col or metric_col not in df_src.columns:
        return [_placeholder(1), _placeholder(2), _placeholder(3)], [], 0

    df_work = df_src.copy()
    vals = pd.to_numeric(df_work[metric_col], errors="coerce")
    df_work = df_work.assign(_metric_val=vals).dropna(subset=["_metric_val"])
    if require_positive:
        df_pos = df_work[df_work["_metric_val"] > 0.0]
        if not df_pos.empty:
            df_work = df_pos
    else:
        df_pos = df_work[df_work["_metric_val"] > 0.0]
        if not df_pos.empty:
            df_work = df_pos

    n_valid = len(df_work)
    if n_valid == 0:
        return [_placeholder(1), _placeholder(2), _placeholder(3)], [], 0

    df_top_sorted = df_work.sort_values(by="_metric_val", ascending=ascending_best)
    df_top3 = df_top_sorted.head(3)

    top_players: List[Dict[str, Any]] = []
    for rank, (_, row) in enumerate(df_top3.iterrows(), start=1):
        name = str(row.get("Nom ; Prénom", row.get("Player Name", "-"))).strip()
        poste = str(row.get("Poste_Detaille", row.get("Poste", "")))
        top_players.append(_make_entry(rank, name, fmt_fn(float(row["_metric_val"])), poste))
    while len(top_players) < 3:
        top_players.append(_placeholder(len(top_players) + 1))

    if n_valid <= 3:
        vigilance_players: List[Dict[str, Any]] = []
    else:
        df_remaining = df_top_sorted.iloc[3:].copy()
        df_vig_sorted = df_remaining.sort_values(by="_metric_val", ascending=(not ascending_best))
        df_vig3 = df_vig_sorted.head(3)
        vigilance_players = []
        for rank, (_, row) in enumerate(df_vig3.iterrows(), start=1):
            name = str(row.get("Nom ; Prénom", row.get("Player Name", "-"))).strip()
            poste = str(row.get("Poste_Detaille", row.get("Poste", "")))
            vigilance_players.append(_make_entry(rank, name, fmt_fn(float(row["_metric_val"])), poste))

    return top_players, vigilance_players, n_valid


def calculer_podiums_seance(df_subset: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    """
    Calcule les Podiums (3 Premiers + 3 Derniers Vigilance) d'une séance journalière ou d'un sous-groupe :
    - Top speed (km/h) : 3 Premiers desc / 3 Derniers asc
    - % de marche : 3 Premiers asc (>0%) / 3 Derniers desc
    - Haute intensité (m) : 3 Premiers desc / 3 Derniers asc
    """
    if df_subset is None or not isinstance(df_subset, pd.DataFrame) or df_subset.empty:
        t_e, v_e, _ = _compute_top_and_vigilance_players_list(pd.DataFrame(), None, False, lambda v: "-")
        return {
            "Top speed (km/h)": {"icon": "⚡", "players": list(t_e), "vigilance_players": list(v_e), "total_valid": 0},
            "% de marche": {"icon": "🚶", "players": list(t_e), "vigilance_players": list(v_e), "total_valid": 0},
            "Haute intensité (m)": {"icon": "🔥", "players": list(t_e), "vigilance_players": list(v_e), "total_valid": 0},
        }

    ts_col = "Top Speed (km/h)" if "Top Speed (km/h)" in df_subset.columns else None
    t_ts, v_ts, n_ts = _compute_top_and_vigilance_players_list(
        df_subset, ts_col, ascending_best=False, fmt_fn=lambda v: f"{v:.1f} km/h"
    )

    walk_col = "%Marche " if "%Marche " in df_subset.columns else ("% de marche moyen" if "% de marche moyen" in df_subset.columns else None)
    def _fmt_walk(v: float) -> str:
        pct = v * 100.0 if v <= 1.0 else v
        return f"{pct:.1f} %"
    t_wk, v_wk, n_wk = _compute_top_and_vigilance_players_list(
        df_subset, walk_col, ascending_best=True, fmt_fn=_fmt_walk, require_positive=True
    )

    hi_col = None
    for c in ["Haute intensité Podium (m)", " Distance haute intensité (m)", "Haute intensité (m)"]:
        if c in df_subset.columns:
            hi_col = c
            break
    t_hi, v_hi, n_hi = _compute_top_and_vigilance_players_list(
        df_subset, hi_col, ascending_best=False, fmt_fn=lambda v: f"{int(round(v)):,d} m".replace(",", " ")
    )

    return {
        "Top speed (km/h)": {"icon": "⚡", "players": t_ts, "vigilance_players": v_ts, "total_valid": n_ts},
        "% de marche": {"icon": "🚶", "players": t_wk, "vigilance_players": v_wk, "total_valid": n_wk},
        "Haute intensité (m)": {"icon": "🔥", "players": t_hi, "vigilance_players": v_hi, "total_valid": n_hi},
    }


def generer_podiums_hebdomadaires(
    df_hebdo: pd.DataFrame,
    groupe: Optional[str] = None,
    poste_specifique: Optional[str] = None,
) -> Dict[str, Dict[str, Any]]:
    """
    Génère les podiums de la semaine (3 Premiers & 3 Derniers Vigilance) : Top Speed, % Marche, Haute Intensité.
    Supporte le filtrage par ligne de poste (groupe) ou par poste spécifique (9 postes).
    """
    filter_cat = poste_specifique or groupe
    if df_hebdo is None or not isinstance(df_hebdo, pd.DataFrame) or df_hebdo.empty:
        return calculer_podiums_seance(pd.DataFrame())

    if filter_cat and str(filter_cat).strip().upper() not in ["ALL", "EFFECTIF GLOBAL", "TOUS LES JOUEURS", "TOUS LES GROUPES"]:
        df_subset = _filter_df_by_group(df_hebdo, str(filter_cat).strip())
    else:
        df_subset = df_hebdo.copy()

    return calculer_podiums_seance(df_subset)


# ==============================================================================
# GESTION DES OBJECTIFS DE CHARGE ATHLÉTIQUE (PAR GROUPE DE POSTE)
# ==============================================================================

PREREGLAGES_OBJECTIFS: Dict[str, Dict[str, Dict[str, float]]] = {
    "Standard / Référence RCA": {
        "5 de devant": {"Objectif Distance (m)": 4500.0, "Obj. HSR": 250.0, "Obj. VHSR": 50.0, "Obj. ACC/DEC": 35.0},
        "3ème Ligne": {"Objectif Distance (m)": 5500.0, "Obj. HSR": 450.0, "Obj. VHSR": 100.0, "Obj. ACC/DEC": 45.0},
        "Trois-quarts": {"Objectif Distance (m)": 6200.0, "Obj. HSR": 700.0, "Obj. VHSR": 250.0, "Obj. ACC/DEC": 50.0},
    },
    "Mardi (Aérobie & Volume)": {
        "5 de devant": {"Objectif Distance (m)": 4800.0, "Obj. HSR": 220.0, "Obj. VHSR": 40.0, "Obj. ACC/DEC": 32.0},
        "3ème Ligne": {"Objectif Distance (m)": 5600.0, "Obj. HSR": 380.0, "Obj. VHSR": 80.0, "Obj. ACC/DEC": 40.0},
        "Trois-quarts": {"Objectif Distance (m)": 6200.0, "Obj. HSR": 550.0, "Obj. VHSR": 180.0, "Obj. ACC/DEC": 45.0},
    },
    "Mercredi (Intensité & Vitesse)": {
        "5 de devant": {"Objectif Distance (m)": 4200.0, "Obj. HSR": 280.0, "Obj. VHSR": 60.0, "Obj. ACC/DEC": 38.0},
        "3ème Ligne": {"Objectif Distance (m)": 5200.0, "Obj. HSR": 500.0, "Obj. VHSR": 120.0, "Obj. ACC/DEC": 48.0},
        "Trois-quarts": {"Objectif Distance (m)": 5800.0, "Obj. HSR": 750.0, "Obj. VHSR": 300.0, "Obj. ACC/DEC": 55.0},
    },
    "Vendredi (Mise en place / Affûtage)": {
        "5 de devant": {"Objectif Distance (m)": 3200.0, "Obj. HSR": 140.0, "Obj. VHSR": 30.0, "Obj. ACC/DEC": 22.0},
        "3ème Ligne": {"Objectif Distance (m)": 3800.0, "Obj. HSR": 250.0, "Obj. VHSR": 60.0, "Obj. ACC/DEC": 28.0},
        "Trois-quarts": {"Objectif Distance (m)": 4200.0, "Obj. HSR": 350.0, "Obj. VHSR": 120.0, "Obj. ACC/DEC": 32.0},
    },
    "Dimanche (Match)": {
        "5 de devant": {"Objectif Distance (m)": 5500.0, "Obj. HSR": 350.0, "Obj. VHSR": 80.0, "Obj. ACC/DEC": 45.0},
        "3ème Ligne": {"Objectif Distance (m)": 6800.0, "Obj. HSR": 650.0, "Obj. VHSR": 180.0, "Obj. ACC/DEC": 55.0},
        "Trois-quarts": {"Objectif Distance (m)": 7500.0, "Obj. HSR": 900.0, "Obj. VHSR": 400.0, "Obj. ACC/DEC": 60.0},
    },
}

DEFAULT_OBJECTIFS_CONFIG: Dict[str, Dict[str, float]] = {
    "5 de devant": {"Objectif Distance (m)": 4500.0, "Obj. HSR": 250.0, "Obj. VHSR": 50.0, "Obj. ACC/DEC": 35.0},
    "3ème Ligne": {"Objectif Distance (m)": 5500.0, "Obj. HSR": 450.0, "Obj. VHSR": 100.0, "Obj. ACC/DEC": 45.0},
    "Trois-quarts": {"Objectif Distance (m)": 6200.0, "Obj. HSR": 700.0, "Obj. VHSR": 250.0, "Obj. ACC/DEC": 50.0},
    "Tous les groupes": {"Objectif Distance (m)": 5400.0, "Obj. HSR": 460.0, "Obj. VHSR": 130.0, "Obj. ACC/DEC": 43.0},
}


def _calculer_moyenne_tous_groupes(cfg: Dict[str, Dict[str, float]]) -> Dict[str, float]:
    """Calcule la moyenne pondérée / équiprobable de référence pour 'Tous les groupes'."""
    keys = ["Objectif Distance (m)", "Obj. HSR", "Obj. VHSR", "Obj. ACC/DEC"]
    grps = ["5 de devant", "3ème Ligne", "Trois-quarts"]
    res = {}
    for k in keys:
        vals = [float(cfg.get(g, {}).get(k, 0.0)) for g in grps if g in cfg]
        res[k] = round(sum(vals) / max(len(vals), 1), 1)
    return res


def charger_objectifs_config() -> Dict[str, Dict[str, float]]:
    """
    Charge les objectifs configurés depuis data/targets_config.json.
    Si le fichier n'existe pas encore, il est automatiquement initialisé avec les standards RCA.
    """
    if os.path.exists(TARGETS_CONFIG_FILE):
        try:
            with open(TARGETS_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict) and "5 de devant" in data:
                    if "Tous les groupes" not in data:
                        data["Tous les groupes"] = _calculer_moyenne_tous_groupes(data)
                    return data
        except Exception as e:
            print(f"Avertissement targets_config.json : {e}")

    cfg = dict(DEFAULT_OBJECTIFS_CONFIG)
    sauvegarder_objectifs_config(cfg)
    return cfg


def sauvegarder_objectifs_config(cfg: Dict[str, Dict[str, float]]) -> bool:
    """Enregistre les objectifs configurés dans data/targets_config.json."""
    try:
        os.makedirs(os.path.dirname(TARGETS_CONFIG_FILE), exist_ok=True)
        cfg_to_save = dict(cfg)
        cfg_to_save["Tous les groupes"] = _calculer_moyenne_tous_groupes(cfg_to_save)
        with open(TARGETS_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg_to_save, f, ensure_ascii=False, indent=2)

        try:
            try:

                import github_sync

            except (ImportError, ModuleNotFoundError):

                from modules import github_sync
            github_sync.push_file_to_github(
                TARGETS_CONFIG_FILE,
                commit_message="config: mise à jour des objectifs de performance par poste"
            )
        except Exception:
            pass

        return True
    except Exception as e:
        print(f"Erreur enregistrement targets_config.json : {e}")
        return False


# ==============================================================================
# SÉCURITÉ CLOUD : SAUVEGARDE & RESTAURATION (ARCHIVE ZIP & JSON)
# ==============================================================================

def creer_archive_backup_zip() -> bytes:
    """
    Génère une archive ZIP complète de toute la mémoire hebdomadaire :
    - Tous les JSON de data/weekly_memory/
    - Tous les PDFs de data/weekly_pdfs/
    - La configuration des objectifs data/targets_config.json
    - Les mappings persistants config_rca_gps.json & config_rca_joueurs_custom.json
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        if os.path.exists(WEEKLY_MEMORY_DIR):
            for fname in os.listdir(WEEKLY_MEMORY_DIR):
                if fname.endswith(".json"):
                    fpath = os.path.join(WEEKLY_MEMORY_DIR, fname)
                    zf.write(fpath, arcname=f"weekly_memory/{fname}")

        if os.path.exists(WEEKLY_PDFS_DIR):
            for fname in os.listdir(WEEKLY_PDFS_DIR):
                if fname.endswith(".pdf"):
                    fpath = os.path.join(WEEKLY_PDFS_DIR, fname)
                    zf.write(fpath, arcname=f"weekly_pdfs/{fname}")

        if os.path.exists(TARGETS_CONFIG_FILE):
            zf.write(TARGETS_CONFIG_FILE, arcname="targets_config.json")

        for cfg_name in ["config_rca_gps.json", "config_rca_joueurs_custom.json"]:
            cp = os.path.join(str(PROJECT_ROOT), cfg_name)
            if os.path.exists(cp):
                zf.write(cp, arcname=cfg_name)

    buf.seek(0)
    return buf.getvalue()


def restaurer_archive_backup(uploaded_file: Any) -> Tuple[bool, str]:
    """
    Restaure une sauvegarde à partir d'un fichier ZIP ou JSON uploadé par l'utilisateur.
    Garantit une persistance absolue en cas de réinitialisation d'un conteneur Streamlit Cloud.
    """
    if uploaded_file is None:
        return False, "Aucun fichier fourni."

    filename = getattr(uploaded_file, "name", "").lower()
    file_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()

    if not file_bytes:
        return False, "Le fichier de sauvegarde est vide."

    os.makedirs(WEEKLY_MEMORY_DIR, exist_ok=True)
    os.makedirs(WEEKLY_PDFS_DIR, exist_ok=True)

    # 1. Format ZIP
    if filename.endswith(".zip") or zipfile.is_zipfile(io.BytesIO(file_bytes)):
        try:
            nb_json = 0
            nb_pdf = 0
            restored_paths: List[str] = []
            with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as zf:
                for member in zf.infolist():
                    m_name = member.filename.replace("\\", "/")
                    if m_name.endswith("/") or ".." in m_name:
                        continue
                    base = os.path.basename(m_name)

                    if m_name.startswith("weekly_memory/") or (base.startswith("week_") and base.endswith(".json")):
                        target_path = os.path.join(WEEKLY_MEMORY_DIR, base)
                        with open(target_path, "wb") as f_out:
                            f_out.write(zf.read(member))
                        restored_paths.append(target_path)
                        nb_json += 1
                    elif m_name.startswith("weekly_pdfs/") or base.endswith(".pdf"):
                        target_path = os.path.join(WEEKLY_PDFS_DIR, base)
                        with open(target_path, "wb") as f_out:
                            f_out.write(zf.read(member))
                        nb_pdf += 1
                    elif base == "targets_config.json":
                        with open(TARGETS_CONFIG_FILE, "wb") as f_out:
                            f_out.write(zf.read(member))
                        restored_paths.append(TARGETS_CONFIG_FILE)
                    elif base in ["config_rca_gps.json", "config_rca_joueurs_custom.json"]:
                        target_path = os.path.join(str(PROJECT_ROOT), base)
                        with open(target_path, "wb") as f_out:
                            f_out.write(zf.read(member))
                        restored_paths.append(target_path)

            try:
                try:

                    import github_sync

                except (ImportError, ModuleNotFoundError):

                    from modules import github_sync
                for rp in restored_paths:
                    github_sync.push_file_to_github(
                        rp,
                        commit_message=f"backup: restauration {os.path.basename(rp)}"
                    )
            except Exception:
                pass

            return True, f"Restauration réussie ! {nb_json} semaine(s) et {nb_pdf} PDF(s) restaurés."
        except Exception as e:
            return False, f"Erreur lors de la décompression : {str(e)}"

    # 2. Format JSON
    elif filename.endswith(".json"):
        try:
            raw_text = file_bytes.decode("utf-8")
            data = json.loads(raw_text)
            restored_path = None
            if isinstance(data, dict):
                if "sessions" in data:
                    y = data.get("year", 2026)
                    w = data.get("week_num", 1)
                    target_path = get_week_file_path(y, w)
                    with open(target_path, "w", encoding="utf-8") as f_out:
                        json.dump(data, f_out, ensure_ascii=False, indent=2)
                    restored_path = target_path
                    res_msg = f"Semaine {w} ({y}) restaurée avec succès ({len(data.get('sessions', {}))} séances)."
                elif "5 de devant" in data and "Trois-quarts" in data:
                    with open(TARGETS_CONFIG_FILE, "w", encoding="utf-8") as f_out:
                        json.dump(data, f_out, ensure_ascii=False, indent=2)
                    restored_path = TARGETS_CONFIG_FILE
                    res_msg = "Configuration des objectifs restaurée avec succès."
                else:
                    target_path = os.path.join(WEEKLY_MEMORY_DIR, os.path.basename(filename))
                    with open(target_path, "w", encoding="utf-8") as f_out:
                        json.dump(data, f_out, ensure_ascii=False, indent=2)
                    restored_path = target_path
                    res_msg = f"Fichier {filename} importé dans la mémoire hebdomadaire."

                if restored_path:
                    try:
                        try:

                            import github_sync

                        except (ImportError, ModuleNotFoundError):

                            from modules import github_sync
                        github_sync.push_file_to_github(
                            restored_path,
                            commit_message=f"backup: restauration {os.path.basename(restored_path)}"
                        )
                    except Exception:
                        pass

                return True, res_msg
        except Exception as e:
            return False, f"Erreur lecture JSON : {str(e)}"

    return False, "Format non reconnu (.zip ou .json attendu)."


# ==============================================================================
# GRAPHIQUES MATPLOTLIB POUR LE RAPPORT HEBDOMADAIRE (PDF & STREAMLIT)
# ==============================================================================


def generer_graphique_hebdo_groupes_matplotlib(
    df_agg: pd.DataFrame,
    figsize: Tuple[float, float] = (6.2, 3.6),
) -> io.BytesIO:
    """
    Génère le diagramme combiné à barres comparatif de la semaine pour la Page 1 du PDF et le Dashboard :
    - Distance Moyenne (m) par Groupe (barres bleu marine #1F4E79)
    - HSR Z3 Moyen (m) par Groupe (barres rouge #A30000)
    sur les 3 groupes officiels du RCA : 5 de devant, 3ème Ligne, Trois-quarts.
    """
    is_wide = figsize[0] >= 10.0
    fig, ax1 = plt.subplots(figsize=figsize, dpi=140)
    fig.patch.set_facecolor("#FFFFFF")
    ax1.set_facecolor("#FFFFFF")

    groupes = ["5 de devant", "3ème Ligne", "Trois-quarts"]
    dist_vals: List[float] = []
    hsr_vals: List[float] = []

    for g in groupes:
        sub = df_agg[df_agg["Groupe"] == g] if not df_agg.empty and "Groupe" in df_agg.columns else pd.DataFrame()
        if not sub.empty:
            d_val = float(sub["Distance Totale (m)"].mean()) if "Distance Totale (m)" in sub.columns else (float(sub["Distance Totale (km)"].mean()) * 1000.0 if "Distance Totale (km)" in sub.columns else 0.0)
            h_val = float(sub["HSR Z3 (m)"].mean()) if "HSR Z3 (m)" in sub.columns else (float(sub["Haute intensité (m)"].mean()) if "Haute intensité (m)" in sub.columns else 0.0)
        else:
            d_val = 0.0
            h_val = 0.0
        dist_vals.append(d_val)
        hsr_vals.append(h_val)

    x = np.arange(len(groupes))
    width = 0.30 if is_wide else 0.35

    rects1 = ax1.bar(x - width/2, dist_vals, width, label="Distance Moyenne (m)", color="#1F4E79", edgecolor="none", zorder=3)
    ax2 = ax1.twinx()
    rects2 = ax2.bar(x + width/2, hsr_vals, width, label="HSR Z3 Moyen (m)", color="#A30000", edgecolor="none", zorder=3)

    fs_lbl = 10.0 if is_wide else 8.5
    fs_tick = 10.5 if is_wide else 9.0
    fs_title = 12.0 if is_wide else 10.0
    fs_ann = 9.0 if is_wide else 7.5
    fs_leg = 9.5 if is_wide else 8.0

    ax1.set_ylabel("Distance Moyenne (m)", color="#1F4E79", fontweight="bold", fontsize=fs_lbl)
    ax2.set_ylabel("HSR Z3 Moyen (m)", color="#A30000", fontweight="bold", fontsize=fs_lbl)
    ax1.set_xticks(x)
    ax1.set_xticklabels(groupes, fontweight="bold", fontsize=fs_tick, color="#1E293B")
    ax1.set_title("Charge Moyenne Hebdomadaire par Ligne de Poste", fontweight="bold", color="#1F4E79", fontsize=fs_title, pad=10 if is_wide else 8)

    # Annotations au-dessus des barres
    for rect in rects1:
        h = rect.get_height()
        if h > 0:
            ax1.annotate(f"{int(round(h)):,d} m".replace(",", " "),
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 3 if is_wide else 2), textcoords="offset points",
                         ha="center", va="bottom", fontsize=fs_ann, fontweight="bold", color="#1F4E79")

    for rect in rects2:
        h = rect.get_height()
        if h > 0:
            ax2.annotate(f"{int(round(h)):,d} m".replace(",", " "),
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 3 if is_wide else 2), textcoords="offset points",
                         ha="center", va="bottom", fontsize=fs_ann, fontweight="bold", color="#A30000")

    max_dist = max(dist_vals) if dist_vals and max(dist_vals) > 0 else 1000.0
    max_hsr = max(hsr_vals) if hsr_vals and max(hsr_vals) > 0 else 100.0
    ax1.set_ylim(0, max_dist * 1.25)
    ax2.set_ylim(0, max_hsr * 1.25)
    ax1.grid(axis="y", linestyle="--", alpha=0.3, zorder=1)
    ax1.tick_params(axis="y", labelcolor="#1F4E79", labelsize=9.0 if is_wide else 8.0)
    ax2.tick_params(axis="y", labelcolor="#A30000", labelsize=9.0 if is_wide else 8.0)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper center", bbox_to_anchor=(0.5, -0.10 if is_wide else -0.12), ncol=2, frameon=False, fontsize=fs_leg)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def generer_graphique_hebdo_repartition_jours_matplotlib(synthese_jours: Dict[str, Any]) -> io.BytesIO:
    """
    Génère le graphique comparatif de la répartition de la charge par séance (Mardi vs Mercredi vs Vendredi...)
    pour la Page 1 du PDF et le Dashboard.
    """
    fig, ax1 = plt.subplots(figsize=(6.2, 2.6), dpi=180)
    fig.patch.set_facecolor("#FFFFFF")
    ax1.set_facecolor("#FFFFFF")

    days_labels: List[str] = []
    dist_vals: List[float] = []
    hi_vals: List[float] = []

    for s_id, s in synthese_jours.items():
        if not isinstance(s, dict):
            continue
        raw_day = s.get("day_type", s_id)
        short_day = raw_day.split("(")[0].strip()
        date_clean = format_readable_date(s.get("date", ""))
        label = f"{short_day} ({date_clean})" if date_clean else short_day
        days_labels.append(label)
        dist_vals.append(float(s.get("total_distance_km", 0.0)))
        hi_vals.append(float(s.get("total_hi_m", 0.0)))

    if not days_labels:
        days_labels = ["Séance 1"]
        dist_vals = [0.0]
        hi_vals = [0.0]

    x = np.arange(len(days_labels))
    width = 0.35

    rects1 = ax1.bar(x - width/2, dist_vals, width, label="Distance Totale (km)", color="#1F4E79", edgecolor="none", zorder=3)
    ax2 = ax1.twinx()
    rects2 = ax2.bar(x + width/2, hi_vals, width, label="Haute Intensité (m)", color="#A30000", edgecolor="none", zorder=3)

    ax1.set_ylabel("Distance (km)", color="#1F4E79", fontweight="bold", fontsize=8.2)
    ax2.set_ylabel("H.I. (m)", color="#A30000", fontweight="bold", fontsize=8.2)
    ax1.set_xticks(x)
    ax1.set_xticklabels(days_labels, fontweight="bold", fontsize=8.2, color="#1E293B")
    ax1.set_title("Volume Journalier par Séance", fontweight="bold", color="#1F4E79", fontsize=9.5, pad=6)

    for rect in rects1:
        h = rect.get_height()
        if h > 0:
            ax1.annotate(f"{h:.1f} km",
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 2), textcoords="offset points",
                         ha="center", va="bottom", fontsize=7.2, fontweight="bold", color="#1F4E79")

    for rect in rects2:
        h = rect.get_height()
        if h > 0:
            ax2.annotate(f"{int(round(h)):,d} m".replace(",", " "),
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 2), textcoords="offset points",
                         ha="center", va="bottom", fontsize=7.2, fontweight="bold", color="#A30000")

    max_dist = max(dist_vals) if dist_vals and max(dist_vals) > 0 else 5.0
    max_hi = max(hi_vals) if hi_vals and max(hi_vals) > 0 else 500.0
    ax1.set_ylim(0, max_dist * 1.25)
    ax2.set_ylim(0, max_hi * 1.25)
    ax1.grid(axis="y", linestyle="--", alpha=0.3, zorder=1)
    ax1.tick_params(axis="y", labelcolor="#1F4E79", labelsize=7.5)
    ax2.tick_params(axis="y", labelcolor="#A30000", labelsize=7.5)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper center", bbox_to_anchor=(0.5, -0.16), ncol=2, frameon=False, fontsize=7.5)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def generer_graphique_hebdo_jauges_matplotlib(
    group_name: str,
    df_grp: pd.DataFrame,
    nb_seances: int = 3,
    objectifs_config: Optional[Dict[str, Dict[str, float]]] = None,
    layout: str = "1x4",
) -> io.BytesIO:
    """
    Génère une figure propre (1x4 ou 2x2) regroupant les 4 duels de barres
    'Objectif Hebdomadaire vs Réalisé Cumulé' pour une ligne de poste ou un poste spécifique RCA :
    1. Distance Totale (m)
    2. HSR Z3 (m)
    3. VHSR Sprint (m)
    4. ACC / DEC (> 3 m/s²)
    """
    active_cfg = objectifs_config or charger_objectifs_config()
    targets_daily = get_objectifs_for_category(active_cfg, group_name)

    multiplier = max(nb_seances, 1)

    # Calcul des réalisés moyens par joueur du groupe sur la semaine
    if not df_grp.empty:
        def _get_mean(df_m: pd.DataFrame, candidates: List[str], mult: float = 1.0) -> float:
            for col in candidates:
                if col in df_m.columns:
                    s = pd.to_numeric(df_m[col], errors="coerce").dropna()
                    if not s.empty:
                        return float(s.mean() * mult)
            return 0.0

        val_dist = _get_mean(df_grp, ["Distance Totale (m)", "Distance (m)"])
        if val_dist == 0.0:
            val_dist = _get_mean(df_grp, ["Distance Totale (km)", "Distance (km)"], mult=1000.0)

        val_hsr = _get_mean(df_grp, ["HSR Z3 (m)", "Haute intensité (m)", "HSR_m", "Distance haute intensité (m)"])
        if val_hsr == 0.0:
            val_hsr = _get_mean(df_grp, ["Distance in Speed Zone 3  (km)"], mult=1000.0)

        val_vhsr = _get_mean(df_grp, ["Distance Sprint VHSR (m)", "VHSR (m)", "Sprint (m)"])
        if val_vhsr == 0.0 and "Distance in Speed Zone 4  (km)" in df_grp.columns and "Distance in Speed Zone 5  (km)" in df_grp.columns:
            s4 = pd.to_numeric(df_grp["Distance in Speed Zone 4  (km)"], errors="coerce").fillna(0.0)
            s5 = pd.to_numeric(df_grp["Distance in Speed Zone 5  (km)"], errors="coerce").fillna(0.0)
            val_vhsr = float((s4 + s5).mean() * 1000.0)

        val_acc = _get_mean(df_grp, ["Accel/Dec total", "Accel/Dec", "ACC/DEC"])
    else:
        val_dist = val_hsr = val_vhsr = val_acc = 0.0

    val_dist = 0.0 if (val_dist is None or np.isnan(val_dist)) else val_dist
    val_hsr = 0.0 if (val_hsr is None or np.isnan(val_hsr)) else val_hsr
    val_vhsr = 0.0 if (val_vhsr is None or np.isnan(val_vhsr)) else val_vhsr
    val_acc = 0.0 if (val_acc is None or np.isnan(val_acc)) else val_acc

    cards_def = [
        ("Distance Totale (m)", targets_daily["Objectif Distance (m)"] * multiplier, val_dist, "m"),
        ("HSR Z3 (m)", targets_daily["Obj. HSR"] * multiplier, val_hsr, "m"),
        ("VHSR Sprint (m)", targets_daily["Obj. VHSR"] * multiplier, val_vhsr, "m"),
        ("ACC / DEC (> 3 m/s²)", targets_daily["Obj. ACC/DEC"] * multiplier, val_acc, ""),
    ]

    if layout == "2x2":
        fig, axes_grid = plt.subplots(2, 2, figsize=(11.5, 4.0), dpi=130)
        axes = axes_grid.flatten()
    else:
        fig, axes = plt.subplots(1, 4, figsize=(11.5, 2.5), dpi=130)
    fig.patch.set_facecolor("#FFFFFF")

    for idx, (title, obj_val, real_val, unit) in enumerate(cards_def):
        ax = axes[idx]
        ax.set_facecolor("#F8FAFC")
        for spine in ax.spines.values():
            spine.set_color("#CBD5E1")
            spine.set_linewidth(0.8)

        pct = (real_val / obj_val * 100.0) if obj_val > 0 else 0.0

        x = np.array([0, 1])
        width = 0.44

        ax.bar(
            x,
            [obj_val, real_val],
            width=width,
            color=["#1F4E79", "#A30000"],
            edgecolor="none",
            zorder=3,
        )

        ax.set_title(title, fontsize=9.2, fontweight="bold", color="#1F4E79", pad=6)
        ax.set_xticks(x)
        ax.set_xticklabels(["Obj.", "Réalisé"], fontsize=8.5, fontweight="bold", color="#334155")

        max_v = max(obj_val, real_val, 1.0)
        ax.set_ylim(0, max_v * 1.35)
        ax.grid(axis="y", linestyle="--", alpha=0.35, color="#94A3B8", zorder=1)
        ax.set_yticklabels([])
        ax.tick_params(axis="both", which="both", length=0)

        def _fmt(v: float, u: str) -> str:
            if abs(v - round(v)) < 0.05:
                s = f"{int(round(v)):,d}".replace(",", " ")
            else:
                s = f"{v:.1f}"
            return f"{s} {u}".strip()

        ax.text(
            0, obj_val + (max_v * 0.03),
            _fmt(obj_val, unit),
            ha="center", va="bottom",
            fontsize=8.0, fontweight="bold", color="#1F4E79"
        )

        ax.text(
            1, real_val + (max_v * 0.03),
            _fmt(real_val, unit),
            ha="center", va="bottom",
            fontsize=8.0, fontweight="bold", color="#A30000"
        )

        # Badge % d'atteinte
        badge_bg = "#DCFCE7" if pct >= 95.0 else ("#FEF3C7" if pct >= 75.0 else "#FEE2E2")
        badge_fg = "#166534" if pct >= 95.0 else ("#92400E" if pct >= 75.0 else "#991B1B")

        ax.text(
            0.5, 0.88,
            f"{pct:.0f}%",
            transform=ax.transAxes,
            ha="center", va="center",
            fontsize=8.0, fontweight="bold",
            color=badge_fg,
            bbox=dict(boxstyle="round,pad=0.25", facecolor=badge_bg, edgecolor="none", alpha=0.95),
            zorder=4,
        )

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


# ==============================================================================
# EXPORT PDF MULTI-PAGES RAPPORT HEBDOMADAIRE (FORMAT PAYSAGE COACHS)
# ==============================================================================

def safe_latin1(s: Any) -> str:
    """Assure qu'une chaîne est compatible avec l'encodage Latin-1 (Helvetica standard FPDF)."""
    if s is None:
        return ""
    txt = str(s)
    replacements = {
        "\u2019": "'",
        "\u2018": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u201e": '"',
        "\u2013": "-",
        "\u2014": "-",
        "\u2022": "-",
        "\u2026": "...",
        "\u0153": "oe",
        "\u0152": "OE",
        "\u20ac": "EUR",
        "\xa0": " ",
        "\u200b": "",
        "\ufeff": "",
    }
    for old, new in replacements.items():
        txt = txt.replace(old, new)
    try:
        return txt.encode("latin-1", errors="replace").decode("latin-1")
    except Exception:
        return "".join([c if ord(c) < 256 else "?" for c in txt])


def calculer_sous_titres_podium(
    df_subset: pd.DataFrame,
    label_prefix: str = "Moyenne",
) -> Dict[str, str]:
    """
    Calcule les moyennes d'équipe ou de groupe pour les sous-titres des cartes podiums.
    Gère gracieusement le cas où le sous-ensemble compte moins de 3 joueurs ou est vide.
    """
    if df_subset is None or not isinstance(df_subset, pd.DataFrame) or df_subset.empty:
        return {
            "Top speed (km/h)": f"{label_prefix} : 0.0 km/h",
            "% de marche": f"{label_prefix} : 0.0 %",
            "Haute intensité (m)": f"{label_prefix} : 0 m",
        }

    # Top Speed
    ts_col = "Top Speed (km/h)"
    if ts_col in df_subset.columns:
        s_ts = pd.to_numeric(df_subset[ts_col], errors="coerce")
        s_ts_pos = s_ts[s_ts > 0]
        vmax_moy = float(s_ts_pos.mean()) if not s_ts_pos.empty else 0.0
    else:
        vmax_moy = 0.0
    if pd.isna(vmax_moy) or np.isnan(vmax_moy):
        vmax_moy = 0.0

    # % Marche
    walk_col = "%Marche " if "%Marche " in df_subset.columns else ("% de marche moyen" if "% de marche moyen" in df_subset.columns else None)
    if walk_col and walk_col in df_subset.columns:
        s_w = pd.to_numeric(df_subset[walk_col], errors="coerce")
        s_w_pos = s_w[s_w > 0.0]
        walk_moy = float(s_w_pos.mean()) if not s_w_pos.empty else 0.0
    else:
        walk_moy = 0.0
    if pd.isna(walk_moy) or np.isnan(walk_moy):
        walk_moy = 0.0
    walk_pct = walk_moy * 100.0 if walk_moy <= 1.0 else walk_moy

    # Haute Intensité
    hi_col = None
    for c in ["Haute intensité Podium (m)", " Distance haute intensité (m)", "Haute intensité (m)"]:
        if c in df_subset.columns:
            hi_col = c
            break
    if hi_col:
        s_hi = pd.to_numeric(df_subset[hi_col], errors="coerce")
        hi_moy = float(s_hi.mean()) if not s_hi.empty else 0.0
    else:
        hi_moy = 0.0
    if pd.isna(hi_moy) or np.isnan(hi_moy):
        hi_moy = 0.0
    hi_int = int(round(hi_moy))

    return {
        "Top speed (km/h)": f"{label_prefix} : {vmax_moy:.1f} km/h",
        "% de marche": f"{label_prefix} : {walk_pct:.1f} %",
        "Haute intensité (m)": f"{label_prefix} : {hi_int:,d} m".replace(",", " "),
    }


# ==============================================================================
# MOTEUR D'ANALYSE COMPARATIVE MULTI-SÉANCES & MULTI-SEMAINES (HISTORIQUE)
# ==============================================================================

CALENDRIER_ADVERSAIRES_RCA: Dict[Any, str] = {
    (2026, 41): "RHEU",
    (2026, 42): "SALLES",
    (2026, 43): "EVREUX",
    (2026, 44): "REPOS",
    (2026, 45): "VERSAILLES",
    (2026, 46): "GENNEVILLIERS",
    (2026, 47): "COURBEVOIE",
    "10/10": "RHEU",
    "11/10": "RHEU",
    "16/10": "SALLES",
    "17/10": "SALLES",
    "18/10": "SALLES",
    "23/10": "EVREUX",
    "24/10": "EVREUX",
    "25/10": "EVREUX",
    "06/11": "VERSAILLES",
    "07/11": "VERSAILLES",
    "08/11": "VERSAILLES",
}


def parse_session_date(val: Any) -> Optional[datetime.date]:
    """
    Convertit de manière déterministe une valeur de date quelconque en objet datetime.date.
    Supporte :
    - datetime.date, datetime.datetime, pd.Timestamp
    - Numéros de série Excel (ex: 46332 ou "46332" ou 46270)
    - Chaînes ISO (YYYY-MM-DD) et françaises (JJ/MM/AAAA ou JJ-MM-AAAA)
    - Ne bascule JAMAIS vers datetime.now() ou date.today().
    """
    if val is None or pd.isna(val):
        return None
    if isinstance(val, (datetime.datetime, pd.Timestamp)):
        return val.date()
    if isinstance(val, datetime.date):
        return val

    val_str = str(val).strip()
    if not val_str or val_str.lower() in ["nan", "none", "nat", ""]:
        return None

    # Test si numéro de série Excel (ex: 46332 ou 46332.0)
    try:
        f = float(val_str.replace(",", "."))
        if 30000 <= f <= 80000:
            return pd.to_datetime(f, unit="D", origin="1899-12-30").date()
    except (ValueError, TypeError):
        pass

    # Format ISO : YYYY-MM-DD ou YYYY/MM/DD
    m_iso = re.search(r"(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})", val_str)
    if m_iso:
        try:
            y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
            return datetime.date(y, m, d)
        except Exception:
            pass

    # Format français : JJ/MM/AAAA ou JJ-MM-AAAA
    m_fr = re.search(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})", val_str)
    if m_fr:
        try:
            d, m, y = int(m_fr.group(1)), int(m_fr.group(2)), int(m_fr.group(3))
            if y < 100:
                y += 2000
            return datetime.date(y, m, d)
        except Exception:
            pass

    try:
        dt = pd.to_datetime(val_str, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            return dt.date()
    except Exception:
        pass

    return None


def detect_day_of_week(val: Any) -> str:
    """
    Détecte le jour de la semaine en français (Lundi, Mardi, Mercredi, Jeudi, Vendredi, Samedi, Dimanche).
    Fonctionne avec une date, un numéro de série Excel (ex: 46332), ou un nom de jour textuel.
    Ne bascule JAMAIS sur la date système today().
    """
    JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
    if val is None:
        return ""

    dt = parse_session_date(val)
    if dt is not None:
        return JOURS[dt.weekday()]

    val_str = str(val).strip().lower()
    for j in JOURS:
        if j.lower() in val_str:
            return j

    return ""


def _parse_date_to_obj(
    date_val: Any,
    fallback_year: int = 2026,
    fallback_week: int = 1,
    fallback_day_name: str = "",
) -> Optional[datetime.date]:
    """
    Convertit une valeur de date quelconque en objet datetime.date pour tri chronologique strict.
    Ne bascule JAMAIS sur la date système today().
    """
    dt = parse_session_date(date_val)
    if dt is not None:
        return dt

    JOURS_MAP = {"lundi": 1, "mardi": 2, "mercredi": 3, "jeudi": 4, "vendredi": 5, "samedi": 6, "dimanche": 7}
    day_num = JOURS_MAP.get(str(fallback_day_name or "").strip().lower(), None)
    if day_num is not None:
        try:
            return datetime.date.fromisocalendar(fallback_year, max(1, min(53, fallback_week)), day_num)
        except Exception:
            pass

    return None


def extraire_adversaire_seance(
    date_obj: Optional[datetime.date],
    day_name: str = "",
    adversaire: str = "",
    notes: str = "",
    week_num: Optional[int] = None,
) -> str:
    """
    Extrait le nom de l'adversaire selon la règle du sprint :
    1. Si `adversaire` est explicitement défini (non vide), l'utiliser en majuscules.
    2. Règle de rétrocompatibilité (CRITIQUE) :
       - Si la date de la séance correspond au vendredi de la S43 (ex: 23/10 ou semaine 43), l'adversaire par défaut est "EVREUX".
       - Si la date correspond au vendredi de la S45 (ex: 06/11 ou semaine 45), l'adversaire par défaut est "VERSAILLES".
    3. Pour toutes les autres dates : utilise la valeur du champ Adversaire lue dans la base historique.
       Si elle est vide, retourne une chaîne vide.
    """
    adv_clean = str(adversaire or "").strip()
    if adv_clean:
        return adv_clean.upper()

    # Détection si c'est un vendredi
    is_friday = ("vendredi" in str(day_name or "").lower())
    if date_obj is not None and date_obj.weekday() == 4:
        is_friday = True

    # Règle de rétrocompatibilité pour S43 et S45 (Vendredis historiques)
    if is_friday:
        if date_obj is not None:
            y, w, _ = date_obj.isocalendar()
            d_str = date_obj.strftime("%d/%m")
            if w == 43 or d_str == "23/10":
                return "EVREUX"
            if w == 45 or d_str == "06/11":
                return "VERSAILLES"
        if week_num == 43:
            return "EVREUX"
        if week_num == 45:
            return "VERSAILLES"

    return ""


def formater_libelle_seance_x(
    date_obj: datetime.date,
    day_name: str,
    adversaire: str = "",
    notes: str = "",
    week_num: Optional[int] = None,
) -> str:
    """
    Formate le libellé pour l'axe X des graphiques comparatifs :
    Format : {jour_semaine} {jour/mois} + ' ({ADVERSAIRE})' si un adversaire est défini.
    Si l'adversaire est vide, n'affiche que la date sans parenthèses.
    """
    j_str = (day_name or "séance").strip().lower()
    d_str = date_obj.strftime("%d/%m") if date_obj else ""
    opp = extraire_adversaire_seance(
        date_obj=date_obj,
        day_name=day_name,
        adversaire=adversaire,
        notes=notes,
        week_num=week_num,
    )
    if opp:
        return f"{j_str} {d_str} ({opp})"
    return f"{j_str} {d_str}".strip()


CANONICAL_9_POSTES: List[str] = [
    "Pilier", "Talonneur", "2ème ligne", "3ème ligne",
    "Demi de mêlée", "Demi d'ouverture", "Centre", "Ailier", "Arrière",
]

POSTE_TO_GROUPE_MAP: Dict[str, str] = {
    "Pilier": "5 de devant",
    "Talonneur": "5 de devant",
    "2ème ligne": "5 de devant",
    "3ème ligne": "3ème Ligne",
    "Flanker": "3ème Ligne",
    "Numéro 8": "3ème Ligne",
    "Numero 8": "3ème Ligne",
    "Demi de mêlée": "Trois-quarts",
    "Demi d'ouverture": "Trois-quarts",
    "Centre": "Trois-quarts",
    "Ailier": "Trois-quarts",
    "Arrière": "Trois-quarts",
}


def is_grand_groupe(category: str, is_detailed_context: bool = False) -> bool:
    """
    Détermine si une catégorie correspond à une grande ligne (5 de devant, 3ème Ligne, Trois-quarts)
    ou à un poste individuel / spécifique (Pilier, Talonneur, Demi de mêlée, etc.).
    """
    c_norm = _strip_accents_lower(str(category or "")).strip()
    if c_norm in ("5 de devant", "cinq de devant", "devant", "avants", "trois quarts", "trois-quarts", "3 4", "arrieres", "arrières"):
        return True
    if c_norm in ("3eme ligne", "3e ligne", "troisieme ligne", "3ème ligne"):
        return not is_detailed_context
    return False


def build_custom_target_groups(
    choice_5: str,
    choice_3: str,
    choice_tq: str,
    df: Optional[pd.DataFrame] = None,
    only_active: bool = False,
    joueurs_lookup: Optional[Dict[str, Any]] = None,
) -> List[str]:
    """
    Construit la liste dynamique `target_groups` à partir des sélections de l'UI :
    - 5 de devant : 'Ensemble (1 groupe)' -> ['5 de devant']
                    'Séparés (Pilier, Talonneur, 2ème ligne)' -> ['Pilier', 'Talonneur', '2ème ligne']
    - 3ème ligne : 'Ensemble (1 groupe)' -> ['3ème ligne']
                   'Séparés (Flanker, Numéro 8)' -> ['Flanker', 'Numéro 8'] si présents, sinon ['3ème ligne']
    - Trois-quarts : 'Ensemble (1 groupe)' -> ['Trois-quarts']
                     'Séparés (Demis, Centres, Ailiers, Arrières)' -> ['Demi de mêlée', "Demi d'ouverture", 'Centre', 'Ailier', 'Arrière']
    """
    groups: List[str] = []

    # 1. 5 de devant
    c5 = str(choice_5 or "").lower()
    if "séparé" in c5 or "separe" in c5:
        groups.extend(["Pilier", "Talonneur", "2ème ligne"])
    else:
        groups.append("5 de devant")

    # 2. 3ème ligne
    c3 = str(choice_3 or "").lower()
    if "séparé" in c3 or "separe" in c3:
        has_f8 = False
        if df is not None and not df.empty:
            for c in ("Poste_Detaille", "Poste"):
                if c in df.columns and df[c].astype(str).str.lower().str.contains("flanker|numero 8|numéro 8|n8").any():
                    has_f8 = True
                    break
        if has_f8:
            groups.extend(["Flanker", "Numéro 8"])
        else:
            groups.append("3ème ligne")
    else:
        groups.append("3ème ligne")

    # 3. Trois-quarts
    ctq = str(choice_tq or "").lower()
    if "séparé" in ctq or "separe" in ctq:
        groups.extend(["Demi de mêlée", "Demi d'ouverture", "Centre", "Ailier", "Arrière"])
    else:
        groups.append("Trois-quarts")

    if only_active and df is not None and not df.empty:
        active_res = []
        for g in groups:
            sub = _filter_df_by_group(df, g)
            if not sub.empty:
                active_res.append(g)
        return active_res or groups

    return groups


def get_grand_groupe_isolated_excluded(
    grp: str,
    isolated_positions: Optional[List[str]] = None,
) -> List[str]:
    """
    Retourne la liste ordonnée des postes isolés appartenant à cette grande ligne.
    Ex: grp='5 de devant', isolated=['Pilier'] -> ['Pilier']
    """
    if not isolated_positions:
        return []
    grp_clean = _strip_accents_lower(grp).strip()
    res = []
    for pos in isolated_positions:
        parent = POSTE_TO_GROUPE_MAP.get(pos, "")
        p_clean = _strip_accents_lower(parent).strip()
        if (
            p_clean == grp_clean
            or (grp_clean in ["5 de devant", "devant"] and parent == "5 de devant")
            or (grp_clean in ["3eme ligne", "3e ligne", "troisieme ligne", "3ème ligne"] and parent == "3ème Ligne")
            or (grp_clean in ["trois quarts", "trois-quarts", "3/4", "arrieres"] and parent == "Trois-quarts")
        ):
            if pos not in res:
                res.append(pos)
    return res


def get_group_display_name(
    grp: str,
    isolated_positions: Optional[List[str]] = None,
    is_isolated: bool = False,
) -> str:
    """
    Retourne l'intitulé d'affichage pour une ligne ou un poste.
    Si la ligne est amputée de certains postes, ajoute '(hors Pos1, Pos2)'.
    Ex: '5 DE DEVANT (hors Pilier)' ou 'PILIER'
    """
    grp_str = str(grp or "").strip()
    if is_isolated:
        return grp_str.upper()
    excluded = get_grand_groupe_isolated_excluded(grp_str, isolated_positions)
    if excluded:
        return f"{grp_str.upper()} (hors {', '.join(excluded)})"
    return grp_str.upper()


def resolve_active_target_groups(
    df: Optional[pd.DataFrame] = None,
    target_groups: Optional[List[str]] = None,
    isolated_positions: Optional[List[str]] = None,
    joueurs_lookup: Optional[Dict[str, Any]] = None,
) -> List[str]:
    """
    Retourne la liste des libellés de groupes et postes actifs qui seront effectivement
    générés sur le PDF, en appliquant l'exclusion des postes isolés et la règle anti-vide.
    """
    if not target_groups:
        return ["5 de devant", "3ème Ligne", "Trois-quarts"]

    isolated = list(isolated_positions or [])
    active_labels = []

    for idx, grp in enumerate(target_groups):
        is_iso = (idx >= 3 and grp in isolated) or (grp in isolated and not is_grand_groupe(grp))
        if df is not None and not df.empty:
            df_sub = _filter_df_by_group(
                df, grp, isolated_positions=isolated if not is_iso else None, is_isolated=is_iso
            )
            # Règle anti-vide : si une grande ligne est vide après exclusion, on l'ignore !
            if not is_iso and is_grand_groupe(grp) and df_sub.empty:
                continue
            if is_iso and df_sub.empty and len(target_groups) > 1:
                continue

        disp = get_group_display_name(grp, isolated_positions=isolated, is_isolated=is_iso)
        active_labels.append(disp)

    return active_labels or target_groups

_REGISTERED_JOUEURS_LOOKUP: Dict[str, str] = {}


def register_joueurs_lookup(mapping: Optional[Dict[str, str]]) -> None:
    """Enregistre le dictionnaire {Nom joueur: Poste} de l'effectif pour la normalisation."""
    global _REGISTERED_JOUEURS_LOOKUP
    if isinstance(mapping, dict):
        _REGISTERED_JOUEURS_LOOKUP = dict(mapping)


def _get_effective_joueurs_lookup() -> Dict[str, str]:
    """Retourne le référentiel effectif enregistré ou charge celui de gps_processing."""
    if _REGISTERED_JOUEURS_LOOKUP:
        return _REGISTERED_JOUEURS_LOOKUP
    try:
        try:
            from modules import gps_processing
        except (ImportError, ModuleNotFoundError):
            import gps_processing
        j_data = getattr(gps_processing, "joueurs_data", None)
        if isinstance(j_data, dict) and j_data:
            return j_data
    except Exception:
        pass
    return {}


def _strip_accents_lower(val: Any) -> str:
    import unicodedata
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    s = unicodedata.normalize("NFKD", str(val).strip().lower())
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[-_'’./,;()]+", " ", s)
    return " ".join(s.split())


def normaliser_poste_specifique(
    poste_raw: Any,
    player_name: str = "",
    joueurs_lookup: Optional[Dict[str, str]] = None,
) -> str:
    """
    Normalise un libellé de poste (y compris fautes/variantes : 'Allier', 'Demi de mélée',
    '2eme ligne', 'Pilier droit', etc.) vers l'un des 9 postes canoniques RCA.
    Si poste_raw est générique ('5 de devant', 'Trois-quarts') ou vide, résout via joueurs_lookup.
    """
    def _map_token(txt: Any) -> str:
        k = _strip_accents_lower(txt)
        if not k or k in ("5 de devant", "trois quarts", "3 4", "arrieres", "avants", "non specifie", "inconnu", "-"):
            return ""
        if any(w in k for w in ("pilier", "1ere ligne", "premiere ligne", "prop")):
            return "Pilier"
        if any(w in k for w in ("talonneur", "talon", "hooker")):
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

    canon = _map_token(poste_raw)
    if canon:
        return canon

    lookup = joueurs_lookup if isinstance(joueurs_lookup, dict) else _get_effective_joueurs_lookup()
    if player_name and lookup:
        p_norm = _strip_accents_lower(player_name)
        p_tokens = set(p_norm.split())
        for k_ref, v_ref in lookup.items():
            k_norm = _strip_accents_lower(k_ref)
            if p_norm == k_norm or (p_tokens and p_tokens == set(k_norm.split())):
                c2 = _map_token(v_ref)
                if c2:
                    return c2

    return str(poste_raw).strip() if poste_raw is not None and not pd.isna(poste_raw) else ""


def get_active_positions(
    df: pd.DataFrame,
    joueurs_lookup: Optional[Dict[str, str]] = None,
) -> List[str]:
    """
    Retourne la liste ordonnée des postes parmi CANONICAL_9_POSTES effectivement
    représentés (> 0 joueur) dans df.
    """
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return []
    lookup = joueurs_lookup if isinstance(joueurs_lookup, dict) else _get_effective_joueurs_lookup()
    active_set = set()
    for _, row in df.iterrows():
        p_name = str(row.get("Nom ; Prénom", row.get("Player Name", row.get("Nom", ""))) or "")
        p_raw = row.get("Poste_Detaille", row.get("Poste", row.get("Groupe", "")))
        canon = normaliser_poste_specifique(p_raw, player_name=p_name, joueurs_lookup=lookup)
        if canon in CANONICAL_9_POSTES:
            active_set.add(canon)
    return [p for p in CANONICAL_9_POSTES if p in active_set]


def get_objectifs_for_category(
    config_objectifs: Optional[Dict[str, Dict[str, float]]],
    category_name: str,
) -> Dict[str, float]:
    """
    Retourne les objectifs athlétiques pour un groupe (5 de devant, 3ème Ligne, Trois-quarts)
    ou un poste spécifique (en héritant de son groupe parent).
    """
    cfg = config_objectifs or DEFAULT_OBJECTIFS_CONFIG
    if category_name in cfg:
        return cfg[category_name]
    parent_grp = POSTE_TO_GROUPE_MAP.get(category_name)
    if parent_grp and parent_grp in cfg:
        return cfg[parent_grp]
    return cfg.get("Tous les groupes", DEFAULT_OBJECTIFS_CONFIG["Tous les groupes"])


def _filter_df_by_group(
    df: pd.DataFrame,
    grp: str,
    isolated_positions: Optional[List[str]] = None,
    is_isolated: bool = False,
) -> pd.DataFrame:
    """
    Filtre un DataFrame de séance/semaine par groupe de poste ('5 de devant', '3ème Ligne',
    'Trois-quarts') OU par poste spécifique ('Pilier', 'Ailier', etc.).
    Supporte l'exclusion dynamique des postes isolés (isolated_positions) sur les grandes lignes.
    """
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()

    grp_str = str(grp or "").strip()
    grp_clean = _strip_accents_lower(grp_str)
    lookup = _get_effective_joueurs_lookup()
    col_pd = "Poste_Detaille" if "Poste_Detaille" in df.columns else ("Poste" if "Poste" in df.columns else None)
    col_nm = "Nom ; Prénom" if "Nom ; Prénom" in df.columns else ("Player Name" if "Player Name" in df.columns else ("Nom" if "Nom" in df.columns else None))

    # Cas 1 : Poste spécifique isolé ou demandé directement
    if is_isolated or (grp_str in CANONICAL_9_POSTES and not is_grand_groupe(grp_str)):
        if col_pd or col_nm:
            mask = df.apply(
                lambda r: normaliser_poste_specifique(
                    r.get(col_pd, "") if col_pd else "",
                    player_name=str(r.get(col_nm, "")) if col_nm else "",
                    joueurs_lookup=lookup,
                ) == grp_str,
                axis=1,
            )
            return df[mask]
        return pd.DataFrame()

    # Support spécifique des sous-postes de 3ème ligne si demandés explicitement
    if grp_clean in ["flanker", "troisieme ligne aile", "flankers"]:
        if col_pd:
            mask = df[col_pd].astype(str).str.lower().str.contains("flanker|troisieme ligne aile|troisième ligne aile")
            return df[mask]
    if grp_clean in ["numero 8", "n8", "troisieme ligne centre"]:
        if col_pd:
            mask = df[col_pd].astype(str).str.lower().str.contains("numero 8|numéro 8|n8|troisieme ligne centre|troisième ligne centre")
            return df[mask]

    # Cas 2 : Grande ligne de poste
    df_base = pd.DataFrame()
    if "Groupe" in df.columns:
        if grp_clean in ["5 de devant", "devant"]:
            mask = df["Groupe"].astype(str).str.strip().str.lower().isin(["5 de devant", "devant", "cinq de devant", "avants (5)"])
        elif grp_clean in ["3ème ligne", "3eme ligne", "3e ligne"]:
            mask = df["Groupe"].astype(str).str.strip().str.lower().isin(["3ème ligne", "3eme ligne", "3e ligne", "troisième ligne"])
        elif grp_clean in ["trois-quarts", "trois quarts", "3/4", "arrieres", "arrières"]:
            mask = df["Groupe"].astype(str).str.strip().str.lower().isin(["trois-quarts", "trois quarts", "arrières", "arrieres", "3/4"])
        else:
            mask = df["Groupe"].astype(str).str.strip().str.lower() == grp_clean
        df_base = df[mask]

    if df_base.empty and col_pd:
        if grp_clean in ["5 de devant", "devant"]:
            kw = ["pilier", "talonneur", "2ème ligne", "2eme ligne", "deuxième ligne", "première ligne", "5 de devant"]
        elif grp_clean in ["3ème ligne", "3eme ligne", "3e ligne"]:
            kw = ["3ème ligne", "3eme ligne", "troisième ligne", "flanker", "n8", "numéro 8"]
        else:
            kw = ["demi", "ouvreur", "ouverture", "mêlée", "melee", "centre", "ailier", "allier", "arrière", "arriere", "trois-quarts"]
        mask = df[col_p].astype(str).str.strip().str.lower().apply(lambda s: any(k in s for k in kw))
        df_base = df[mask]

    if df_base.empty and grp_str in CANONICAL_9_POSTES:
        mask = df.apply(
            lambda r: normaliser_poste_specifique(
                r.get(col_pd, "") if col_pd else "",
                player_name=str(r.get(col_nm, "")) if col_nm else "",
                joueurs_lookup=lookup,
            ) == grp_str,
            axis=1,
        )
        df_base = df[mask]

    # Application de l'exclusion des postes isolés pour cette grande ligne
    if isolated_positions and not df_base.empty:
        iso_clean = set(_strip_accents_lower(p) for p in isolated_positions)
        mask_keep = df_base.apply(
            lambda r: _strip_accents_lower(normaliser_poste_specifique(
                r.get(col_pd, "") if col_pd else "",
                player_name=str(r.get(col_nm, "")) if col_nm else "",
                joueurs_lookup=lookup,
            )) not in iso_clean,
            axis=1,
        )
        df_base = df_base[mask_keep]

    return df_base


def calculer_metriques_groupes_seance(
    df_s: pd.DataFrame,
    categories: Optional[List[str]] = None,
    isolated_positions: Optional[List[str]] = None,
) -> Dict[str, Dict[str, Any]]:
    """
    Calcule les moyennes exactes par ligne de poste ou par poste spécifique pour le triptyque :
    Retourne {categorie: {"volume": m, "hsr": m, "vhsr": m, "rpe": float, "num_players": n}}
    """
    groupes = categories if categories else ["5 de devant", "3ème Ligne", "Trois-quarts"]
    res: Dict[str, Dict[str, Any]] = {
        g: {"volume": 0.0, "hsr": 0.0, "vhsr": 0.0, "rpe": None, "num_players": 0}
        for g in groupes
    }
    if df_s is None or not isinstance(df_s, pd.DataFrame) or df_s.empty:
        return res

    isolated_set = set(isolated_positions or [])
    for grp in groupes:
        is_iso = grp in isolated_set and not is_grand_groupe(grp)
        df_sub = _filter_df_by_group(df_s, grp, isolated_positions=isolated_positions if not is_iso else None, is_isolated=is_iso)
        n_p = len(df_sub)
        res[grp]["num_players"] = n_p
        if n_p == 0:
            continue

        # 1. Volume / Distance totale (m)
        vols = []
        for _, r in df_sub.iterrows():
            d = _extract_metric_dist_m(r)
            vols.append(d)
        res[grp]["volume"] = round(float(np.mean(vols)), 1) if vols else 0.0

        # 2. HSR (> 20 km/h ou Z3) (m)
        hsrs = []
        for _, r in df_sub.iterrows():
            h = _extract_metric_hsr_m(r)
            hsrs.append(h)
        res[grp]["hsr"] = round(float(np.mean(hsrs)), 1) if hsrs else 0.0

        # 3. VHSR / Sprint (> 25 km/h ou Z5) (m)
        vhsrs = []
        for _, r in df_sub.iterrows():
            vh = 0.0
            for col in ["VHSR (m)", "Distance Sprint VHSR (m)", "VHSR_m"]:
                if col in r and r[col] is not None and not pd.isna(r[col]):
                    try:
                        vh = float(r[col])
                        break
                    except Exception:
                        pass
            if vh == 0.0 and "Distance in Speed Zone 5  (km)" in r and pd.notna(r["Distance in Speed Zone 5  (km)"]):
                try:
                    vh = float(r["Distance in Speed Zone 5  (km)"]) * 1000.0
                except Exception:
                    pass
            vhsrs.append(vh)
        res[grp]["vhsr"] = round(float(np.mean(vhsrs)), 1) if vhsrs else 0.0

        # 4. RPE Moyen du groupe (1-10)
        rpes = []
        for _, r in df_sub.iterrows():
            v_rpe = _extract_metric_rpe(r)
            if v_rpe is not None and v_rpe > 0:
                rpes.append(v_rpe)
        res[grp]["rpe"] = round(float(np.mean(rpes)), 1) if rpes else None

    return res


def _is_avant(row: Any) -> bool:
    """Détecte si un joueur appartient aux Avants (5 de devant + 3ème Ligne)."""
    grp = str(row.get("Groupe", "")).strip().lower() if hasattr(row, "get") else ""
    pos = str(row.get("Poste_Detaille", row.get("Poste", ""))).strip().lower() if hasattr(row, "get") else ""
    if any(k in grp for k in ["5 de devant", "devant", "3ème ligne", "3eme ligne", "avants", "pack"]):
        return True
    if any(k in pos for k in ["pilier", "talonneur", "deuxième ligne", "2ème ligne", "troisième ligne", "3ème ligne", "flanker", "n8", "numéro 8"]):
        return True
    return False


def _is_arriere(row: Any) -> bool:
    """Détecte si un joueur appartient aux Arrières (Trois-quarts / charnière)."""
    grp = str(row.get("Groupe", "")).strip().lower() if hasattr(row, "get") else ""
    pos = str(row.get("Poste_Detaille", row.get("Poste", ""))).strip().lower() if hasattr(row, "get") else ""
    if any(k in grp for k in ["trois-quarts", "trois quarts", "arrières", "arrieres", "3/4", "backs"]):
        return True
    if any(k in pos for k in ["demi", "mêlée", "ouverture", "centre", "ailier", "arrière", "arriere", "10", "9", "12", "13", "11", "14", "15"]):
        return True
    return False


def _extract_metric_dist_m(row: Any) -> float:
    for col in ["Distance Totale (m)", "Distance_m"]:
        if col in row and pd.notna(row[col]):
            try: return float(row[col])
            except Exception: pass
    for col in ["Distance (km)", "Distance Totale (km)"]:
        if col in row and pd.notna(row[col]):
            try: return float(row[col]) * 1000.0
            except Exception: pass
    return 0.0


def _extract_metric_hsr_m(row: Any) -> float:
    for col in ["HSR Z3 (m)", " Distance haute intensité (m)", "Haute intensité (m)", "Haute intensité Podium (m)", "HSR_m"]:
        if col in row and pd.notna(row[col]):
            try: return float(row[col])
            except Exception: pass
    return 0.0


def _extract_metric_acc_dec(row: Any) -> float:
    for col in ["Accel/Dec total", "Accel/Dec"]:
        if col in row and pd.notna(row[col]):
            try: return float(row[col])
            except Exception: pass
    return 0.0


def _extract_metric_dpm(row: Any) -> float:
    for col in ["Distance Per Min (m/min)", "Distance/min (m/min)", "Distance_Per_Min"]:
        if col in row and pd.notna(row[col]):
            try: return float(row[col])
            except Exception: pass
    return 0.0


def _extract_metric_player_load(row: Any) -> float:
    for col in ["Player Load", "PlayerLoad", "Charge", "Impacts", "Total Impacts"]:
        if col in row and pd.notna(row[col]):
            try:
                v = float(row[col])
                if v > 0:
                    return v
            except Exception: pass
    d_m = _extract_metric_dist_m(row)
    acc = _extract_metric_acc_dec(row)
    return round(d_m * 0.08 + acc * 1.8, 1)


def _extract_metric_rpe(row: Any) -> Optional[float]:
    for col in ["RPE", "Score RPE", "rpe"]:
        if col in row and row[col] is not None and not pd.isna(row[col]):
            try:
                v = float(row[col])
                if 0.0 <= v <= 10.0:
                    return v
            except Exception:
                pass
    return None


def _extract_metric_session_rpe(row: Any) -> Optional[float]:
    for col in ["Session_RPE", "Session RPE", "Charge RPE", "Foster", "session_rpe"]:
        if col in row and row[col] is not None and not pd.isna(row[col]):
            try:
                v = float(row[col])
                if v > 0:
                    return v
            except Exception:
                pass
    return None


def calculer_moyennes_seance(df: pd.DataFrame) -> Dict[str, Dict[str, float]]:
    """
    Calcule les moyennes collectives globales et par groupe (Avants vs Arrières)
    sur les métriques clés spécifiées :
    1. Distance Totale (m)
    2. High Speed Running (> 20 km/h) (m)
    3. Accélérations & Décélérations intenses (> 3 m/s²)
    4. Mètres par minute (m/min)
    5. Player Load / Charge globale
    6. RPE Moyen (1-10)
    7. Session-RPE (Foster = RPE × Durée)
    """
    res: Dict[str, Dict[str, float]] = {
        "Distance Totale (m)": {"Total": 0.0, "Avants": 0.0, "Arrières": 0.0},
        "HSR (> 20 km/h) (m)": {"Total": 0.0, "Avants": 0.0, "Arrières": 0.0},
        "Acc/Dec (> 3 m/s²)": {"Total": 0.0, "Avants": 0.0, "Arrières": 0.0},
        "m/min": {"Total": 0.0, "Avants": 0.0, "Arrières": 0.0},
        "Player Load": {"Total": 0.0, "Avants": 0.0, "Arrières": 0.0},
        "RPE Moyen (1-10)": {"Total": 0.0, "Avants": 0.0, "Arrières": 0.0},
        "Session-RPE (Foster)": {"Total": 0.0, "Avants": 0.0, "Arrières": 0.0},
    }
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return res

    is_av = df.apply(_is_avant, axis=1)
    is_ar = df.apply(_is_arriere, axis=1)
    if not is_ar.any() and is_av.any():
        is_ar = ~is_av
    elif not is_av.any() and is_ar.any():
        is_av = ~is_ar

    df_av = df[is_av]
    df_ar = df[is_ar]

    metric_extractors = [
        ("Distance Totale (m)", _extract_metric_dist_m),
        ("HSR (> 20 km/h) (m)", _extract_metric_hsr_m),
        ("Acc/Dec (> 3 m/s²)", _extract_metric_acc_dec),
        ("m/min", _extract_metric_dpm),
        ("Player Load", _extract_metric_player_load),
        ("RPE Moyen (1-10)", _extract_metric_rpe),
        ("Session-RPE (Foster)", _extract_metric_session_rpe),
    ]

    for m_name, extractor in metric_extractors:
        vals_tot = [extractor(r) for _, r in df.iterrows()]
        vals_av = [extractor(r) for _, r in df_av.iterrows()] if not df_av.empty else []
        vals_ar = [extractor(r) for _, r in df_ar.iterrows()] if not df_ar.empty else []

        if "RPE" in m_name:
            v_tot = [v for v in vals_tot if v is not None and not pd.isna(v) and float(v) > 0]
            v_av = [v for v in vals_av if v is not None and not pd.isna(v) and float(v) > 0]
            v_ar = [v for v in vals_ar if v is not None and not pd.isna(v) and float(v) > 0]
            res[m_name]["Total"] = round(float(np.mean(v_tot)), 1) if v_tot else 0.0
            res[m_name]["Avants"] = round(float(np.mean(v_av)), 1) if v_av else 0.0
            res[m_name]["Arrières"] = round(float(np.mean(v_ar)), 1) if v_ar else 0.0
        else:
            res[m_name]["Total"] = round(float(np.mean(vals_tot)), 1) if vals_tot else 0.0
            res[m_name]["Avants"] = round(float(np.mean(vals_av)), 1) if vals_av else 0.0
            res[m_name]["Arrières"] = round(float(np.mean(vals_ar)), 1) if vals_ar else 0.0

    return res


def recuperer_historique_seances_meme_jour(
    session_date_or_day: Any,
    df_current: pd.DataFrame,
    n_max: int = 4,
    current_notes_or_tag: str = "",
    current_adversaire: str = "",
    current_rpe_stats: Optional[Dict[str, Any]] = None,
    categories: Optional[List[str]] = None,
    isolated_positions: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    """
    Extrait l'historique des N dernières séances enregistrées correspondant au même jour de la semaine
    (ex. si séance du vendredi -> les N derniers vendredis disponibles, séance du jour incluse).
    Rattrape automatiquement les données RPE manquantes sur les séances passées via Google Sheets.
    Gère automatiquement les semaines creuses sans entraînement et les modes 3 groupes / 9 postes.
    Ne dépend JAMAIS de la date système du jour.
    """
    target_cats = list(categories) if categories else ["5 de devant", "3ème Ligne", "Trois-quarts"]
    all_cats = list(dict.fromkeys(["5 de devant", "3ème Ligne", "Trois-quarts"] + target_cats))

    dt_curr = parse_session_date(session_date_or_day)
    target_day = detect_day_of_week(session_date_or_day)

    if dt_curr is not None:
        target_weekday = dt_curr.weekday()
        if not target_day:
            JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
            target_day = JOURS[target_weekday]
        curr_iso_year, curr_iso_week, _ = dt_curr.isocalendar()
    else:
        JOURS_MAP = {"lundi": 0, "mardi": 1, "mercredi": 2, "jeudi": 3, "vendredi": 4, "samedi": 5, "dimanche": 6}
        target_weekday = JOURS_MAP.get(str(target_day or "").strip().lower(), 4)
        if not target_day:
            target_day = "Vendredi"
        curr_iso_year, curr_iso_week = 2026, 45
        dt_curr = _parse_date_to_obj(session_date_or_day, fallback_year=curr_iso_year, fallback_week=curr_iso_week, fallback_day_name=target_day)
        if dt_curr is None:
            dt_curr = datetime.date(curr_iso_year, 11, 6)

    matched_sessions: Dict[str, Dict[str, Any]] = {}
    saved_weeks = list_saved_weeks()

    df_rpe_cache: Optional[pd.DataFrame] = None
    rpe_fetched = False

    def _get_rpe_sheet_df() -> Optional[pd.DataFrame]:
        nonlocal df_rpe_cache, rpe_fetched
        if not rpe_fetched:
            rpe_fetched = True
            if rpe_connector and hasattr(rpe_connector, "fetch_raw_rpe_data"):
                try:
                    df_rpe_cache = rpe_connector.fetch_raw_rpe_data()
                except Exception:
                    df_rpe_cache = None
        return df_rpe_cache

    j_lookup = _get_effective_joueurs_lookup()
    isolated_set = set(isolated_positions or [])

    for w_info in saved_weeks:
        y = w_info["year"]
        w = w_info["week_num"]
        w_data = load_week_data(y, w)

        # Rattrapage préventif des RPE manquants sur cette semaine historique
        if rpe_connector and hasattr(rpe_connector, "enrich_historical_week_data_with_rpe"):
            sessions_map = w_data.get("sessions", {}) if isinstance(w_data, dict) else {}
            needs_rpe = False
            for s in sessions_map.values():
                if not isinstance(s, dict):
                    continue
                s_day_raw = s.get("day_type", "") + " " + s.get("date", "")
                s_day = detect_day_of_week(s_day_raw)
                s_dt = _parse_date_to_obj(s.get("date"), fallback_year=y, fallback_week=w, fallback_day_name=s_day or target_day)
                if s_dt is None:
                    continue
                if (s_dt.weekday() == target_weekday) or (s_day and s_day.lower() == target_day.lower()):
                    p_list_tmp = s.get("players_data", [])
                    has_rpe = any(
                        isinstance(p, dict) and p.get("RPE") not in (None, "", "NR")
                        and not (isinstance(p.get("RPE"), float) and np.isnan(p.get("RPE")))
                        for p in p_list_tmp
                    )
                    if not has_rpe:
                        needs_rpe = True
                        break
            if needs_rpe:
                try:
                    df_rpe_raw = _get_rpe_sheet_df()
                    if df_rpe_raw is not None and not df_rpe_raw.empty:
                        w_data, changed = rpe_connector.enrich_historical_week_data_with_rpe(
                            w_data, df_rpe=df_rpe_raw, joueurs_lookup=j_lookup
                        )
                        if changed:
                            fpath = get_week_file_path(y, w)
                            try:
                                with open(fpath, "w", encoding="utf-8") as fw:
                                    json.dump(w_data, fw, indent=2, ensure_ascii=False, default=str)
                            except Exception:
                                pass
                except Exception:
                    pass

        for s_id, s in w_data.get("sessions", {}).items():
            if not isinstance(s, dict):
                continue
            s_day_raw = s.get("day_type", "") + " " + s.get("date", "")
            s_day = detect_day_of_week(s_day_raw)
            s_dt = _parse_date_to_obj(s.get("date"), fallback_year=y, fallback_week=w, fallback_day_name=s_day or target_day)
            if s_dt is None:
                continue

            is_same_weekday = (s_dt.weekday() == target_weekday)
            is_same_dayname = bool(s_day and s_day.lower() == target_day.lower())

            if is_same_weekday or is_same_dayname:
                date_key = s_dt.strftime("%Y-%m-%d")
                p_list = s.get("players_data", [])
                if p_list:
                    df_s = pd.DataFrame(p_list)
                    if gsheet_sync and hasattr(gsheet_sync, "filter_out_test_players"):
                        df_s = gsheet_sync.filter_out_test_players(df_s)
                    label_str = f"{target_day} {s_dt.strftime('%d/%m')}"
                    short_str = f"{s_dt.strftime('%d/%m')}"
                    adv_s = s.get("Adversaire") or s.get("adversaire") or s.get("opponent") or ""
                    label_x = formater_libelle_seance_x(
                        date_obj=s_dt,
                        day_name=target_day,
                        adversaire=adv_s,
                        notes=s.get("notes", ""),
                        week_num=w,
                    )

                    g_metrics = calculer_metriques_groupes_seance(df_s, categories=all_cats, isolated_positions=isolated_positions)
                    r_stats_s = s.get("rpe_stats", {}) or {}
                    by_p_s = r_stats_s.get("by_poste") if isinstance(r_stats_s, dict) else {}
                    for grp_key in all_cats:
                        is_iso = grp_key in isolated_set and not is_grand_groupe(grp_key)
                        df_group = _filter_df_by_group(df_s, grp_key, isolated_positions=isolated_positions if not is_iso else None, is_isolated=is_iso)
                        rpe_moyen = None
                        if not df_group.empty and "RPE" in df_group.columns:
                            s_rpe = pd.to_numeric(df_group["RPE"], errors="coerce").dropna()
                            s_rpe = s_rpe[(s_rpe >= 1.0) & (s_rpe <= 10.0)]
                            if not s_rpe.empty:
                                rpe_moyen = round(float(s_rpe.mean()), 1)
                        if rpe_moyen is None and isinstance(r_stats_s, dict):
                            r_val = (
                                r_stats_s.get(grp_key)
                                or (by_p_s.get(grp_key) if isinstance(by_p_s, dict) else None)
                                or r_stats_s.get(f"rpe_moyen_{grp_key.lower().replace(' ', '_').replace('è', 'e')}")
                            )
                            if r_val is not None:
                                try:
                                    f_r = float(r_val)
                                    if f_r > 0:
                                        rpe_moyen = round(f_r, 1)
                                except Exception:
                                    pass
                        if rpe_moyen is None and isinstance(w_data.get("rpe_summary"), dict):
                            r_w_val = w_data["rpe_summary"].get(grp_key)
                            if r_w_val is not None:
                                try:
                                    f_rw = float(r_w_val)
                                    if f_rw > 0:
                                        rpe_moyen = round(f_rw, 1)
                                except Exception:
                                    pass
                        g_metrics[grp_key]["rpe"] = rpe_moyen

                    matched_sessions[date_key] = {
                        "date_key": date_key,
                        "sort_date": s_dt,
                        "label": label_str,
                        "short_label": short_str,
                        "label_x": label_x,
                        "full_date": s_dt.strftime("%d/%m/%Y"),
                        "is_current": False,
                        "num_players": len(df_s),
                        "players_data": p_list,
                        "rpe_stats": r_stats_s,
                        "metrics": calculer_moyennes_seance(df_s),
                        "group_metrics": g_metrics,
                    }

    # Intégration de la séance actuelle
    curr_date_key = dt_curr.strftime("%Y-%m-%d")
    curr_label = f"{target_day} {dt_curr.strftime('%d/%m')}"
    curr_short = f"{dt_curr.strftime('%d/%m')}"
    curr_label_x = formater_libelle_seance_x(
        date_obj=dt_curr,
        day_name=target_day,
        adversaire=current_adversaire,
        notes=current_notes_or_tag,
        week_num=curr_iso_week,
    )

    curr_g_metrics = calculer_metriques_groupes_seance(df_current, categories=all_cats, isolated_positions=isolated_positions)
    curr_by_p = current_rpe_stats.get("by_poste") if isinstance(current_rpe_stats, dict) else {}
    for grp_key in all_cats:
        is_iso = grp_key in isolated_set and not is_grand_groupe(grp_key)
        df_group = _filter_df_by_group(df_current, grp_key, isolated_positions=isolated_positions if not is_iso else None, is_isolated=is_iso)
        rpe_moyen = None
        if not df_group.empty and "RPE" in df_group.columns:
            s_r = pd.to_numeric(df_group["RPE"], errors="coerce").dropna()
            s_r = s_r[(s_r >= 1.0) & (s_r <= 10.0)]
            if not s_r.empty:
                rpe_moyen = float(s_r.mean())
        if (rpe_moyen is None or pd.isna(rpe_moyen)) and current_rpe_stats and isinstance(current_rpe_stats, dict):
            r_val = (
                current_rpe_stats.get(grp_key)
                or (curr_by_p.get(grp_key) if isinstance(curr_by_p, dict) else None)
                or current_rpe_stats.get(f"rpe_moyen_{grp_key.lower().replace(' ', '_').replace('è', 'e')}")
            )
            if r_val is not None:
                try:
                    f_r = float(r_val)
                    if f_r > 0:
                        rpe_moyen = f_r
                except Exception:
                    pass
        curr_g_metrics[grp_key]["rpe"] = round(rpe_moyen, 1) if (rpe_moyen is not None and not pd.isna(rpe_moyen)) else None

    matched_sessions[curr_date_key] = {
        "date_key": curr_date_key,
        "sort_date": dt_curr,
        "label": curr_label,
        "short_label": curr_short,
        "label_x": curr_label_x,
        "full_date": dt_curr.strftime("%d/%m/%Y"),
        "is_current": True,
        "num_players": len(df_current) if df_current is not None else 0,
        "metrics": calculer_moyennes_seance(df_current),
        "group_metrics": curr_g_metrics,
    }

    sorted_sessions = sorted(matched_sessions.values(), key=lambda x: x["sort_date"])
    n_clamped = max(1, min(4, n_max))
    final_sessions = sorted_sessions[-n_clamped:]

    if final_sessions:
        final_sessions[-1]["is_current"] = True
        for it in final_sessions[:-1]:
            it["is_current"] = False

    return final_sessions


def _calculate_rpe_groups_for_week(
    df_week: pd.DataFrame,
    sessions_dict: Optional[Dict[str, Any]] = None,
    week_data: Optional[Dict[str, Any]] = None,
    categories: Optional[List[str]] = None,
    isolated_positions: Optional[List[str]] = None,
) -> Dict[str, Optional[float]]:
    """
    Calcule les moyennes RPE de la semaine pour l'équipe, les 3 groupes de postes
    et éventuellement les 9 postes spécifiques. Jamais de faux 'NR' si des données
    Google Sheets / session / cache existent pour cette semaine.
    """
    cats_extra = [c for c in (categories or CANONICAL_9_POSTES) if c not in ("5 de devant", "3ème Ligne", "Trois-quarts")]
    res: Dict[str, Optional[float]] = {
        "Total": None,
        "5 de devant": None,
        "3ème Ligne": None,
        "Trois-quarts": None,
    }
    for c in cats_extra:
        res[c] = None

    # 1. Collecte directe depuis les séances enregistrées (joueurs + rpe_stats)
    if sessions_dict and rpe_connector:
        try:
            w_rpe = rpe_connector.calculate_weekly_rpe_from_sessions(sessions_dict)
            if w_rpe:
                res["Total"] = w_rpe.get("rpe_moyen_equipe")
                res["5 de devant"] = w_rpe.get("rpe_5_devant")
                res["3ème Ligne"] = w_rpe.get("rpe_3eme_ligne")
                res["Trois-quarts"] = w_rpe.get("rpe_trois_quarts")
        except Exception:
            pass

    # 2. Complément depuis df_week (colonnes RPE individuelles normalisées par groupe / poste)
    if df_week is not None and not df_week.empty and "RPE" in df_week.columns:
        r_all = pd.to_numeric(df_week["RPE"], errors="coerce").dropna()
        if not r_all.empty and res["Total"] is None:
            res["Total"] = round(float(r_all.mean()), 1)
        isolated_set = set(isolated_positions or [])
        for cat_k in ["5 de devant", "3ème Ligne", "Trois-quarts"] + cats_extra:
            if res.get(cat_k) is None:
                is_iso = cat_k in isolated_set and not is_grand_groupe(cat_k)
                df_sub = _filter_df_by_group(df_week, cat_k, isolated_positions=isolated_positions if not is_iso else None, is_isolated=is_iso)
                if not df_sub.empty and "RPE" in df_sub.columns:
                    sub_r = pd.to_numeric(df_sub["RPE"], errors="coerce").dropna()
                    if not sub_r.empty:
                        res[cat_k] = round(float(sub_r.mean()), 1)

    # 3. Complément depuis rpe_stats / by_poste par séance si toujours manquant
    if sessions_dict:
        poste_vals: Dict[str, List[float]] = {c: [] for c in cats_extra}
        for s_info in sessions_dict.values():
            if not isinstance(s_info, dict):
                continue
            r_st = s_info.get("rpe_stats") or {}
            by_p = r_st.get("by_poste") if isinstance(r_st, dict) else None
            if isinstance(by_p, dict):
                for c in cats_extra:
                    if by_p.get(c) is not None:
                        try:
                            poste_vals[c].append(float(by_p[c]))
                        except Exception:
                            pass
        for c, vals in poste_vals.items():
            if res.get(c) is None and vals:
                res[c] = round(float(np.mean(vals)), 1)

    # 4. Complément final depuis week_data["rpe_summary"] (agrégation directe Google Sheets de la semaine ISO)
    if isinstance(week_data, dict):
        r_sum = week_data.get("rpe_summary")
        if isinstance(r_sum, dict):
            for k in list(res.keys()):
                if res[k] is None and r_sum.get(k) is not None:
                    try:
                        res[k] = round(float(r_sum[k]), 1)
                    except Exception:
                        pass

    return res


def recuperer_historique_semaines(
    year_curr: int,
    week_curr: int,
    df_curr_hebdo: pd.DataFrame,
    n_max: int = 4,
    categories: Optional[List[str]] = None,
    synthese_curr: Optional[Dict[str, Any]] = None,
    isolated_positions: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    """
    Extrait l'historique des N dernières semaines enregistrées (semaine en cours incluse).
    Rattrape automatiquement les données RPE manquantes (ex: S43) via Google Sheets
    et prépare `group_metrics` (Volume, HSR, VHSR, RPE) pour le comparatif triptyque.
    """
    target_cats = list(categories) if categories else ["5 de devant", "3ème Ligne", "Trois-quarts"]
    all_cats = list(dict.fromkeys(["5 de devant", "3ème Ligne", "Trois-quarts"] + target_cats))

    saved_weeks = list_saved_weeks()
    saved_weeks_chrono = sorted(saved_weeks, key=lambda x: (x["year"], x["week_num"]))

    # Cache unique des saisies Google Sheets RPE pour rattraper les semaines historiques avec 'NR'
    df_rpe_cache: Optional[pd.DataFrame] = None
    lookup_ref = _get_effective_joueurs_lookup()

    def _get_rpe_sheet_once() -> pd.DataFrame:
        nonlocal df_rpe_cache
        if df_rpe_cache is None:
            if rpe_connector and hasattr(rpe_connector, "fetch_raw_rpe_data"):
                try:
                    df_rpe_cache = rpe_connector.fetch_raw_rpe_data()
                except Exception:
                    df_rpe_cache = pd.DataFrame()
            else:
                df_rpe_cache = pd.DataFrame()
        return df_rpe_cache

    weeks_dict: Dict[str, Dict[str, Any]] = {}
    for w_info in saved_weeks_chrono:
        y = w_info["year"]
        w = w_info["week_num"]
        if (y < year_curr) or (y == year_curr and w <= week_curr):
            w_key = f"{y}-W{w:02d}"
            w_data = load_week_data(y, w)

            # Rattrapage dynamique des RPE historiques manquants (S43, etc.)
            if rpe_connector and hasattr(rpe_connector, "enrich_historical_week_data_with_rpe"):
                try:
                    sessions_w = w_data.get("sessions", {})
                    needs_rpe = (
                        not isinstance(w_data.get("rpe_summary"), dict)
                        or w_data["rpe_summary"].get("Total") is None
                        or any(
                            not isinstance(s.get("rpe_stats"), dict) or s["rpe_stats"].get("equipe") is None
                            for s in sessions_w.values() if isinstance(s, dict)
                        )
                    )
                    if needs_rpe:
                        df_sheet = _get_rpe_sheet_once()
                        if not df_sheet.empty:
                            w_data, modified = rpe_connector.enrich_historical_week_data_with_rpe(
                                w_data, df_rpe=df_sheet, joueurs_lookup=lookup_ref
                            )
                            if modified:
                                try:
                                    w_path = get_week_file_path(y, w)
                                    with open(w_path, "w", encoding="utf-8") as fw:
                                        json.dump(w_data, fw, ensure_ascii=False, indent=2)
                                except Exception:
                                    pass
                except Exception:
                    pass

            df_w, _ = aggregate_weekly_sessions(w_data)
            label_str = f"Semaine {w:02d}"
            short_str = f"S{w:02d}"
            metrics_w = calculer_moyennes_seance(df_w)
            rpe_grp_w = _calculate_rpe_groups_for_week(
                df_w, w_data.get("sessions", {}), week_data=w_data, categories=all_cats, isolated_positions=isolated_positions
            )
            g_metrics_w = calculer_metriques_groupes_seance(df_w, categories=all_cats, isolated_positions=isolated_positions)
            for cat_k in all_cats:
                if g_metrics_w.get(cat_k, {}).get("rpe") is None and rpe_grp_w.get(cat_k) is not None:
                    g_metrics_w.setdefault(cat_k, {"volume": 0.0, "hsr": 0.0, "vhsr": 0.0, "rpe": None, "num_players": 0})
                    g_metrics_w[cat_k]["rpe"] = rpe_grp_w[cat_k]

            r_av = [v for v in [rpe_grp_w.get("5 de devant"), rpe_grp_w.get("3ème Ligne")] if v is not None]
            metrics_w["RPE Moyen (1-10)"] = {
                "Total": rpe_grp_w.get("Total") or 0.0,
                "Avants": round(float(np.mean(r_av)), 1) if r_av else 0.0,
                "Arrières": rpe_grp_w.get("Trois-quarts") or 0.0,
                "5 de devant": rpe_grp_w.get("5 de devant") or 0.0,
                "3ème Ligne": rpe_grp_w.get("3ème Ligne") or 0.0,
                "Trois-quarts": rpe_grp_w.get("Trois-quarts") or 0.0,
            }

            weeks_dict[w_key] = {
                "key": w_key,
                "year": y,
                "week_num": w,
                "label": label_str,
                "short_label": short_str,
                "label_x": short_str,
                "is_current": False,
                "num_players": len(df_w),
                "metrics": metrics_w,
                "rpe_groups": rpe_grp_w,
                "group_metrics": g_metrics_w,
            }

    curr_key = f"{year_curr}-W{week_curr:02d}"
    curr_w_data = load_week_data(year_curr, week_curr)
    if rpe_connector and hasattr(rpe_connector, "enrich_historical_week_data_with_rpe") and curr_w_data.get("sessions"):
        try:
            df_sheet = _get_rpe_sheet_once()
            if not df_sheet.empty:
                curr_w_data, mod_c = rpe_connector.enrich_historical_week_data_with_rpe(
                    curr_w_data, df_rpe=df_sheet, joueurs_lookup=lookup_ref
                )
                if mod_c:
                    try:
                        with open(get_week_file_path(year_curr, week_curr), "w", encoding="utf-8") as fw:
                            json.dump(curr_w_data, fw, ensure_ascii=False, indent=2)
                    except Exception:
                        pass
        except Exception:
            pass

    curr_label = f"Semaine {week_curr:02d}"
    curr_short = f"S{week_curr:02d}"
    curr_metrics = calculer_moyennes_seance(df_curr_hebdo)
    curr_rpe_grp = _calculate_rpe_groups_for_week(
        df_curr_hebdo,
        synthese_curr if isinstance(synthese_curr, dict) else curr_w_data.get("sessions", {}),
        week_data=curr_w_data,
        categories=all_cats,
        isolated_positions=isolated_positions,
    )
    curr_g_metrics = calculer_metriques_groupes_seance(df_curr_hebdo, categories=all_cats, isolated_positions=isolated_positions)
    for cat_k in all_cats:
        if curr_g_metrics.get(cat_k, {}).get("rpe") is None and curr_rpe_grp.get(cat_k) is not None:
            curr_g_metrics.setdefault(cat_k, {"volume": 0.0, "hsr": 0.0, "vhsr": 0.0, "rpe": None, "num_players": 0})
            curr_g_metrics[cat_k]["rpe"] = curr_rpe_grp[cat_k]

    curr_r_av = [v for v in [curr_rpe_grp.get("5 de devant"), curr_rpe_grp.get("3ème Ligne")] if v is not None]
    curr_metrics["RPE Moyen (1-10)"] = {
        "Total": curr_rpe_grp.get("Total") or 0.0,
        "Avants": round(float(np.mean(curr_r_av)), 1) if curr_r_av else 0.0,
        "Arrières": curr_rpe_grp.get("Trois-quarts") or 0.0,
        "5 de devant": curr_rpe_grp.get("5 de devant") or 0.0,
        "3ème Ligne": curr_rpe_grp.get("3ème Ligne") or 0.0,
        "Trois-quarts": curr_rpe_grp.get("Trois-quarts") or 0.0,
    }

    weeks_dict[curr_key] = {
        "key": curr_key,
        "year": year_curr,
        "week_num": week_curr,
        "label": curr_label,
        "short_label": curr_short,
        "label_x": curr_short,
        "is_current": True,
        "num_players": len(df_curr_hebdo) if df_curr_hebdo is not None else 0,
        "metrics": curr_metrics,
        "rpe_groups": curr_rpe_grp,
        "group_metrics": curr_g_metrics,
    }

    sorted_weeks = sorted(weeks_dict.values(), key=lambda x: (x["year"], x["week_num"]))
    n_clamped = max(1, min(8, n_max))
    final_weeks = sorted_weeks[-n_clamped:]

    if final_weeks:
        final_weeks[-1]["is_current"] = True
        for w_it in final_weeks[:-1]:
            w_it["is_current"] = False

    return final_weeks


def generer_graphique_comparatif_historique_matplotlib(
    items_comparatif: List[Dict[str, Any]],
    type_label: str = "Semaine",
) -> io.BytesIO:
    """
    Génère un graphique d'évolution 3 en 1 via Matplotlib :
    - Subplot 1 (Gauche) : Player Load Moyen par date/semaine avec courbes Avants & Arrières.
    - Subplot 2 (Centre) : High Speed Running (> 20 km/h) par date/semaine avec courbes Avants & Arrières.
    - Subplot 3 (Droite) : Évolution du RPE Moyen (1-10) — Fatigue Ressentie déclinée par ligne de poste
                           (5 de devant, 3ème Ligne, Trois-quarts, Équipe).
    Met en valeur la séance ou semaine actuelle (mise en avant dorée/or, badge ★ Actuelle).
    """
    fig, (ax1, ax2, ax3) = plt.subplots(1, 3, figsize=(14.2, 3.2), dpi=180)
    fig.patch.set_facecolor("#FFFFFF")

    labels = [it.get("short_label", it.get("label", "")) for it in items_comparatif]
    x = np.arange(len(labels))

    pl_tot = [it["metrics"]["Player Load"]["Total"] for it in items_comparatif]
    pl_av = [it["metrics"]["Player Load"]["Avants"] for it in items_comparatif]
    pl_ar = [it["metrics"]["Player Load"]["Arrières"] for it in items_comparatif]

    hsr_tot = [it["metrics"]["HSR (> 20 km/h) (m)"]["Total"] for it in items_comparatif]
    hsr_av = [it["metrics"]["HSR (> 20 km/h) (m)"]["Avants"] for it in items_comparatif]
    hsr_ar = [it["metrics"]["HSR (> 20 km/h) (m)"]["Arrières"] for it in items_comparatif]

    colors_pl = ["#F3C300" if it.get("is_current") else "#162B4C" for it in items_comparatif]
    colors_hsr = ["#F3C300" if it.get("is_current") else "#2B6CB0" for it in items_comparatif]

    # --- AX1 : Player Load ---
    ax1.set_facecolor("#F8FAFC")
    bars1 = ax1.bar(x, pl_tot, width=0.42, color=colors_pl, edgecolor="#0F2B48", linewidth=0.8, zorder=3)
    ax1.plot(x, pl_av, color="#E67E22", linestyle="--", marker="o", markersize=4, label="Avants (Pack)", zorder=5)
    ax1.plot(x, pl_ar, color="#162B4C", linestyle=":", marker="s", markersize=4, label="Arrières (3/4)", zorder=5)

    ax1.set_title(f"Évolution Player Load ({type_label}s)", fontsize=9.2, fontweight="bold", color="#162B4C", pad=6)
    ax1.set_xticks(x)
    ax1.set_xticklabels(labels, fontsize=8.2, fontweight="bold", color="#334155")
    max_pl = max(1.0, max(pl_tot or [1.0]), max(pl_av or [1.0]), max(pl_ar or [1.0]))
    ax1.set_ylim(0, max_pl * 1.35)
    ax1.grid(axis="y", linestyle="--", alpha=0.35, color="#94A3B8", zorder=1)

    for i, rect in enumerate(bars1):
        h = rect.get_height()
        it = items_comparatif[i]
        lbl = f"{h:.1f}" if abs(h - round(h)) > 0.05 else f"{int(round(h))}"
        if it.get("is_current"):
            lbl += " ★"
        ax1.annotate(lbl, xy=(rect.get_x() + rect.get_width() / 2, h),
                     xytext=(0, 2), textcoords="offset points", ha="center", va="bottom",
                     fontsize=7.2, fontweight="bold", color="#D97706" if it.get("is_current") else "#1F4E79")

    ax1.legend(loc="upper right", frameon=False, fontsize=6.5)

    # --- AX2 : HSR ---
    ax2.set_facecolor("#F8FAFC")
    bars2 = ax2.bar(x, hsr_tot, width=0.42, color=colors_hsr, edgecolor="#7A0000", linewidth=0.8, zorder=3)
    ax2.plot(x, hsr_av, color="#A30000", linestyle="--", marker="o", markersize=4, label="Avants (Pack)", zorder=5)
    ax2.plot(x, hsr_ar, color="#162B4C", linestyle=":", marker="s", markersize=4, label="Arrières (3/4)", zorder=5)

    ax2.set_title(f"Évolution HSR (> 20 km/h) ({type_label}s)", fontsize=9.2, fontweight="bold", color="#162B4C", pad=6)
    ax2.set_xticks(x)
    ax2.set_xticklabels(labels, fontsize=8.2, fontweight="bold", color="#334155")
    max_hsr = max(10.0, max(hsr_tot or [1.0]), max(hsr_av or [1.0]), max(hsr_ar or [1.0]))
    ax2.set_ylim(0, max_hsr * 1.35)
    ax2.grid(axis="y", linestyle="--", alpha=0.35, color="#94A3B8", zorder=1)

    for i, rect in enumerate(bars2):
        h = rect.get_height()
        it = items_comparatif[i]
        lbl = f"{int(round(h)):,d} m".replace(",", " ")
        if it.get("is_current"):
            lbl += " ★"
        ax2.annotate(lbl, xy=(rect.get_x() + rect.get_width() / 2, h),
                     xytext=(0, 2), textcoords="offset points", ha="center", va="bottom",
                     fontsize=7.2, fontweight="bold", color="#D97706" if it.get("is_current") else "#2B6CB0")

    ax2.legend(loc="upper right", frameon=False, fontsize=6.5)

    # --- AX3 : Évolution RPE Moyen par Groupe (Fatigue Ressentie) ---
    ax3.set_facecolor("#F8FAFC")
    rpe_5dev = [it.get("rpe_groups", {}).get("5 de devant") or it.get("metrics", {}).get("RPE Moyen (1-10)", {}).get("5 de devant") for it in items_comparatif]
    rpe_3eme = [it.get("rpe_groups", {}).get("3ème Ligne") or it.get("metrics", {}).get("RPE Moyen (1-10)", {}).get("3ème Ligne") for it in items_comparatif]
    rpe_3quarts = [it.get("rpe_groups", {}).get("Trois-quarts") or it.get("metrics", {}).get("RPE Moyen (1-10)", {}).get("Trois-quarts") for it in items_comparatif]
    rpe_equipe = [it.get("rpe_groups", {}).get("Total") or it.get("metrics", {}).get("RPE Moyen (1-10)", {}).get("Total") for it in items_comparatif]

    has_any_rpe = any(v is not None and v > 0 for v in (rpe_5dev + rpe_3eme + rpe_3quarts + rpe_equipe))

    if has_any_rpe:
        # Tracé des courbes des 3 groupes de postes
        for line_data, col, marker, lbl in [
            (rpe_5dev, "#E67E22", "o", "5 de devant"),
            (rpe_3eme, "#F3C300", "s", "3ème Ligne"),
            (rpe_3quarts, "#2B6CB0", "^", "Trois-quarts"),
        ]:
            valid_x = [i for i, v in enumerate(line_data) if v is not None and v > 0]
            valid_y = [line_data[i] for i in valid_x]
            if valid_y:
                ax3.plot(valid_x, valid_y, color=col, linewidth=1.8, marker=marker, markersize=4.5, label=lbl, zorder=4)
                for px, py in zip(valid_x, valid_y):
                    ax3.annotate(f"{py:.1f}", xy=(px, py), xytext=(0, 3), textcoords="offset points",
                                 ha="center", va="bottom", fontsize=6.8, fontweight="bold", color=col)

        # Courbe moyenne équipe (repère global en trait pointillé plus épais)
        valid_eq_x = [i for i, v in enumerate(rpe_equipe) if v is not None and v > 0]
        valid_eq_y = [rpe_equipe[i] for i in valid_eq_x]
        if valid_eq_y:
            ax3.plot(valid_eq_x, valid_eq_y, color="#162B4C", linestyle="--", linewidth=2.0, marker="D", markersize=4.0, label="Équipe", zorder=5)

        ax3.legend(loc="upper right", frameon=False, fontsize=6.2)
    else:
        ax3.text(0.5, 0.5, "Aucune saisie RPE disponible", ha="center", va="center", transform=ax3.transAxes,
                 fontsize=8.5, color="#94A3B8", fontstyle="italic")

    ax3.set_title(f"Évolution RPE Moyen par Groupe ({type_label}s)", fontsize=9.2, fontweight="bold", color="#162B4C", pad=6)
    ax3.set_xticks(x)
    ax3.set_xticklabels(labels, fontsize=8.2, fontweight="bold", color="#334155")
    ax3.set_ylim(0, 10.5)
    ax3.set_ylabel("RPE (1-10)", fontsize=7.8, fontweight="bold", color="#162B4C")
    ax3.grid(axis="y", linestyle="--", alpha=0.35, color="#94A3B8", zorder=1)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def generer_graphique_comparatif_triptyque_postes_matplotlib(
    items_comparatif: List[Dict[str, Any]],
    target_day: str = "Vendredi",
    categories_list: Optional[List[str]] = None,
    isolated_positions: Optional[List[str]] = None,
) -> io.BytesIO:
    """
    Génère un triptyque graphique (jusqu'à 3 panneaux côte à côte) officiel Staff RCA :
    - Par défaut : 5 de devant | 3ème Ligne | Trois-quarts (3/4)
    - En mode détaillé ou sur-mesure : tout lot de 1 à 3 postes/lignes (ex: Pilier | Talonneur | 2ème ligne).
    Histogramme à barres groupées :
    - Volume (m) : Bleu (#1E88E5) sur axe Y gauche
    - HSR (> 20 km/h) (m) : Rouge (#C62828) sur axe Y gauche
    - VHSR (> 25 km/h) (m) : Vert (#2ECC71) sur axe Y gauche
    - RPE (1-10) : Orange (#E67E22) sur axe Y droit twinx (échelle 1-10), ou 'NR' gris uniquement si absent.
    Une seule légende horizontale centrée en haut sous le bandeau.
    """
    cats = list(categories_list) if categories_list else ["5 de devant", "3ème Ligne", "Trois-quarts"]
    cats = cats[:3]
    isolated_set = set(isolated_positions or [])
    group_defs = []
    for c in cats:
        is_iso = c in isolated_set and not is_grand_groupe(c)
        title = get_group_display_name(c, isolated_positions=isolated_positions, is_isolated=is_iso)
        group_defs.append((c, title))

    n_panels = max(1, min(3, len(group_defs)))
    fig, axes_arr = plt.subplots(1, n_panels, figsize=(14.5, 6.8), dpi=130)
    axes = [axes_arr] if n_panels == 1 else list(axes_arr)
    fig.patch.set_facecolor("#FFFFFF")

    K = len(items_comparatif)
    x = np.arange(K)
    w = 0.18 if K >= 4 else (0.19 if K == 3 else 0.20)

    for idx, (grp_key, grp_title) in enumerate(group_defs):
        ax = axes[idx]
        ax.set_facecolor("#FAFAFA")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#CBD5E1")
        ax.spines["bottom"].set_color("#CBD5E1")
        ax.grid(axis="y", linestyle="--", alpha=0.2, color="#94A3B8", zorder=1)

        # Axe Y secondaire découplé dédié exclusivement au RPE (1 à 10)
        ax_rpe = ax.twinx()
        ax_rpe.set_facecolor("none")
        ax_rpe.spines["top"].set_visible(False)
        ax_rpe.spines["left"].set_visible(False)
        ax_rpe.spines["bottom"].set_visible(False)
        ax_rpe.spines["right"].set_color("#CBD5E1")
        ax_rpe.grid(False)
        ax_rpe.set_ylim(0.0, 10.5)
        ax_rpe.set_yticks([1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
        ax_rpe.set_ylabel("RPE (1-10)", fontsize=7.8, fontweight="bold", color="#C2410C")
        ax_rpe.tick_params(axis="y", labelsize=7.5, colors="#C2410C")

        v_vol = [it.get("group_metrics", {}).get(grp_key, {}).get("volume", 0.0) for it in items_comparatif]
        v_hsr = [it.get("group_metrics", {}).get(grp_key, {}).get("hsr", 0.0) for it in items_comparatif]
        v_vhsr = [it.get("group_metrics", {}).get(grp_key, {}).get("vhsr", 0.0) for it in items_comparatif]
        v_rpe = [it.get("group_metrics", {}).get(grp_key, {}).get("rpe") for it in items_comparatif]

        # 3 barres de charge externe sur l'axe Y principal (mètres) — Charte exacte RCA / Comparatif
        b_vol = ax.bar(x - 1.5 * w, v_vol, width=w, color="#1E88E5", edgecolor="none", zorder=3)
        b_hsr = ax.bar(x - 0.5 * w, v_hsr, width=w, color="#C62828", edgecolor="none", zorder=3)
        b_vhsr = ax.bar(x + 0.5 * w, v_vhsr, width=w, color="#2ECC71", edgecolor="none", zorder=3)

        # 1 barre de charge interne RPE sur l'axe Y secondaire (1 à 10)
        rpe_bar_heights = [float(v) if (v is not None and pd.notna(v) and float(v) > 0) else 0.0 for v in v_rpe]
        b_rpe = ax_rpe.bar(
            x + 1.5 * w,
            rpe_bar_heights,
            width=w,
            color="#E67E22",
            edgecolor="#C2410C",
            linewidth=0.8,
            alpha=0.95,
            zorder=4,
        )

        # Annotations numériques des métriques GPS (Volume, HSR, VHSR)
        for bars, vals in [(b_vol, v_vol), (b_hsr, v_hsr), (b_vhsr, v_vhsr)]:
            for rect, val in zip(bars, vals):
                if val > 0:
                    h = rect.get_height()
                    ax.annotate(
                        f"{int(round(val))}",
                        xy=(rect.get_x() + rect.get_width() / 2, h),
                        xytext=(0, 3),
                        textcoords="offset points",
                        ha="center",
                        va="bottom",
                        fontsize=7.3,
                        fontweight="bold",
                        color="#1E293B",
                    )

        # Annotations numériques de la note RPE directement au-dessus de sa colonne dédiée
        for i_idx, rpe_val in enumerate(v_rpe):
            rect = b_rpe[i_idx]
            bar_center_x = rect.get_x() + rect.get_width() / 2.0
            if rpe_val is not None and pd.notna(rpe_val) and float(rpe_val) > 0:
                f_val = float(rpe_val)
                lbl_rpe = f"{f_val:.1f}" if abs(f_val - round(f_val)) > 0.05 else f"{int(round(f_val))}"
                ax_rpe.annotate(
                    lbl_rpe,
                    xy=(bar_center_x, f_val),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha="center",
                    va="bottom",
                    fontsize=7.8,
                    fontweight="bold",
                    color="#C2410C",
                    zorder=6,
                )
            else:
                ax_rpe.annotate(
                    "NR",
                    xy=(bar_center_x, 0.2),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha="center",
                    va="bottom",
                    fontsize=6.8,
                    fontweight="bold",
                    color="#94A3B8",
                    zorder=6,
                )

        ax.set_title(grp_title, fontsize=11.5, fontweight="bold", color="#0C2340", pad=12)
        ax.set_xticks(x)
        x_labels = [it.get("label_x", it.get("short_label", it.get("label", ""))) for it in items_comparatif]
        ax.set_xticklabels([str(lbl).replace(" (", "\n(") for lbl in x_labels], fontsize=8.2, fontweight="bold", color="#334155")
        ax.set_xlim(-0.55, max(float(K) - 0.45, 0.55))

        max_v = max(max(v_vol or [100.0]), 100.0)
        ax.set_ylim(0, max_v * 1.25)
        ax.set_ylabel("Mètres (m)", fontsize=7.8, fontweight="bold", color="#0C2340")
        ax.tick_params(axis="x", length=0)
        ax.tick_params(axis="y", labelsize=7.5, color="#94A3B8")

    leg_elements = [
        Line2D([0], [0], marker="o", color="w", markerfacecolor="#1E88E5", markersize=9, label="Volume (m)"),
        Line2D([0], [0], marker="o", color="w", markerfacecolor="#C62828", markersize=9, label="HSR (> 20 km/h) (m)"),
        Line2D([0], [0], marker="o", color="w", markerfacecolor="#2ECC71", markersize=9, label="VHSR (> 25 km/h) (m)"),
        Line2D([0], [0], marker="s", color="w", markerfacecolor="#E67E22", markeredgecolor="#C2410C", markersize=8, label="RPE (1-10)"),
    ]
    fig.legend(
        handles=leg_elements,
        loc="upper center",
        bbox_to_anchor=(0.5, 0.98),
        ncol=4,
        frameon=False,
        fontsize=9.5,
    )

    fig.subplots_adjust(top=0.88, bottom=0.14, left=0.05, right=0.96, wspace=0.32)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def render_comparatif_triptyque_seances_pdf_page(
    pdf: Any,
    items_comparatif: List[Dict[str, Any]],
    target_day: str = "Vendredi",
    week_num: Optional[int] = None,
    df_current: Optional[pd.DataFrame] = None,
    categories_list: Optional[List[str]] = None,
    info_banner_text: Optional[str] = None,
    isolated_positions: Optional[List[str]] = None,
) -> None:
    """
    Rend une page de comparaison historique au format triptyque Staff RCA (3 panneaux côte à côte) :
    - Applicable autant en comparatif journalier qu'en comparatif hebdomadaire.
    - Supporte les 3 groupes ou un lot de 1 à 3 postes spécifiques.
    """
    if not items_comparatif:
        pdf.set_xy(12.0, 30.0)
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(273.0, 10.0, safe_latin1("Aucune donnée comparative disponible."), align="C")
        return

    cats = list(categories_list) if categories_list else ["5 de devant", "3ème Ligne", "Trois-quarts"]
    isolated_set = set(isolated_positions or [])

    # Garantie absolue d'injection du RPE par catégorie
    for it in items_comparatif:
        g_metrics = it.setdefault("group_metrics", {})
        is_curr = it.get("is_current", False)
        for grp in cats:
            grp_dict = g_metrics.setdefault(grp, {})
            if grp_dict.get("rpe") is None or pd.isna(grp_dict.get("rpe")):
                df_src = None
                if is_curr and df_current is not None and not df_current.empty:
                    df_src = df_current
                elif "players_data" in it and it["players_data"]:
                    df_src = pd.DataFrame(it["players_data"])

                if df_src is not None and not df_src.empty:
                    is_iso = grp in isolated_set and not is_grand_groupe(grp)
                    df_group = _filter_df_by_group(df_src, grp, isolated_positions=isolated_positions if not is_iso else None, is_isolated=is_iso)
                    if not df_group.empty and "RPE" in df_group.columns:
                        s_rpe = pd.to_numeric(df_group["RPE"], errors="coerce").dropna()
                        rpe_moyen = s_rpe.mean() if not s_rpe.empty else None
                        grp_dict["rpe"] = round(float(rpe_moyen), 1) if pd.notna(rpe_moyen) else None

    y_top = 19.5
    has_priors = len(items_comparatif) >= 2
    if not has_priors:
        pdf.set_fill_color(254, 243, 199)
        pdf.set_draw_color(245, 158, 11)
        pdf.rect(12.0, y_top, 273.0, 5.5, style="FD")
        pdf.set_xy(12.0, y_top)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.set_text_color(146, 64, 14)
        info_txt = info_banner_text or f"HISTORIQUE : Aucun élément antérieur enregistré pour un {target_day.lower()}. Données actuelles affichées seules."
        pdf.cell(273.0, 5.5, safe_latin1(info_txt), align="C")
        y_chart = y_top + 7.0
        h_chart = 168.0
    else:
        y_chart = y_top + 1.5
        h_chart = 173.0

    try:
        buf_chart = generer_graphique_comparatif_triptyque_postes_matplotlib(
            items_comparatif=items_comparatif,
            target_day=target_day,
            categories_list=cats,
            isolated_positions=isolated_positions,
        )
        pdf.image(buf_chart, x=12.0, y=y_chart, w=273.0, h=h_chart)
    except Exception as e:
        print(f"Erreur graphique triptyque comparatif PDF : {e}")


def render_comparatif_triptyque_pdf_pages(
    pdf: Any,
    items_comparatif: List[Dict[str, Any]],
    subtitle_base: str,
    categories_list: Optional[List[str]] = None,
    target_day: str = "Vendredi",
    week_num: Optional[int] = None,
    df_current: Optional[pd.DataFrame] = None,
    info_banner_text: Optional[str] = None,
    isolated_positions: Optional[List[str]] = None,
) -> None:
    """
    Pagine automatiquement la section comparative en triptyques de 3 panneaux par page
    (1 page en mode Lignes, jusqu'à 3 pages en mode 9 Postes ou Sur-mesure).
    """
    cats = list(categories_list) if categories_list else ["5 de devant", "3ème Ligne", "Trois-quarts"]
    isolated_set = set(isolated_positions or [])
    chunks = [cats[i : i + 3] for i in range(0, len(cats), 3)]
    for ch_idx, chunk in enumerate(chunks):
        if len(chunks) > 1:
            chunk_lbl = " / ".join(get_group_display_name(c, isolated_positions=isolated_positions, is_isolated=(c in isolated_set and not is_grand_groupe(c))) for c in chunk)
            pdf.current_subtitle = f"{subtitle_base} ({ch_idx + 1}/{len(chunks)} — {chunk_lbl})"
        else:
            pdf.current_subtitle = subtitle_base
        pdf.add_page()
        render_comparatif_triptyque_seances_pdf_page(
            pdf=pdf,
            items_comparatif=items_comparatif,
            target_day=target_day,
            week_num=week_num,
            df_current=df_current,
            categories_list=chunk,
            info_banner_text=info_banner_text,
            isolated_positions=isolated_positions,
        )


def render_comparatif_historique_pdf_page(
    pdf: Any,
    items_comparatif: List[Dict[str, Any]],
    type_label: str = "Semaine",
    target_day: str = "",
    categories_list: Optional[List[str]] = None,
    df_current: Optional[pd.DataFrame] = None,
    isolated_positions: Optional[List[str]] = None,
) -> None:
    """
    Rend la page comparative multi-semaines avec le nouveau triptyque visuel unifié
    (Volume bleu, HSR rouge, VHSR vert sur axe gauche + RPE orange sur axe droit).
    """
    info_txt = (
        "HISTORIQUE : Moins de 2 semaines enregistrées dans la mémoire hebdomadaire. Données de la semaine affichées seules."
        if len(items_comparatif) < 2 else None
    )
    render_comparatif_triptyque_seances_pdf_page(
        pdf=pdf,
        items_comparatif=items_comparatif,
        target_day=target_day or type_label,
        df_current=df_current,
        categories_list=categories_list,
        info_banner_text=info_txt,
        isolated_positions=isolated_positions,
    )


def _render_podium_section_pdf(
    pdf: Any,
    podium_data: Dict[str, Dict[str, Any]],
    x_pod: float,
    y_pod: float,
    w_pod: float,
    title_header: str,
    subtitle_map: Optional[Dict[str, str]] = None,
    include_vigilance: bool = True,
) -> None:
    """Rend le bloc podium 3 Premiers (et 3 Derniers Vigilance juste en dessous) de 3 catégories."""
    categories_config = [
        {
            "key": "Top speed (km/h)",
            "title": "TOP SPEED (km/h)",
            "def_sub": "Vitesse Maximale",
            "vigilance_sub": "Vitesse Minimale",
            "color": (22, 43, 76),
            "vigilance_color": (153, 27, 27),
        },
        {
            "key": "% de marche",
            "title": "% DE MARCHE",
            "def_sub": "Moins marche (> 0%)",
            "vigilance_sub": "Plus de marche (%)",
            "color": (43, 108, 176),
            "vigilance_color": (180, 83, 9),
        },
        {
            "key": "Haute intensité (m)",
            "title": "HAUTE INTENSITÉ (m)",
            "def_sub": "Cumul Z3 + Z4 + Z5 > 14.4 km/h",
            "vigilance_sub": "Moins haute intensité",
            "color": (243, 195, 0),
            "vigilance_color": (194, 65, 12),
        },
    ]

    top_badges = {
        1: {"label": "1er", "bg": (243, 195, 0), "fg": (22, 43, 76)},
        2: {"label": "2e", "bg": (100, 116, 139), "fg": (255, 255, 255)},
        3: {"label": "3e", "bg": (180, 83, 9), "fg": (255, 255, 255)},
    }
    vigilance_badges = {
        1: {"label": "Dern.", "bg": (153, 27, 27), "fg": (255, 255, 255)},
        2: {"label": "Av-D.", "bg": (194, 65, 12), "fg": (255, 255, 255)},
        3: {"label": "3e D.", "bg": (217, 119, 6), "fg": (255, 255, 255)},
    }

    def _draw_block(
        y_start_blk: float,
        blk_title: str,
        is_vigilance: bool,
        compact_half: bool,
    ) -> float:
        hdr_h = 5.0 if compact_half else 6.0
        hdr_bg = (139, 30, 36) if is_vigilance else (22, 43, 76)
        pdf.set_fill_color(*hdr_bg)
        pdf.rect(x_pod, y_start_blk, w_pod, hdr_h, style="F")
        pdf.set_fill_color(243, 195, 0)
        pdf.rect(x_pod, y_start_blk + hdr_h - 0.5, w_pod, 0.5, style="F")

        pdf.set_xy(x_pod, y_start_blk)
        pdf.set_font("Helvetica", "B", 8.0 if compact_half else 9.0)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(w_pod, hdr_h - 0.5, safe_latin1(blk_title), align="C")

        card_h = 20.6 if compact_half else 42.0
        card_gap = 1.2 if compact_half else 2.5
        cat_hdr_h = 4.6 if compact_half else 5.8
        row_h = 4.8 if compact_half else 10.0
        row_gap = 0.35 if compact_half else 1.3
        y_card_start = y_start_blk + hdr_h + (1.2 if compact_half else 2.0)

        badges_map = vigilance_badges if is_vigilance else top_badges
        list_field = "vigilance_players" if is_vigilance else "players"

        for i, cat_cfg in enumerate(categories_config):
            y_card = y_card_start + i * (card_h + card_gap)

            pdf.set_fill_color(248, 250, 252)
            pdf.set_draw_color(203, 213, 225)
            pdf.rect(x_pod, y_card, w_pod, card_h, style="FD")

            accent_color = cat_cfg["vigilance_color"] if is_vigilance else cat_cfg["color"]
            pdf.set_fill_color(*accent_color)
            pdf.rect(x_pod, y_card, w_pod, cat_hdr_h, style="F")

            is_gold_banner = (not is_vigilance) and (
                (accent_color == (243, 195, 0)) or ("HAUTE" in str(cat_cfg.get("title", "")).upper())
            )
            if is_gold_banner:
                title_text_color = (22, 43, 76)
                sub_text_color = (22, 43, 76)
                sub_font_style = "B"
            else:
                title_text_color = (255, 255, 255)
                sub_text_color = (226, 232, 240)
                sub_font_style = "I"

            pdf.set_xy(x_pod + 3.0, y_card)
            pdf.set_font("Helvetica", "B", 7.4 if compact_half else 8.5)
            pdf.set_text_color(*title_text_color)
            pdf.cell(62, cat_hdr_h, safe_latin1(cat_cfg["title"]), align="L")

            if is_vigilance:
                sub_txt = cat_cfg["vigilance_sub"]
            else:
                sub_val = subtitle_map.get(cat_cfg["key"]) if isinstance(subtitle_map, dict) else None
                sub_txt = str(sub_val).strip() if sub_val is not None else cat_cfg["def_sub"]
            pdf.set_xy(x_pod + 62, y_card)
            pdf.set_font("Helvetica", sub_font_style, 6.5 if compact_half else 7.5)
            pdf.set_text_color(*sub_text_color)
            pdf.cell(w_pod - 65, cat_hdr_h, safe_latin1(sub_txt), align="R")

            cat_data = podium_data.get(cat_cfg["key"], {}) if isinstance(podium_data, dict) else {}
            players = cat_data.get(list_field, []) if isinstance(cat_data, dict) else []
            total_valid = int(cat_data.get("total_valid", len(players)) or 0) if isinstance(cat_data, dict) else len(players)

            row_w = w_pod - 4
            y_rows_start = y_card + cat_hdr_h + (0.4 if compact_half else 1.2)

            if is_vigilance and len(players) == 0:
                msg = (
                    "Effectif <= 3 : déjà classé dans les 3 Premiers"
                    if total_valid > 0
                    else "Aucune donnée valide pour ce critère"
                )
                pdf.set_xy(x_pod + 3, y_rows_start + (row_h * 0.8))
                pdf.set_font("Helvetica", "I", 7.0)
                pdf.set_text_color(100, 116, 139)
                pdf.cell(w_pod - 6, row_h, safe_latin1(msg), align="C")
                continue

            for rank_idx in range(1, 4):
                y_row = y_rows_start + (rank_idx - 1) * (row_h + row_gap)
                x_row = x_pod + 2

                p_entry = players[rank_idx - 1] if len(players) >= rank_idx and isinstance(players[rank_idx - 1], dict) else {}
                p_name = safe_latin1(p_entry.get("name", "-"))
                p_stat = safe_latin1(p_entry.get("stat", "-"))

                pdf.set_fill_color(255, 255, 255)
                pdf.set_draw_color(226, 232, 240)
                pdf.rect(x_row, y_row, row_w, row_h, style="FD")

                badge_info = badges_map.get(rank_idx, {"label": f"{rank_idx}e", "bg": (100, 116, 139), "fg": (255, 255, 255)})
                photo_path = photo_manager.get_player_photo_path(p_name) if p_name and p_name != "-" else None
                has_photo = bool(photo_path and os.path.isfile(photo_path) and os.path.basename(photo_path) not in ["avatar_par_defaut.jpg", "avatar_par_defaut.png"])

                if compact_half:
                    badge_w = 10.5
                    badge_h = row_h - 1.0
                    pdf.set_fill_color(*badge_info["bg"])
                    pdf.rect(x_row + 1.5, y_row + 0.5, badge_w, badge_h, style="F")
                    pdf.set_xy(x_row + 1.5, y_row + 0.5)
                    pdf.set_font("Helvetica", "B", 6.0)
                    pdf.set_text_color(*badge_info["fg"])
                    pdf.cell(badge_w, badge_h, badge_info["label"], align="C")

                    text_x = x_row + 13.5
                    if has_photo:
                        try:
                            img_s = row_h - 0.8
                            pdf.image(photo_path, x=text_x, y=y_row + 0.4, w=img_s, h=img_s)
                            text_x += img_s + 1.5
                        except Exception:
                            pass

                    stat_w = 34.0
                    name_w = (x_row + row_w - stat_w - 2.0) - text_x
                    name_disp = p_name if len(p_name) <= 26 else p_name[:24] + ".."
                    pdf.set_xy(text_x, y_row)
                    pdf.set_font("Helvetica", "B", 7.0)
                    pdf.set_text_color(30, 41, 59)
                    pdf.cell(max(name_w, 20.0), row_h, safe_latin1(name_disp), align="L")

                    pdf.set_xy(x_row + row_w - stat_w - 1.5, y_row)
                    pdf.set_font("Helvetica", "B", 7.5)
                    pdf.set_text_color(*accent_color)
                    pdf.cell(stat_w, row_h, safe_latin1(p_stat), align="R")
                else:
                    if has_photo:
                        pdf.set_fill_color(*badge_info["bg"])
                        pdf.rect(x_row + 2, y_row + 1.3, 10.0, 7.4, style="F")
                        pdf.set_xy(x_row + 2, y_row + 1.3)
                        pdf.set_font("Helvetica", "B", 7)
                        pdf.set_text_color(*badge_info["fg"])
                        pdf.cell(10.0, 7.4, badge_info["label"], align="C")

                        try:
                            pdf.image(photo_path, x=x_row + 13.5, y=y_row + 1.0, w=8.0, h=8.0)
                            pdf.set_draw_color(203, 213, 225)
                            pdf.rect(x_row + 13.5, y_row + 1.0, 8.0, 8.0, style="D")
                        except Exception:
                            has_photo = False

                        name_disp = p_name if len(p_name) <= 24 else p_name[:22] + ".."
                        pdf.set_xy(x_row + 24.0, y_row)
                        pdf.set_font("Helvetica", "B", 8)
                        pdf.set_text_color(30, 41, 59)
                        pdf.cell(66.0, row_h, safe_latin1(name_disp), align="L")
                    else:
                        pdf.set_fill_color(*badge_info["bg"])
                        pdf.rect(x_row + 2, y_row + 1.3, 12.0, 7.4, style="F")
                        pdf.set_xy(x_row + 2, y_row + 1.3)
                        pdf.set_font("Helvetica", "B", 7.5)
                        pdf.set_text_color(*badge_info["fg"])
                        pdf.cell(12.0, 7.4, badge_info["label"], align="C")

                        name_disp = p_name if len(p_name) <= 28 else p_name[:26] + ".."
                        pdf.set_xy(x_row + 16.5, y_row)
                        pdf.set_font("Helvetica", "B", 8.5)
                        pdf.set_text_color(30, 41, 59)
                        pdf.cell(73.5, row_h, safe_latin1(name_disp), align="L")

                    pdf.set_xy(x_row + 90.0, y_row)
                    pdf.set_font("Helvetica", "B", 9)
                    pdf.set_text_color(*accent_color)
                    pdf.cell(40.0, row_h, safe_latin1(p_stat), align="R")

        return y_card_start + 3 * card_h + 2 * card_gap

    if include_vigilance:
        y_after_top = _draw_block(
            y_start_blk=y_pod,
            blk_title="3 PREMIERS DE LA SÉANCE (TOP 3)",
            is_vigilance=False,
            compact_half=True,
        )
        _draw_block(
            y_start_blk=y_after_top + 2.0,
            blk_title="3 DERNIERS DE LA SÉANCE (VIGILANCE)",
            is_vigilance=True,
            compact_half=True,
        )
    else:
        _draw_block(
            y_start_blk=y_pod,
            blk_title=title_header,
            is_vigilance=False,
            compact_half=False,
        )


def _render_podium_page_aere(
    pdf: Any,
    podium_data: Dict[str, Dict[str, Any]],
    group_name: str = "",
    subtitle_map: Optional[Dict[str, str]] = None,
    top_title: Optional[str] = None,
    vigilance_title: Optional[str] = None,
) -> None:
    """
    Rend une page podium complète pleine largeur (générale, par ligne de poste ou par poste spécifique RCA) :
    - Moitié supérieure : 3 PREMIERS (TOP 3)
    - Moitié inférieure : 3 DERNIERS (VIGILANCE)
    Affiche côte-à-côte les 3 critères (Top Speed, % de Marche, Haute Intensité),
    avec gestion gracieuse des petits groupes (<= 3 ou < 6 joueurs, sans doublons).
    """
    categories_config = [
        {
            "key": "Top speed (km/h)",
            "title": "TOP SPEED (km/h)",
            "def_sub": "Vitesse Maximale",
            "vigilance_sub": "Vitesse Minimale",
            "color": (31, 78, 121),
            "vigilance_color": (153, 27, 27),
        },
        {
            "key": "% de marche",
            "title": "% DE MARCHE",
            "def_sub": "Moins marche (> 0%)",
            "vigilance_sub": "Plus de marche (%)",
            "color": (43, 108, 176),
            "vigilance_color": (180, 83, 9),
        },
        {
            "key": "Haute intensité (m)",
            "title": "HAUTE INTENSITÉ (m)",
            "def_sub": "Cumul Z3 + Z4 + Z5 > 14.4 km/h",
            "vigilance_sub": "Moins haute intensité (m)",
            "color": (243, 195, 0),
            "vigilance_color": (194, 65, 12),
        },
    ]

    top_badges = {
        1: {"label": "1er", "bg": (243, 195, 0), "fg": (22, 43, 76)},
        2: {"label": "2e", "bg": (100, 116, 139), "fg": (255, 255, 255)},
        3: {"label": "3e", "bg": (180, 83, 9), "fg": (255, 255, 255)},
    }
    vigilance_badges = {
        1: {"label": "Dern.", "bg": (153, 27, 27), "fg": (255, 255, 255)},
        2: {"label": "Av-D.", "bg": (194, 65, 12), "fg": (255, 255, 255)},
        3: {"label": "3e D.", "bg": (217, 119, 6), "fg": (255, 255, 255)},
    }

    col_w = 87.0
    col_gap = 6.0
    x_start = 12.0

    def _draw_podium_half(
        y_section: float,
        section_title: str,
        is_vigilance: bool,
    ) -> float:
        # 1. Bandeau de section (3 Premiers vs 3 Derniers Vigilance)
        sec_banner_h = 5.6
        pdf.set_fill_color(*(139, 30, 36) if is_vigilance else (22, 43, 76))
        pdf.rect(x_start, y_section, 273.0, sec_banner_h, style="F")
        pdf.set_fill_color(243, 195, 0)
        pdf.rect(x_start, y_section + sec_banner_h - 0.5, 273.0, 0.5, style="F")

        pdf.set_xy(x_start + 3.0, y_section)
        pdf.set_font("Helvetica", "B", 8.8)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(267.0, sec_banner_h - 0.5, safe_latin1(section_title), align="C")

        y_start = y_section + sec_banner_h + 1.4
        banner_h = 9.0
        card_h = 22.0
        card_gap = 1.2
        badges_map = vigilance_badges if is_vigilance else top_badges
        list_field = "vigilance_players" if is_vigilance else "players"

        for c_idx, cat in enumerate(categories_config):
            x_c = x_start + c_idx * (col_w + col_gap)
            accent_color = cat["vigilance_color"] if is_vigilance else cat["color"]

            # En-tête de catégorie
            pdf.set_fill_color(*accent_color)
            pdf.rect(x_c, y_start, col_w, banner_h, style="F")

            is_gold_banner = (not is_vigilance) and (
                (accent_color == (243, 195, 0)) or ("HAUTE" in str(cat.get("title", "")).upper())
            )
            if is_gold_banner:
                title_text_color = (22, 43, 76)
                sub_text_color = (22, 43, 76)
                sub_font_style = "B"
                lisere_color = (22, 43, 76)
            else:
                title_text_color = (255, 255, 255)
                sub_text_color = (226, 232, 240)
                sub_font_style = "I"
                lisere_color = (243, 195, 0)

            pdf.set_fill_color(*lisere_color)
            pdf.rect(x_c, y_start + banner_h - 0.8, col_w, 0.8, style="F")

            pdf.set_xy(x_c + 2.0, y_start + 0.6)
            pdf.set_font("Helvetica", "B", 9.0)
            pdf.set_text_color(*title_text_color)
            pdf.cell(col_w - 4.0, 4.0, safe_latin1(cat["title"]), align="C")

            if is_vigilance:
                sub_txt = cat["vigilance_sub"]
            else:
                sub_val = subtitle_map.get(cat["key"]) if isinstance(subtitle_map, dict) else None
                sub_txt = str(sub_val).strip() if sub_val is not None else cat["def_sub"]
            pdf.set_xy(x_c + 2.0, y_start + 4.5)
            pdf.set_font("Helvetica", sub_font_style, 7.2)
            pdf.set_text_color(*sub_text_color)
            pdf.cell(col_w - 4.0, 3.6, safe_latin1(sub_txt), align="C")

            cat_data = podium_data.get(cat["key"], {}) if isinstance(podium_data, dict) else {}
            players = cat_data.get(list_field, []) if isinstance(cat_data, dict) else []
            total_valid = int(cat_data.get("total_valid", len(players)) or 0) if isinstance(cat_data, dict) else len(players)

            y_cards_start = y_start + banner_h + 1.2

            if is_vigilance and len(players) == 0:
                pdf.set_fill_color(248, 250, 252)
                pdf.set_draw_color(203, 213, 225)
                pdf.rect(x_c, y_cards_start, col_w, 26.0, style="FD")
                msg_empty = (
                    "Effectif <= 3 joueurs :\nclassement complet affiché dans les 3 Premiers."
                    if total_valid > 0
                    else "Aucune donnée valide enregistrée\npour ce groupe sur cette catégorie."
                )
                pdf.set_xy(x_c + 3.0, y_cards_start + 7.5)
                pdf.set_font("Helvetica", "I", 7.8)
                pdf.set_text_color(100, 116, 139)
                pdf.multi_cell(col_w - 6.0, 4.5, safe_latin1(msg_empty), align="C")
                continue

            max_slots = 3 if not is_vigilance else max(len(players), 0)
            for r_idx in range(1, max_slots + 1):
                y_card = y_cards_start + (r_idx - 1) * (card_h + card_gap)
                p_entry = players[r_idx - 1] if len(players) >= r_idx and isinstance(players[r_idx - 1], dict) else {}
                p_name = safe_latin1(p_entry.get("name", "-"))
                p_stat = safe_latin1(p_entry.get("stat", "-"))

                pdf.set_fill_color(255, 255, 255)
                pdf.set_draw_color(203, 213, 225)
                pdf.rect(x_c, y_card, col_w, card_h, style="FD")

                # Liseré gauche
                pdf.set_fill_color(*accent_color)
                pdf.rect(x_c, y_card, 2.2, card_h, style="F")

                # Photo à gauche (17x17 mm)
                photo_size = 17.0
                photo_x = x_c + 4.0
                photo_y = y_card + (card_h - photo_size) / 2.0
                photo_path = photo_manager.get_player_photo_path(p_name) if p_name and p_name != "-" else None
                has_photo = bool(photo_path and os.path.isfile(photo_path) and os.path.basename(photo_path) not in ["avatar_par_defaut.jpg", "avatar_par_defaut.png"])

                if has_photo:
                    try:
                        pdf.image(photo_path, x=photo_x, y=photo_y, w=photo_size, h=photo_size)
                        pdf.set_draw_color(203, 213, 225)
                        pdf.rect(photo_x, photo_y, photo_size, photo_size, style="D")
                    except Exception:
                        has_photo = False

                if not has_photo:
                    pdf.set_fill_color(241, 245, 249)
                    pdf.set_draw_color(203, 213, 225)
                    pdf.rect(photo_x, photo_y, photo_size, photo_size, style="FD")
                    pdf.set_xy(photo_x, photo_y)
                    pdf.set_font("Helvetica", "B", 9.0)
                    pdf.set_text_color(148, 163, 184)
                    initials = "".join([part[0].upper() for part in p_name.split() if part])[:2] if (p_name and p_name != "-") else "?"
                    pdf.cell(photo_size, photo_size, safe_latin1(initials), align="C")

                # Zone droite : Badge + Statistique en haut, Nom du joueur en bas
                x_right = photo_x + photo_size + 2.5
                w_right = (x_c + col_w) - x_right - 2.0

                badge = badges_map.get(r_idx, {"label": f"{r_idx}e", "bg": (100, 116, 139), "fg": (255, 255, 255)})
                pdf.set_fill_color(*badge["bg"])
                pdf.rect(x_right, y_card + 2.2, 13.5, 6.0, style="F")
                pdf.set_xy(x_right, y_card + 2.2)
                pdf.set_font("Helvetica", "B", 7.8)
                pdf.set_text_color(*badge["fg"])
                pdf.cell(13.5, 6.0, badge["label"], align="C")

                pdf.set_xy(x_right + 14.5, y_card + 2.2)
                pdf.set_font("Helvetica", "B", 9.5)
                pdf.set_text_color(*accent_color)
                pdf.cell(w_right - 14.5, 6.0, safe_latin1(p_stat), align="R")

                name_disp = p_name if len(p_name) <= 23 else p_name[:21] + ".."
                pdf.set_xy(x_right, y_card + 11.0)
                pdf.set_font("Helvetica", "B", 9.0)
                pdf.set_text_color(30, 41, 59)
                pdf.cell(w_right, 7.5, safe_latin1(name_disp), align="L")

        return y_start + banner_h + 1.2 + 3 * card_h + 2 * card_gap

    grp_label = group_name.upper() if group_name else ""
    t_title = top_title or (f"3 PREMIERS (TOP 3) — {grp_label}" if grp_label else "3 PREMIERS (TOP 3)")
    v_title = vigilance_title or (f"3 DERNIERS (VIGILANCE) — {grp_label}" if grp_label else "3 DERNIERS (VIGILANCE)")

    y_end_top = _draw_podium_half(
        y_section=19.5,
        section_title=t_title,
        is_vigilance=False,
    )

    # Liseré doré séparateur entre 3 Premiers et 3 Derniers Vigilance
    y_sep = y_end_top + 1.4
    pdf.set_fill_color(243, 195, 0)
    pdf.rect(x_start, y_sep, 273.0, 0.5, style="F")

    _draw_podium_half(
        y_section=y_sep + 1.5,
        section_title=v_title,
        is_vigilance=True,
    )


def _render_podiums_compact_horizontal(
    pdf: Any,
    podium_data: Dict[str, Dict[str, Any]],
    y_start: float,
    col_w: float = 87.5,
    col_gap: float = 5.25,
    subtitle_map: Optional[Dict[str, str]] = None,
    is_large: bool = False,
    x_start: float = 12.0,
    mode: str = "players",
    row_h_override: Optional[float] = None,
) -> float:
    """
    Rend un bloc podium horizontal élégant de 3 catégories côte-à-côte :
    - Top Speed (km/h)
    - % de marche
    - Haute intensité (m)

    Supporte `mode="players"` (3 Premiers) et `mode="vigilance_players"` (3 Derniers Vigilance).
    Gère gracieusement les petits effectifs (<= 3 joueurs ou < 6 joueurs) sans doublons ni exception.
    Retourne la hauteur totale occupée (mm).
    """
    is_vigilance = (mode == "vigilance_players")
    categories_config = [
        {
            "key": "Top speed (km/h)",
            "title": "TOP SPEED (km/h)",
            "def_sub": "Vitesse Maximale",
            "vigilance_sub": "Vitesse Minimale",
            "color": (153, 27, 27) if is_vigilance else (31, 78, 121),
        },
        {
            "key": "% de marche",
            "title": "% DE MARCHE",
            "def_sub": "Moins marche (> 0%)",
            "vigilance_sub": "Plus de marche (%)",
            "color": (180, 83, 9) if is_vigilance else (46, 90, 136),
        },
        {
            "key": "Haute intensité (m)",
            "title": "HAUTE INTENSITÉ (m)",
            "def_sub": "Cumul Z3+Z4+Z5 > 14.4 km/h",
            "vigilance_sub": "Moins haute intensité",
            "color": (194, 65, 12) if is_vigilance else (163, 0, 0),
        },
    ]

    top_badges = {
        1: {"label": "1er", "bg": (217, 119, 6), "fg": (255, 255, 255)},
        2: {"label": "2e", "bg": (100, 116, 139), "fg": (255, 255, 255)},
        3: {"label": "3e", "bg": (180, 83, 9), "fg": (255, 255, 255)},
    }
    vigilance_badges = {
        1: {"label": "Dern.", "bg": (153, 27, 27), "fg": (255, 255, 255)},
        2: {"label": "Av-D.", "bg": (194, 65, 12), "fg": (255, 255, 255)},
        3: {"label": "3e D.", "bg": (217, 119, 6), "fg": (255, 255, 255)},
    }
    rank_badges = vigilance_badges if is_vigilance else top_badges

    if row_h_override is not None:
        banner_h = 4.8
        row_h = float(row_h_override)
        row_gap = 0.8
        photo_size = min(9.5, max(7.0, row_h - 2.2))
        title_font = 7.4
        sub_font = 6.0
    elif is_large:
        banner_h = 5.8
        row_h = 20.0
        row_gap = 1.6
        photo_size = 14.0
        title_font = 8.5
        sub_font = 7.0
    else:
        banner_h = 5.0
        row_h = 10.5
        row_gap = 1.0
        photo_size = 8.4
        title_font = 7.5
        sub_font = 6.2

    y_rows_start = y_start + banner_h + (1.2 if is_large else 0.9)

    for c_idx, cat in enumerate(categories_config):
        x_c = x_start + c_idx * (col_w + col_gap)
        accent_color = cat["color"]

        # 1. En-tête de catégorie
        pdf.set_fill_color(*accent_color)
        pdf.rect(x_c, y_start, col_w, banner_h, style="F")

        is_gold_banner = (not is_vigilance) and (
            (accent_color == (243, 195, 0)) or ("HAUTE" in str(cat.get("title", "")).upper())
        )
        if is_gold_banner:
            title_text_color = (22, 43, 76)
            sub_text_color = (22, 43, 76)
            sub_font_style = "B"
            lisere_color = (22, 43, 76)
        else:
            title_text_color = (255, 255, 255)
            sub_text_color = (241, 245, 249)
            sub_font_style = "I"
            lisere_color = (243, 195, 0)

        pdf.set_fill_color(*lisere_color)
        pdf.rect(x_c, y_start + banner_h - 0.5, col_w, 0.5, style="F")

        pdf.set_xy(x_c + 2.0, y_start)
        pdf.set_font("Helvetica", "B", title_font)
        pdf.set_text_color(*title_text_color)
        pdf.cell(col_w * 0.54, banner_h - 0.5, safe_latin1(cat["title"]), align="L")

        if is_vigilance:
            sub_txt = cat["vigilance_sub"]
        else:
            sub_val = subtitle_map.get(cat["key"]) if isinstance(subtitle_map, dict) else None
            sub_txt = str(sub_val).strip() if sub_val is not None else cat["def_sub"]
        pdf.set_xy(x_c + col_w * 0.54, y_start)
        pdf.set_font("Helvetica", sub_font_style, sub_font)
        pdf.set_text_color(*sub_text_color)
        pdf.cell(col_w * 0.46 - 2.0, banner_h - 0.5, safe_latin1(sub_txt), align="R")

        # 2. Lignes joueurs
        cat_data = podium_data.get(cat["key"], {}) if isinstance(podium_data, dict) else {}
        players = cat_data.get(mode, []) if isinstance(cat_data, dict) else []
        total_valid = int(cat_data.get("total_valid", len(players)) or 0) if isinstance(cat_data, dict) else len(players)

        if is_vigilance and len(players) == 0:
            box_h = 3 * row_h + 2 * row_gap
            pdf.set_fill_color(248, 250, 252)
            pdf.set_draw_color(226, 232, 240)
            pdf.rect(x_c, y_rows_start, col_w, box_h, style="FD")
            msg_empty = (
                "Effectif <= 3 : déjà classé dans les 3 Premiers."
                if total_valid > 0
                else "Aucune donnée valide."
            )
            pdf.set_xy(x_c + 2.0, y_rows_start + box_h / 2.0 - 3.0)
            pdf.set_font("Helvetica", "I", 7.0)
            pdf.set_text_color(100, 116, 139)
            pdf.cell(col_w - 4.0, 6.0, safe_latin1(msg_empty), align="C")
            continue

        for r_idx in range(1, 4):
            y_row = y_rows_start + (r_idx - 1) * (row_h + row_gap)
            p_entry = players[r_idx - 1] if len(players) >= r_idx and isinstance(players[r_idx - 1], dict) else {}
            p_name = safe_latin1(p_entry.get("name", "-")).strip()
            p_poste = safe_latin1(p_entry.get("poste", "")).strip()
            p_stat = safe_latin1(p_entry.get("stat", "-")).strip()

            pdf.set_fill_color(255, 255, 255)
            pdf.set_draw_color(226, 232, 240)
            pdf.rect(x_c, y_row, col_w, row_h, style="FD")

            accent_w = 2.0 if is_large else 1.5
            pdf.set_fill_color(*accent_color)
            pdf.rect(x_c, y_row, accent_w, row_h, style="F")

            badge = rank_badges.get(r_idx, {"label": f"{r_idx}e", "bg": (100, 116, 139), "fg": (255, 255, 255)})
            b_w = 10.0 if is_large else (9.5 if is_vigilance else 7.8)
            b_h = min(8.0 if is_large else 7.5, row_h - 2.0)
            b_x = x_c + accent_w + 1.2
            b_y = y_row + (row_h - b_h) / 2.0
            pdf.set_fill_color(*badge["bg"])
            pdf.rect(b_x, b_y, b_w, b_h, style="F")
            pdf.set_xy(b_x, b_y)
            pdf.set_font("Helvetica", "B", 6.8 if is_large else (5.6 if is_vigilance else 6.0))
            pdf.set_text_color(*badge["fg"])
            pdf.cell(b_w, b_h, badge["label"], align="C")

            photo_x = b_x + b_w + 1.4
            photo_y = y_row + (row_h - photo_size) / 2.0
            photo_path = photo_manager.get_player_photo_path(p_name) if p_name and p_name != "-" else None
            has_photo = bool(photo_path and os.path.isfile(photo_path) and os.path.basename(photo_path) not in ["avatar_par_defaut.jpg", "avatar_par_defaut.png"])

            if has_photo:
                try:
                    pdf.image(photo_path, x=photo_x, y=photo_y, w=photo_size, h=photo_size)
                    pdf.set_draw_color(203, 213, 225)
                    pdf.rect(photo_x, photo_y, photo_size, photo_size, style="D")
                except Exception:
                    has_photo = False

            if not has_photo:
                pdf.set_fill_color(241, 245, 249)
                pdf.set_draw_color(203, 213, 225)
                pdf.rect(photo_x, photo_y, photo_size, photo_size, style="FD")
                pdf.set_xy(photo_x, photo_y)
                pdf.set_font("Helvetica", "B", 7.0 if is_large else 5.5)
                pdf.set_text_color(148, 163, 184)
                initials = "".join([part[0].upper() for part in p_name.split() if part])[:2] if (p_name and p_name != "-") else "?"
                pdf.cell(photo_size, photo_size, safe_latin1(initials), align="C")

            if is_large:
                x_text = photo_x + photo_size + 2.5
                w_text = (x_c + col_w) - x_text - 1.5

                disp_name = p_name if len(p_name) <= 22 else p_name[:20] + ".."
                pdf.set_xy(x_text, y_row + 1.8)
                pdf.set_font("Helvetica", "B", 7.8)
                pdf.set_text_color(30, 41, 59)
                pdf.cell(w_text, 4.8, safe_latin1(disp_name), align="L")

                disp_poste = p_poste if p_poste else "RCA"
                if len(disp_poste) > 18:
                    disp_poste = disp_poste[:16] + ".."
                pdf.set_xy(x_text, y_row + 6.8)
                pdf.set_font("Helvetica", "", 6.8)
                pdf.set_text_color(100, 116, 139)
                pdf.cell(w_text, 4.2, safe_latin1(disp_poste), align="L")

                pdf.set_xy(x_text, y_row + 11.8)
                pdf.set_font("Helvetica", "B", 9.5)
                pdf.set_text_color(*accent_color)
                pdf.cell(w_text, 5.5, safe_latin1(p_stat), align="L")
            else:
                x_text = photo_x + photo_size + 1.6
                x_stat = x_c + col_w - 28.5
                w_name = x_stat - x_text - 1.0

                disp_name = p_name if len(p_name) <= 18 else p_name[:16] + ".."
                pdf.set_xy(x_text, y_row + (1.4 if row_h >= 11.5 else 1.0))
                pdf.set_font("Helvetica", "B", 7.0 if row_h >= 11.5 else 6.8)
                pdf.set_text_color(30, 41, 59)
                pdf.cell(w_name, 4.2, safe_latin1(disp_name), align="L")

                disp_poste = p_poste if p_poste else "RCA"
                if len(disp_poste) > 16:
                    disp_poste = disp_poste[:14] + ".."
                pdf.set_xy(x_text, y_row + (6.0 if row_h >= 11.5 else 5.2))
                pdf.set_font("Helvetica", "", 6.0 if row_h >= 11.5 else 5.8)
                pdf.set_text_color(100, 116, 139)
                pdf.cell(w_name, 3.8, safe_latin1(disp_poste), align="L")

                pdf.set_xy(x_stat, y_row)
                pdf.set_font("Helvetica", "B", 7.8)
                pdf.set_text_color(*accent_color)
                pdf.cell(27.0, row_h, safe_latin1(p_stat), align="R")

    total_h = banner_h + (1.2 if is_large else 0.9) + 3 * row_h + 2 * row_gap
    return total_h


class RCAWeeklyPDFReport(FPDF):
    """Classe FPDF Landscape dédiée au rapport hebdomadaire du RCA."""
    def __init__(self, week_label: str, dates_label: str, total_pages: int = 4):
        super().__init__(orientation="L", unit="mm", format="A4")
        self.week_label = week_label
        self.dates_label = dates_label
        self.current_subtitle = ""
        self.total_pages = total_pages
        self.set_auto_page_break(auto=False)
        self.set_margins(left=12, top=20, right=12)
        self.alias_nb_pages()

    def header(self):
        # 1. Bandeau officiel RCA (#162B4C Bleu marine officiel)
        self.set_fill_color(22, 43, 76)
        self.rect(0, 0, 297, 16.5, "F")

        # 2. Liseré doré officiel (#F3C300) de 3 à 4 pt (~1.3 mm)
        self.set_fill_color(243, 195, 0)
        self.rect(0, 16.5, 297, 1.3, "F")

        # 3. Logo officiel du RCA avec respect strict de l'aspect ratio (w=0, h=13.0)
        candidates = [
            os.path.join(PROJECT_ROOT, "assets", "logo_rca.png"),
            os.path.join(PROJECT_ROOT, "logo_rca.png"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo_rca.png"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_rca.png"),
            "assets/logo_rca.png",
            "logo_rca.png",
        ]
        logo_path = None
        for c in candidates:
            if os.path.exists(c):
                logo_path = c
                break

        x_title = 12
        if logo_path and os.path.exists(logo_path):
            try:
                self.image(logo_path, x=12, y=1.75, w=0, h=13.0)
                x_title = 26
            except Exception:
                pass

        # 4. Titre officiel et sous-titre dynamique en doré
        if self.current_subtitle:
            self.set_xy(x_title, 2.2)
            self.set_font("Helvetica", "B", 10.0)
            self.set_text_color(255, 255, 255)
            self.cell(150, 6, safe_latin1("RUGBY CLUB AMIÉNOIS | Rapport Hebdomadaire de Charge GPS"), align="L")

            self.set_xy(x_title, 8.2)
            self.set_font("Helvetica", "B", 9.0)
            self.set_text_color(243, 195, 0)
            self.cell(150, 6, safe_latin1(self.current_subtitle), align="L")
        else:
            self.set_xy(x_title, 3.2)
            self.set_font("Helvetica", "B", 12.0)
            self.set_text_color(255, 255, 255)
            self.cell(150, 10, safe_latin1("RUGBY CLUB AMIÉNOIS | Rapport Hebdomadaire de Charge GPS"), align="L")

        # 5. Semaine et dates
        self.set_font("Helvetica", "B", 9.5)
        self.set_text_color(243, 195, 0)
        self.set_xy(175, 3.2)
        dates_disp = f" - {self.dates_label}" if self.dates_label else ""
        self.cell(110, 10, safe_latin1(f"{self.week_label.upper()}{dates_disp}"), align="R")

    def footer(self):
        # Liseré discret jaune or
        self.set_fill_color(243, 195, 0)
        self.rect(12, 199.5, 273, 0.4, "F")

        self.set_y(-9.5)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(100, 116, 139)
        self.cell(160, 6, safe_latin1("Document Staff Technique RCA — Pôle Préparation Physique"), align="L")
        self.cell(113, 6, safe_latin1(f"Page {self.page_no()}/{{nb}}"), align="R")


def generate_weekly_pdf_report(
    df_hebdo: Optional[pd.DataFrame] = None,
    year: int = 2026,
    week_num: int = 1,
    synthese_jours: Optional[Union[Dict[str, Any], List[Dict[str, Any]]]] = None,
    notes_coach: str = "",
    objectifs_config: Optional[Dict[str, Dict[str, float]]] = None,
    inclure_comparatif: bool = False,
    nb_comparatif: int = 4,
    granularity_mode: str = "Par ligne de poste",
    target_groups: Optional[List[str]] = None,
    isolated_positions: Optional[List[str]] = None,
    **kwargs: Any,
) -> bytes:
    """
    Génère le rapport hebdomadaire PDF complet officiel RCA au format Paysage (A4).
    - Mode 'Par ligne de poste' (9 pages avec comparatif, 8 sans) :
      Page 1 (Synthèse Générale du Club), Page 2 (Podiums Généraux de la Semaine),
      puis 2 pages par ligne (5 de devant, 3ème Ligne, Trois-quarts),
      puis Page 9 Comparative Multi-Semaines.
    - Mode 'Détaillé par poste' (9 postes actifs) :
      Page 1 + Page 2 (Podiums Généraux) + 2 pages par poste actif + pages comparatives triptyques (3 postes/page).
    - Mode 'Sur mesure / Personnalisé (Mixte)' :
      Lignes classiques amputées des postes isolés + pages dédiées pour chaque poste isolé.
      Règle anti-vide : omission totale des lignes vides sans joueur.
    """
    if df_hebdo is None:
        df_hebdo = kwargs.get("df_agg", pd.DataFrame())
    if df_hebdo is not None and isinstance(df_hebdo, pd.DataFrame) and not df_hebdo.empty:
        df_hebdo = df_hebdo.copy()
        if gsheet_sync and hasattr(gsheet_sync, "filter_out_test_players"):
            df_hebdo = gsheet_sync.filter_out_test_players(df_hebdo)
        j_lookup = _get_effective_joueurs_lookup()
        name_col = "Nom ; Prénom" if "Nom ; Prénom" in df_hebdo.columns else ("Nom" if "Nom" in df_hebdo.columns else None)
        pos_col = "Poste_Detaille" if "Poste_Detaille" in df_hebdo.columns else ("Poste" if "Poste" in df_hebdo.columns else None)
        df_hebdo["Poste_Detaille"] = [
            normaliser_poste_specifique(
                r.get(pos_col, "") if pos_col else "",
                player_name=str(r.get(name_col, "")) if name_col else "",
                joueurs_lookup=j_lookup,
            ) or (str(r.get(pos_col, "")).strip() if pos_col else "Non spécifié")
            for _, r in df_hebdo.iterrows()
        ]

    if synthese_jours is None:
        synthese_jours = kwargs.get("selected_sessions_meta", {})

    if isinstance(synthese_jours, list):
        synthese_dict = {f"s_{i}": s for i, s in enumerate(synthese_jours) if isinstance(s, dict)}
    elif isinstance(synthese_jours, dict):
        synthese_dict = dict(synthese_jours)
    else:
        synthese_dict = {}

    granularity_mode = str(kwargs.get("granularity_mode", granularity_mode) or "Par ligne de poste")
    g_low = _strip_accents_lower(granularity_mode)
    is_detailed = ("detaille" in g_low) or ("9 poste" in g_low)

    target_groups_arg = kwargs.get("target_groups", target_groups)
    isolated_arg = kwargs.get("isolated_positions", isolated_positions)
    if target_groups_arg:
        categories = list(target_groups_arg)
        is_detailed = any(not is_grand_groupe(c) for c in categories)
    elif is_detailed:
        categories = get_active_positions(df_hebdo)
        if not categories:
            categories = ["5 de devant", "3ème Ligne", "Trois-quarts"]
    else:
        categories = ["5 de devant", "3ème Ligne", "Trois-quarts"]

    # Construction dynamique des éléments à afficher avec exclusion et règle anti-vide
    items_to_render = []
    isolated_set = set(isolated_arg or [])
    for idx, grp in enumerate(categories):
        is_iso = (idx >= 3 and grp in isolated_set) or (grp in isolated_set and not is_grand_groupe(grp))
        df_sub = _filter_df_by_group(
            df_hebdo, grp, isolated_positions=isolated_arg if not is_iso else None, is_isolated=is_iso
        )
        # Règle anti-vide : si une grande ligne se retrouve vide après exclusion, on saute la page !
        if not is_iso and is_grand_groupe(grp) and df_sub.empty:
            continue
        if is_iso and df_sub.empty and len(categories) > 1:
            continue
        disp_name = get_group_display_name(grp, isolated_positions=isolated_arg, is_isolated=is_iso)
        items_to_render.append({
            "key": grp,
            "display_name": disp_name,
            "is_isolated": is_iso,
            "df_sub": df_sub,
        })

    if not items_to_render:
        items_to_render = [{
            "key": c,
            "display_name": get_group_display_name(c, isolated_positions=isolated_arg),
            "is_isolated": False,
            "df_sub": _filter_df_by_group(df_hebdo, c),
        } for c in categories]

    week_label = f"Semaine {week_num:02d} - {year}"
    raw_dates = [format_readable_date(s.get("date", "")) for s in synthese_dict.values() if isinstance(s, dict) and s.get("date")]
    raw_dates = [d for d in raw_dates if d]
    dates_str = f"Du {raw_dates[0]} au {raw_dates[-1]}" if len(raw_dates) >= 2 else (raw_dates[0] if raw_dates else "")

    n_comp_pages = ((len(items_to_render) + 2) // 3) if inclure_comparatif else 0
    total_pages_calc = 2 + 2 * len(items_to_render) + n_comp_pages
    pdf = RCAWeeklyPDFReport(week_label=week_label, dates_label=dates_str, total_pages=total_pages_calc)

    if df_hebdo is None or df_hebdo.empty:
        pdf.add_page()
        pdf.set_font("Helvetica", "", 12)
        pdf.cell(0, 10, safe_latin1("Aucune donnee disponible pour cette semaine."), align="L")
        return bytes(pdf.output())

    # Chargement + rattrapage RPE Google Sheets pour la semaine en cours (ex: S43)
    w_data_curr = load_week_data(year, week_num)
    if rpe_connector and hasattr(rpe_connector, "enrich_historical_week_data_with_rpe"):
        sessions_curr = w_data_curr.get("sessions", {}) if isinstance(w_data_curr, dict) else {}
        has_any_rpe_curr = ("RPE" in df_hebdo.columns) and pd.to_numeric(df_hebdo["RPE"], errors="coerce").notna().any()
        if not has_any_rpe_curr or not w_data_curr.get("rpe_summary"):
            try:
                w_data_curr, changed_curr = rpe_connector.enrich_historical_week_data_with_rpe(
                    w_data_curr, joueurs_lookup=_get_effective_joueurs_lookup()
                )
                if changed_curr:
                    fpath_curr = get_week_file_path(year, week_num)
                    try:
                        with open(fpath_curr, "w", encoding="utf-8") as fw:
                            json.dump(w_data_curr, fw, indent=2, ensure_ascii=False, default=str)
                    except Exception:
                        pass
            except Exception:
                pass
        # Propager les RPE rattrapés sur df_hebdo si la colonne était absente ou vide
        if isinstance(w_data_curr, dict) and w_data_curr.get("sessions"):
            if not synthese_dict:
                synthese_dict = w_data_curr["sessions"]
            if ("RPE" not in df_hebdo.columns) or pd.to_numeric(df_hebdo["RPE"], errors="coerce").isna().all():
                player_rpe_vals: Dict[str, List[float]] = {}
                for s_info in w_data_curr["sessions"].values():
                    if not isinstance(s_info, dict):
                        continue
                    for p in s_info.get("players_data", []):
                        if not isinstance(p, dict):
                            continue
                        pname = str(p.get("Nom ; Prénom", p.get("Nom", ""))).strip()
                        prpe = p.get("RPE")
                        if pname and prpe is not None and str(prpe).upper() not in ("", "NR", "NONE", "NAN"):
                            try:
                                frpe = float(prpe)
                                if 1.0 <= frpe <= 10.0:
                                    player_rpe_vals.setdefault(pname, []).append(frpe)
                            except Exception:
                                pass
                if player_rpe_vals and "Nom ; Prénom" in df_hebdo.columns:
                    df_hebdo["RPE"] = df_hebdo["Nom ; Prénom"].map(
                        lambda n: round(float(np.mean(player_rpe_vals[str(n).strip()])), 1)
                        if str(n).strip() in player_rpe_vals else np.nan
                    )


    rpe_groups = _calculate_rpe_groups_for_week(
        df_hebdo,
        sessions_dict=w_data_curr.get("sessions") if isinstance(w_data_curr, dict) and w_data_curr.get("sessions") else synthese_dict,
        week_data=w_data_curr,
        categories=[it["key"] for it in items_to_render],
        isolated_positions=isolated_arg,
    )

    # =========================================================================
    # PAGE 1 : SYNTHÈSE GLOBALE DU CLUB & CHARGE HEBDOMADAIRE
    # =========================================================================
    pdf.current_subtitle = "SYNTHÈSE GLOBALE DU CLUB & CHARGE HEBDOMADAIRE"
    pdf.add_page()
    pdf.set_y(19.5)

    tot_km = float(df_hebdo["Distance Totale (km)"].sum()) if not df_hebdo.empty and "Distance Totale (km)" in df_hebdo.columns else 0.0
    tot_hi = float(df_hebdo["Haute intensité (m)"].sum()) if not df_hebdo.empty and "Haute intensité (m)" in df_hebdo.columns else 0.0
    nb_joueurs_tot = len(df_hebdo)
    max_speed = float(df_hebdo["Top Speed (km/h)"].max()) if not df_hebdo.empty and "Top Speed (km/h)" in df_hebdo.columns else 0.0
    nb_seances = len(synthese_dict)

    kpis = [
        ("Volume Cumulé Équipe", f"{tot_km:.1f} km"),
        ("Haute Intensité Cumulée", f"{int(round(tot_hi)):,d} m".replace(",", " ")),
        ("Joueurs Actifs Semaine", f"{nb_joueurs_tot} joueurs"),
        ("Vitesse Max Semaine", f"{max_speed:.1f} km/h"),
        ("Séances Fusionnées", f"{nb_seances} séances"),
    ]
    if rpe_groups.get("Total") is not None:
        kpis.append(("RPE Moyen Semaine", f"{rpe_groups['Total']:.1f}/10"))
    else:
        kpis.append(("RPE Moyen Semaine", "NR"))

    col_w = (273.0 - (len(kpis) - 1) * 2.25) / len(kpis)
    for idx, (label, val) in enumerate(kpis):
        x = 12 + idx * (col_w + 2.25)
        pdf.set_fill_color(248, 250, 252)
        pdf.set_draw_color(226, 232, 240)
        pdf.rect(x, 19.5, col_w, 11.5, "FD")
        pdf.set_xy(x, 20.3)
        pdf.set_font("Helvetica", "B", 6.8)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(col_w, 3.5, safe_latin1(label.upper()), align="C")
        pdf.set_xy(x, 24.0)
        pdf.set_font("Helvetica", "B", 10.0)
        pdf.set_text_color(31, 78, 121)
        pdf.cell(col_w, 6.0, safe_latin1(val), align="C")

    # 1. Encart narratif : Note du Préparateur Physique | Synthèse Athlétique
    y_note = 32.6
    h_note = 10.6
    pdf.set_fill_color(248, 250, 252)
    pdf.set_draw_color(203, 213, 225)
    pdf.rect(12, y_note, 273.0, h_note, style="FD")
    pdf.set_fill_color(22, 43, 76)
    pdf.rect(12, y_note, 3.5, h_note, style="F")
    pdf.set_fill_color(243, 195, 0)
    pdf.rect(15.5, y_note, 0.8, h_note, style="F")

    pdf.set_xy(18, y_note + 1.0)
    pdf.set_font("Helvetica", "B", 7.4)
    pdf.set_text_color(22, 43, 76)
    pdf.cell(150, 3.4, safe_latin1("NOTE DU PRÉPARATEUR PHYSIQUE | SYNTHÈSE ATHLÉTIQUE"), align="L")

    rpe_eq_str = f"{rpe_groups['Total']:.1f} / 10" if rpe_groups.get("Total") is not None else "NR"
    badge_w_txt = f"CHARGE HEBDO : {tot_km:.1f} km  |  RPE MOY. {rpe_eq_str}"
    pdf.set_xy(170, y_note + 0.9)
    pdf.set_fill_color(244, 247, 251)
    pdf.set_draw_color(243, 195, 0)
    pdf.set_font("Helvetica", "B", 6.8)
    pdf.set_text_color(22, 43, 76)
    pdf.cell(112, 3.6, safe_latin1(badge_w_txt), border=1, align="R", fill=True)

    note_txt = str(notes_coach or "").strip() or (
        f"Bilan consolidé sur {nb_seances} séance(s) et {nb_joueurs_tot} joueurs actifs "
        f"({tot_km:.1f} km cumulés, {int(round(tot_hi)):,d} m à haute intensité, Vmax {max_speed:.1f} km/h).".replace(",", " ")
    )
    pdf.set_xy(18, y_note + 4.8)
    pdf.set_font("Helvetica", "", 7.3)
    pdf.set_text_color(51, 65, 85)
    pdf.multi_cell(w=264.0, h=3.4, txt=safe_latin1(note_txt), align="L")

    # 2. Bandeau intermédiaire côte-à-côte (Séances + Volume Journalier) — y=45.0..90.0
    y_mid = 45.0

    pdf.set_xy(12, y_mid)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(134, 3.8, safe_latin1("1. SÉANCES ENREGISTRÉES & VOLUMES :"), align="L")

    y_table = y_mid + 4.2
    pdf.set_xy(12, y_table)
    pdf.set_fill_color(31, 78, 121)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 7.3)
    pdf.cell(46, 4.8, "Séance / Type", 1, 0, "L", fill=True)
    pdf.cell(24, 4.8, "Date", 1, 0, "C", fill=True)
    pdf.cell(20, 4.8, "Effectif", 1, 0, "C", fill=True)
    pdf.cell(22, 4.8, "Dist. (km)", 1, 0, "C", fill=True)
    pdf.cell(22, 4.8, "H.I. (m)", 1, 1, "C", fill=True)

    pdf.set_font("Helvetica", "", 7.1)
    pdf.set_text_color(15, 23, 42)
    sessions_to_show = list(synthese_dict.items())[:6]
    fill_sess = False
    for _, s in sessions_to_show:
        pdf.set_x(12)
        pdf.set_fill_color(*(248, 250, 252) if fill_sess else (255, 255, 255))
        pdf.cell(46, 4.6, safe_latin1(str(s.get("day_type", ""))[:22]), 1, 0, "L", fill=True)
        pdf.cell(24, 4.6, safe_latin1(format_readable_date(s.get("date", ""))), 1, 0, "C", fill=True)
        pdf.cell(20, 4.6, safe_latin1(f"{s.get('num_players', 0)}"), 1, 0, "C", fill=True)
        pdf.cell(22, 4.6, safe_latin1(f"{float(s.get('total_distance_km', 0.0)):.1f} km"), 1, 0, "C", fill=True)
        pdf.cell(22, 4.6, safe_latin1(f"{int(round(float(s.get('total_hi_m', 0.0)))):,d}".replace(",", " ")), 1, 1, "C", fill=True)
        fill_sess = not fill_sess

    pdf.set_xy(151, y_mid)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(134, 3.8, safe_latin1("2. RÉPARTITION DE CHARGE PAR SÉANCE :"), align="L")
    try:
        buf_days = generer_graphique_hebdo_repartition_jours_matplotlib(synthese_dict)
        pdf.image(buf_days, x=151, y=y_mid + 4.2, w=134, h=41.0)
    except Exception as e:
        print(f"Erreur graphique repartition jours PDF : {e}")

    # 3. Graphique principal élargi et centré (Charge Moyenne Hebdomadaire par Ligne de Poste)
    y_main_chart_hdr = 91.5
    pdf.set_xy(12, y_main_chart_hdr)
    pdf.set_font("Helvetica", "B", 8.8)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(273, 4.0, safe_latin1("3. CHARGE MOYENNE HEBDOMADAIRE PAR LIGNE DE POSTE (5 DE DEVANT | 3ÈME LIGNE | TROIS-QUARTS) :"), align="C")

    try:
        buf_groups = generer_graphique_hebdo_groupes_matplotlib(df_hebdo, figsize=(12.2, 4.7))
        chart_w_main = 255.0
        chart_h_main = 98.5
        chart_x_main = 12.0 + (273.0 - chart_w_main) / 2.0
        pdf.image(buf_groups, x=chart_x_main, y=y_main_chart_hdr + 4.5, w=chart_w_main, h=chart_h_main)
    except Exception as e:
        print(f"Erreur graphique groupes hebdo PDF : {e}")

    # =========================================================================
    # PAGE 2 : PODIUMS GÉNÉRAUX DE LA SEMAINE (PLEINE LARGEUR)
    # =========================================================================
    pdf.current_subtitle = "PODIUMS GÉNÉRAUX DE LA SEMAINE"
    pdf.add_page()
    podium_general = generer_podiums_hebdomadaires(df_hebdo)
    sub_map_gen = calculer_sous_titres_podium(df_hebdo, label_prefix="Moy. semaine")
    _render_podium_page_aere(
        pdf=pdf,
        podium_data=podium_general,
        group_name="",
        subtitle_map=sub_map_gen,
        top_title="3 PREMIERS DE LA SEMAINE (TOP 3)",
        vigilance_title="3 DERNIERS DE LA SEMAINE (VIGILANCE)",
    )

    # =========================================================================
    # PAGES DÉTAILLÉES : 2 PAGES PAR CATÉGORIE (RETOUR HEBDO + PODIUMS)
    # =========================================================================
    def _safe_float(v: Any, default: float = 0.0) -> float:
        try:
            if v is None or pd.isna(v):
                return default
            return float(v)
        except (ValueError, TypeError):
            return default

    active_cfg = objectifs_config or charger_objectifs_config()
    multiplier = max(nb_seances, 1)

    for item in items_to_render:
        grp = item["key"]
        disp_name = item["display_name"]
        is_iso = item["is_isolated"]
        df_sub = item["df_sub"]
        n_players_grp = len(df_sub)

        # ---------------------------------------------------------------------
        # PAGE A : RETOUR HEBDOMADAIRE — OBJECTIFS, JAUGES & TABLEAU JOUEURS
        # ---------------------------------------------------------------------
        pdf.current_subtitle = f"RETOUR HEBDOMADAIRE — {disp_name} ({n_players_grp} JOUEUR{'S' if n_players_grp > 1 else ''})"
        pdf.add_page()
        pdf.set_y(19.5)

        targets_daily = get_objectifs_for_category(active_cfg, grp)
        obj_dist = targets_daily["Objectif Distance (m)"] * multiplier
        obj_hsr = targets_daily["Obj. HSR"] * multiplier
        obj_vhsr = targets_daily["Obj. VHSR"] * multiplier
        obj_acc = targets_daily["Obj. ACC/DEC"] * multiplier

        if not df_sub.empty:
            def _extract_metric_mean(df_m: pd.DataFrame, candidates: List[str], mult: float = 1.0) -> float:
                for col in candidates:
                    if col in df_m.columns:
                        s = pd.to_numeric(df_m[col], errors="coerce").dropna()
                        if not s.empty:
                            return float(s.mean() * mult)
                return 0.0

            real_dist = _extract_metric_mean(df_sub, ["Distance Totale (m)", "Distance (m)"])
            if real_dist == 0.0:
                real_dist = _extract_metric_mean(df_sub, ["Distance Totale (km)", "Distance (km)"], mult=1000.0)

            real_hsr = _extract_metric_mean(df_sub, ["HSR Z3 (m)", "Haute intensité (m)", "HSR_m", "Distance haute intensité (m)"])
            if real_hsr == 0.0:
                real_hsr = _extract_metric_mean(df_sub, ["Distance in Speed Zone 3  (km)"], mult=1000.0)

            real_vhsr = _extract_metric_mean(df_sub, ["Distance Sprint VHSR (m)", "VHSR (m)", "Sprint (m)"])
            if real_vhsr == 0.0 and "Distance in Speed Zone 4  (km)" in df_sub.columns and "Distance in Speed Zone 5  (km)" in df_sub.columns:
                s4 = pd.to_numeric(df_sub["Distance in Speed Zone 4  (km)"], errors="coerce").fillna(0.0)
                s5 = pd.to_numeric(df_sub["Distance in Speed Zone 5  (km)"], errors="coerce").fillna(0.0)
                real_vhsr = float((s4 + s5).mean() * 1000.0)

            real_dpm = _extract_metric_mean(df_sub, ["Distance/min (m/min)", "Distance Per Min (m/min)", "m/min"])
            real_acc = _extract_metric_mean(df_sub, ["Accel/Dec total", "Accel/Dec", "ACC/DEC"])
        else:
            real_dist = real_hsr = real_vhsr = real_dpm = real_acc = 0.0

        real_dist = 0.0 if np.isnan(real_dist) else real_dist
        real_hsr = 0.0 if np.isnan(real_hsr) else real_hsr
        real_vhsr = 0.0 if np.isnan(real_vhsr) else real_vhsr
        real_dpm = 0.0 if np.isnan(real_dpm) else real_dpm
        real_acc = 0.0 if np.isnan(real_acc) else real_acc

        pct_dist = (real_dist / obj_dist * 100.0) if obj_dist > 0 else 0.0
        pct_hsr = (real_hsr / obj_hsr * 100.0) if obj_hsr > 0 else 0.0
        pct_vhsr = (real_vhsr / obj_vhsr * 100.0) if obj_vhsr > 0 else 0.0
        pct_acc = (real_acc / obj_acc * 100.0) if obj_acc > 0 else 0.0

        rpe_cat_val = rpe_groups.get(grp)
        if rpe_cat_val is None and not df_sub.empty and "RPE" in df_sub.columns:
            s_rpe_c = pd.to_numeric(df_sub["RPE"], errors="coerce").dropna()
            if not s_rpe_c.empty:
                rpe_cat_val = round(float(s_rpe_c.mean()), 1)
        rpe_cat_str = f"{rpe_cat_val:.1f} / 10" if (rpe_cat_val is not None and rpe_cat_val > 0) else "NR"

        # 1. Tableau Récapitulatif « RETOUR HEBDOMADAIRE — Objectif vs Réalisé »
        table_x = 12.0
        table_y = 19.8
        col_w_grp = 45.0
        col_w_num = 38.0  # 45 + 6*38 = 273 mm

        pdf.set_xy(table_x, table_y)
        pdf.set_fill_color(31, 78, 121)
        pdf.set_draw_color(31, 78, 121)
        pdf.set_font("Helvetica", "B", 7.8)
        pdf.set_text_color(255, 255, 255)
        grp_is_line = not is_iso and is_grand_groupe(grp, is_detailed_context=is_detailed)
        cat_label_prefix = "LIGNE" if grp_is_line else "POSTE"
        headers_tbl = [
            (col_w_grp, f"{cat_label_prefix} : {disp_name}", "L"),
            (col_w_num, "DISTANCE TOTALE", "C"),
            (col_w_num, "DISTANCE HSR (Z3)", "C"),
            (col_w_num, "DISTANCE VHSR (Sprint)", "C"),
            (col_w_num, "m/min", "C"),
            (col_w_num, "ACC / DEC n°", "C"),
            (col_w_num, "RPE MOYEN", "C"),
        ]
        for w, h_title, align in headers_tbl:
            pdf.cell(w, 5.8, safe_latin1(h_title), border=1, align=align, fill=True)

        pdf.set_xy(table_x, table_y + 5.8)
        pdf.set_fill_color(241, 245, 249)
        pdf.set_draw_color(203, 213, 225)
        pdf.set_text_color(71, 85, 105)
        pdf.set_font("Helvetica", "B", 7.6)
        pdf.cell(col_w_grp, 5.4, safe_latin1(f"  Objectif ({multiplier}x)"), border=1, align="L", fill=True)
        pdf.set_font("Helvetica", "", 7.6)
        pdf.set_text_color(31, 78, 121)
        pdf.cell(col_w_num, 5.4, f"{int(round(obj_dist)):,d} m".replace(",", " "), border=1, align="C", fill=True)
        pdf.cell(col_w_num, 5.4, f"{int(round(obj_hsr)):,d} m".replace(",", " "), border=1, align="C", fill=True)
        pdf.cell(col_w_num, 5.4, f"{int(round(obj_vhsr)):,d} m".replace(",", " "), border=1, align="C", fill=True)
        pdf.set_text_color(148, 163, 184)
        pdf.cell(col_w_num, 5.4, "/", border=1, align="C", fill=True)
        pdf.set_text_color(31, 78, 121)
        pdf.cell(col_w_num, 5.4, f"{int(round(obj_acc)):,d}".replace(",", " "), border=1, align="C", fill=True)
        pdf.set_text_color(148, 163, 184)
        pdf.cell(col_w_num, 5.4, "/", border=1, align="C", fill=True)

        pdf.set_xy(table_x, table_y + 11.2)
        pdf.set_fill_color(255, 255, 255)
        pdf.set_draw_color(203, 213, 225)
        pdf.set_text_color(163, 0, 0)
        pdf.set_font("Helvetica", "B", 7.6)
        pdf.cell(col_w_grp, 5.8, safe_latin1("  Réalisé Hebdo"), border=1, align="L", fill=True)
        pdf.set_text_color(30, 41, 59)
        pdf.cell(col_w_num, 5.8, f"{int(round(real_dist)):,d} m ({pct_dist:.0f}%)".replace(",", " "), border=1, align="C", fill=True)
        pdf.cell(col_w_num, 5.8, f"{int(round(real_hsr)):,d} m ({pct_hsr:.0f}%)".replace(",", " "), border=1, align="C", fill=True)
        pdf.cell(col_w_num, 5.8, f"{int(round(real_vhsr)):,d} m ({pct_vhsr:.0f}%)".replace(",", " "), border=1, align="C", fill=True)
        pdf.cell(col_w_num, 5.8, f"{real_dpm:.1f} m/min", border=1, align="C", fill=True)
        pdf.cell(col_w_num, 5.8, f"{int(round(real_acc)):,d} ({pct_acc:.0f}%)".replace(",", " "), border=1, align="C", fill=True)
        pdf.cell(col_w_num, 5.8, safe_latin1(rpe_cat_str), border=1, align="C", fill=True)

        # 2. 4 Graphiques comparatifs 2x2
        try:
            buf_jauges = generer_graphique_hebdo_jauges_matplotlib(
                disp_name, df_sub, nb_seances=nb_seances, objectifs_config=active_cfg, layout="2x2"
            )
            pdf.image(buf_jauges, x=12, y=38.0, w=273, h=60.0)
        except Exception as e:
            print(f"Erreur graphique jauges hebdo PDF pour {disp_name} : {e}")

        # 3. Tableau individuel des statistiques cumulées
        y_sec_tbl = 100.0
        pdf.set_xy(12, y_sec_tbl)
        pdf.set_font("Helvetica", "B", 8.8)
        pdf.set_text_color(11, 25, 44)
        pdf.cell(273, 4.0, safe_latin1(f"STATISTIQUES INDIVIDUELLES CUMULÉES SUR LA SEMAINE — {disp_name} :"), align="L")

        y_tbl_hdr = y_sec_tbl + 4.4
        pdf.set_xy(12, y_tbl_hdr)
        pdf.set_fill_color(31, 78, 121)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 7.5)

        cols_def = [
            ("Joueur", 44, "L"),
            ("Poste", 24, "C"),
            ("Présences", 26, "C"),
            ("Temps (min)", 20, "C"),
            ("Dist. (km)", 23, "C"),
            ("H.I. (m)", 24, "C"),
            ("Sprint (m)", 22, "C"),
            ("Vmax (km/h)", 23, "C"),
            ("m/min", 22, "C"),
            ("Acc/Dec", 23, "C"),
            ("RPE", 22, "C"),
        ]

        for title, w, align in cols_def:
            pdf.cell(w, 5.0, safe_latin1(title), 1, 0, align, fill=True)
        pdf.ln(5.0)

        if df_sub.empty:
            pdf.set_x(12)
            pdf.set_fill_color(248, 250, 252)
            pdf.set_font("Helvetica", "I", 8.0)
            pdf.set_text_color(100, 116, 139)
            pdf.cell(273, 7.0, safe_latin1("Aucun joueur enregistré pour cette catégorie cette semaine."), 1, 1, "C", fill=True)
        else:
            avail_h = 197.0 - (y_tbl_hdr + 5.0)
            row_h = min(5.2, max(3.6, avail_h / max(n_players_grp, 1)))
            font_sz = 7.5 if row_h >= 4.4 else 6.5
            pdf.set_font("Helvetica", "", font_sz)
            pdf.set_text_color(15, 23, 42)
            fill_row = False

            for _, row in df_sub.iterrows():
                pdf.set_x(12)
                bg = (248, 250, 252) if fill_row else (255, 255, 255)
                pdf.set_fill_color(*bg)

                p_nom = str(row.get("Nom ; Prénom", "-"))[:22]
                p_pos = str(row.get("Poste_Detaille", row.get("Poste", "RCA")))[:13]
                nb_s = int(_safe_float(row.get("Nombre de séances", 1), 1))
                p_pres = f"{nb_s} séance{'s' if nb_s > 1 else ''}"
                p_tps = f"{_safe_float(row.get('Temps total (min)', 0.0)):.0f}"
                p_dist = f"{_safe_float(row.get('Distance Totale (km)', 0.0)):.2f}"
                p_hi = f"{int(round(_safe_float(row.get('Haute intensité (m)', 0.0)))):,d}".replace(",", " ")
                p_vhsr = f"{int(round(_safe_float(row.get('Distance Sprint VHSR (m)', 0.0)))):,d}".replace(",", " ")
                p_vmax = f"{_safe_float(row.get('Top Speed (km/h)', 0.0)):.1f}"
                p_dpm = f"{_safe_float(row.get('Distance/min (m/min)', 0.0)):.1f}"
                p_acc = f"{int(round(_safe_float(row.get('Accel/Dec total', 0.0))))}"

                p_rpe_raw = row.get("RPE")
                if p_rpe_raw is not None and not pd.isna(p_rpe_raw) and str(p_rpe_raw).upper() not in ["NR", "NONE", "NAN"]:
                    try:
                        p_rpe = f"{float(p_rpe_raw):.1f}"
                    except Exception:
                        p_rpe = str(p_rpe_raw)[:4]
                else:
                    p_rpe = "NR"

                pdf.cell(44, row_h, safe_latin1(p_nom), 1, 0, "L", fill=True)
                pdf.cell(24, row_h, safe_latin1(p_pos), 1, 0, "C", fill=True)
                pdf.cell(26, row_h, safe_latin1(p_pres), 1, 0, "C", fill=True)
                pdf.cell(20, row_h, safe_latin1(p_tps), 1, 0, "C", fill=True)
                pdf.cell(23, row_h, safe_latin1(p_dist), 1, 0, "C", fill=True)
                pdf.cell(24, row_h, safe_latin1(p_hi), 1, 0, "C", fill=True)
                pdf.cell(22, row_h, safe_latin1(p_vhsr), 1, 0, "C", fill=True)
                pdf.cell(23, row_h, safe_latin1(p_vmax), 1, 0, "C", fill=True)
                pdf.cell(22, row_h, safe_latin1(p_dpm), 1, 0, "C", fill=True)
                pdf.cell(23, row_h, safe_latin1(p_acc), 1, 0, "C", fill=True)
                pdf.cell(22, row_h, safe_latin1(p_rpe), 1, 1, "C", fill=True)
                fill_row = not fill_row

        # ---------------------------------------------------------------------
        # PAGE B : PODIUMS HEBDOMADAIRES — 3 PREMIERS & 3 DERNIERS SUR LA MÊME PAGE
        # ---------------------------------------------------------------------
        pdf.current_subtitle = f"PODIUMS HEBDOMADAIRES — {cat_label_prefix} {disp_name}"
        pdf.add_page()
        podium_grp = generer_podiums_hebdomadaires(df_sub)
        sub_map_grp = calculer_sous_titres_podium(df_sub, label_prefix="Moyenne")
        _render_podium_page_aere(
            pdf=pdf,
            podium_data=podium_grp,
            group_name=disp_name,
            subtitle_map=sub_map_grp,
        )

    # =========================================================================
    # DERNIÈRE(S) PAGE(S) : ANALYSE COMPARATIVE HISTORIQUE (TRIPTYQUE)
    # =========================================================================
    if inclure_comparatif:
        active_comp_categories = [it["key"] for it in items_to_render]
        items_comp = recuperer_historique_semaines(
            year_curr=year,
            week_curr=week_num,
            df_curr_hebdo=df_hebdo,
            n_max=nb_comparatif,
            categories=active_comp_categories,
            synthese_curr=synthese_dict,
            isolated_positions=isolated_arg,
        )
        render_comparatif_triptyque_pdf_pages(
            pdf=pdf,
            items_comparatif=items_comp,
            subtitle_base="ANALYSE COMPARATIVE HISTORIQUE — MULTI-SEMAINES",
            categories_list=active_comp_categories,
            df_current=df_hebdo,
            isolated_positions=isolated_arg,
        )

    return bytes(pdf.output())
