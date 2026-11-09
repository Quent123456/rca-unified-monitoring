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
            import github_sync
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
                    import github_sync
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
                    import github_sync
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
                    import github_sync
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
    for name, group in grouped:
        nb_seances = len(group)
        jours_presence = ", ".join(group["_session_day"].unique().tolist())

        # Groupe et poste de référence (mode / dernier)
        poste_det = group["Poste_Detaille"].iloc[-1] if "Poste_Detaille" in group.columns else "Inconnu"
        groupe_rca = group["Groupe"].iloc[-1] if "Groupe" in group.columns else "Trois-quarts"

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

def generer_podiums_hebdomadaires(
    df_hebdo: pd.DataFrame,
    groupe: Optional[str] = None,
) -> Dict[str, Dict[str, Any]]:
    """
    Génère les podiums de la semaine : Top Speed, % Marche, Haute Intensité.
    Applique les mêmes critères de classement (top 3) et formules de calcul
    que ceux utilisés pour les podiums des rapports journaliers.
    Gère gracieusement le cas où un groupe positionnel compte strictement moins de 3 joueurs.
    """
    def _make_player_entry(rank: int, name: Any, stat_str: Any, poste: Any = "") -> Dict[str, Any]:
        p_name = str(name).strip() if pd.notna(name) and name is not None else "-"
        if p_name in ["nan", "None", ""]:
            p_name = "-"
        p_poste = str(poste).strip() if pd.notna(poste) and poste is not None else ""
        if p_poste in ["nan", "None"]:
            p_poste = ""
        return {
            "rank": rank,
            "name": p_name,
            "poste": p_poste if p_poste else "RCA",
            "stat": str(stat_str) if stat_str is not None else "-",
            "avatar_src": photo_manager.get_player_image_b64(p_name, size=80 if rank == 1 else 65) if p_name != "-" else photo_manager.get_player_image_b64("", size=80 if rank == 1 else 65),
        }

    def _placeholder(rank: int) -> Dict[str, Any]:
        return {
            "rank": rank,
            "name": "-",
            "poste": "",
            "stat": "-",
            "avatar_src": photo_manager.get_player_image_b64("", size=80 if rank == 1 else 65),
        }

    empty_podium = {
        "Top speed (km/h)": {"icon": "⚡", "players": [_placeholder(1), _placeholder(2), _placeholder(3)]},
        "% de marche": {"icon": "🚶", "players": [_placeholder(1), _placeholder(2), _placeholder(3)]},
        "Haute intensité (m)": {"icon": "🔥", "players": [_placeholder(1), _placeholder(2), _placeholder(3)]},
    }

    if df_hebdo is None or not isinstance(df_hebdo, pd.DataFrame) or df_hebdo.empty:
        return empty_podium

    if groupe and str(groupe).strip().upper() not in ["ALL", "EFFECTIF GLOBAL", "TOUS LES JOUEURS", "TOUS LES GROUPES"]:
        if "Groupe" in df_hebdo.columns:
            df_subset = df_hebdo[df_hebdo["Groupe"] == groupe].copy()
        else:
            df_subset = pd.DataFrame()
    else:
        df_subset = df_hebdo.copy()

    if df_subset.empty:
        return empty_podium

    podium_result: Dict[str, Dict[str, Any]] = {}

    # 1. TOP SPEED (km/h) - Vitesse Maximale hebdomadaire (du plus rapide au moins rapide)
    ts_col = "Top Speed (km/h)"
    if ts_col in df_subset.columns:
        s_ts = pd.to_numeric(df_subset[ts_col], errors="coerce").fillna(0.0)
        df_sorted_ts = df_subset.assign(_ts_sort=s_ts).sort_values(by="_ts_sort", ascending=False)
    else:
        df_sorted_ts = df_subset

    players_ts: List[Dict[str, Any]] = []
    for rank, (_, row) in enumerate(df_sorted_ts.head(3).iterrows(), start=1):
        name = str(row.get("Nom ; Prénom", "-")).strip()
        val_raw = row.get(ts_col, 0.0)
        val = float(val_raw) if pd.notna(val_raw) else 0.0
        if np.isnan(val):
            val = 0.0
        poste = str(row.get("Poste_Detaille", row.get("Poste", "")))
        players_ts.append(_make_player_entry(rank, name, f"{val:.1f} km/h", poste))
    while len(players_ts) < 3:
        players_ts.append(_placeholder(len(players_ts) + 1))

    podium_result["Top speed (km/h)"] = {
        "icon": "⚡",
        "players": players_ts,
    }

    # 2. % DE MARCHE - Moins de marche (> 0% : ordre croissant)
    walk_col = "%Marche " if "%Marche " in df_subset.columns else ("% de marche moyen" if "% de marche moyen" in df_subset.columns else None)
    if walk_col and walk_col in df_subset.columns:
        s_walk = pd.to_numeric(df_subset[walk_col], errors="coerce").fillna(0.0)
        df_valid_walk = df_subset[s_walk > 0.0]
        if not df_valid_walk.empty:
            df_sorted_walk = df_valid_walk.sort_values(by=walk_col, ascending=True)
        else:
            df_sorted_walk = df_subset.sort_values(by=walk_col, ascending=True)
    else:
        df_sorted_walk = df_subset

    players_walk: List[Dict[str, Any]] = []
    for rank, (_, row) in enumerate(df_sorted_walk.head(3).iterrows(), start=1):
        name = str(row.get("Nom ; Prénom", "-")).strip()
        val_raw = row.get(walk_col, 0.0) if walk_col else 0.0
        val_float = float(val_raw) if pd.notna(val_raw) else 0.0
        if np.isnan(val_float):
            val_float = 0.0
        val_pct = val_float * 100.0 if val_float <= 1.0 else val_float
        poste = str(row.get("Poste_Detaille", row.get("Poste", "")))
        players_walk.append(_make_player_entry(rank, name, f"{val_pct:.1f} %", poste))
    while len(players_walk) < 3:
        players_walk.append(_placeholder(len(players_walk) + 1))

    podium_result["% de marche"] = {
        "icon": "🚶",
        "players": players_walk,
    }

    # 3. HAUTE INTENSITÉ (m) - Cumul total Zones 3+4+5 (> 14.4 km/h) (ordre décroissant)
    hi_col = None
    for c in ["Haute intensité Podium (m)", " Distance haute intensité (m)", "Haute intensité (m)"]:
        if c in df_subset.columns:
            hi_col = c
            break

    if hi_col and hi_col in df_subset.columns:
        s_hi = pd.to_numeric(df_subset[hi_col], errors="coerce").fillna(0.0)
        df_sorted_hi = df_subset.assign(_hi_sort=s_hi).sort_values(by="_hi_sort", ascending=False)
    else:
        df_sorted_hi = df_subset

    players_hi: List[Dict[str, Any]] = []
    for rank, (_, row) in enumerate(df_sorted_hi.head(3).iterrows(), start=1):
        name = str(row.get("Nom ; Prénom", "-")).strip()
        val_raw = row.get(hi_col, 0.0) if hi_col else 0.0
        val_m = float(val_raw) if pd.notna(val_raw) else 0.0
        if np.isnan(val_m):
            val_m = 0.0
        hi_int = int(round(val_m))
        poste = str(row.get("Poste_Detaille", row.get("Poste", "")))
        players_hi.append(_make_player_entry(rank, name, f"{hi_int:,d} m".replace(",", " "), poste))
    while len(players_hi) < 3:
        players_hi.append(_placeholder(len(players_hi) + 1))

    podium_result["Haute intensité (m)"] = {
        "icon": "🔥",
        "players": players_hi,
    }

    return podium_result


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
            import github_sync
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
                import github_sync
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
                        import github_sync
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


def generer_graphique_hebdo_groupes_matplotlib(df_agg: pd.DataFrame) -> io.BytesIO:
    """
    Génère le diagramme combiné à barres comparatif de la semaine pour la Page 1 du PDF et le Dashboard :
    - Distance Moyenne (m) par Groupe (barres bleu marine #1F4E79)
    - HSR Z3 Moyen (m) par Groupe (barres rouge #A30000)
    sur les 3 groupes officiels du RCA : 5 de devant, 3ème Ligne, Trois-quarts.
    """
    fig, ax1 = plt.subplots(figsize=(6.2, 3.6), dpi=180)
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
    width = 0.35

    rects1 = ax1.bar(x - width/2, dist_vals, width, label="Distance Moyenne (m)", color="#1F4E79", edgecolor="none", zorder=3)
    ax2 = ax1.twinx()
    rects2 = ax2.bar(x + width/2, hsr_vals, width, label="HSR Z3 Moyen (m)", color="#A30000", edgecolor="none", zorder=3)

    ax1.set_ylabel("Distance Moyenne (m)", color="#1F4E79", fontweight="bold", fontsize=8.5)
    ax2.set_ylabel("HSR Z3 Moyen (m)", color="#A30000", fontweight="bold", fontsize=8.5)
    ax1.set_xticks(x)
    ax1.set_xticklabels(groupes, fontweight="bold", fontsize=9.0, color="#1E293B")
    ax1.set_title("Charge Moyenne Hebdomadaire par Ligne", fontweight="bold", color="#1F4E79", fontsize=10.0, pad=8)

    # Annotations au-dessus des barres
    for rect in rects1:
        h = rect.get_height()
        if h > 0:
            ax1.annotate(f"{int(round(h)):,d} m".replace(",", " "),
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 2), textcoords="offset points",
                         ha="center", va="bottom", fontsize=7.5, fontweight="bold", color="#1F4E79")

    for rect in rects2:
        h = rect.get_height()
        if h > 0:
            ax2.annotate(f"{int(round(h)):,d} m".replace(",", " "),
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 2), textcoords="offset points",
                         ha="center", va="bottom", fontsize=7.5, fontweight="bold", color="#A30000")

    max_dist = max(dist_vals) if dist_vals and max(dist_vals) > 0 else 1000.0
    max_hsr = max(hsr_vals) if hsr_vals and max(hsr_vals) > 0 else 100.0
    ax1.set_ylim(0, max_dist * 1.25)
    ax2.set_ylim(0, max_hsr * 1.25)
    ax1.grid(axis="y", linestyle="--", alpha=0.3, zorder=1)
    ax1.tick_params(axis="y", labelcolor="#1F4E79", labelsize=8)
    ax2.tick_params(axis="y", labelcolor="#A30000", labelsize=8)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=2, frameon=False, fontsize=8.0)

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
) -> io.BytesIO:
    """
    Génère une figure 2x2 propre regroupant les 4 duels de barres 'Objectif Hebdomadaire vs Réalisé Cumulé'
    pour une ligne de poste RCA (5 de devant, 3ème Ligne, Trois-quarts) :
    1. Distance Totale (km ou m)
    2. HSR Z3 (m)
    3. VHSR Sprint (m)
    4. ACC / DEC (> 3 m/s²)
    """
    active_cfg = objectifs_config or charger_objectifs_config()
    targets_daily = active_cfg.get(group_name, active_cfg.get("Tous les groupes", DEFAULT_OBJECTIFS_CONFIG["Tous les groupes"]))

    multiplier = max(nb_seances, 1)

    # Calcul des réalisés moyens par joueur du groupe sur la semaine
    if not df_grp.empty:
        try:
            val_dist = float(pd.to_numeric(df_grp.get("Distance Totale (m)", df_grp.get("Distance Totale (km)", 0) * 1000.0), errors="coerce").mean())
        except Exception:
            val_dist = 0.0
        try:
            val_hsr = float(pd.to_numeric(df_grp.get("HSR Z3 (m)", df_grp.get("Haute intensité (m)", 0)), errors="coerce").mean())
        except Exception:
            val_hsr = 0.0
        try:
            val_vhsr = float(pd.to_numeric(df_grp.get("Distance Sprint VHSR (m)", df_grp.get("VHSR (m)", 0)), errors="coerce").mean())
        except Exception:
            val_vhsr = 0.0
        try:
            val_acc = float(pd.to_numeric(df_grp.get("Accel/Dec total", df_grp.get("Accel/Dec", 0)), errors="coerce").mean())
        except Exception:
            val_acc = 0.0
    else:
        val_dist = 0.0
        val_hsr = 0.0
        val_vhsr = 0.0
        val_acc = 0.0

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

    fig, axes = plt.subplots(1, 4, figsize=(11.5, 2.5), dpi=180)
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


def _filter_df_by_group(df: pd.DataFrame, grp: str) -> pd.DataFrame:
    """Filtre un DataFrame de séance par groupe de poste (5 de devant, 3ème Ligne, Trois-quarts)."""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()

    grp_clean = grp.strip().lower()
    if "Groupe" in df.columns:
        if grp_clean in ["5 de devant", "devant"]:
            mask = df["Groupe"].astype(str).str.strip().str.lower().isin(["5 de devant", "devant", "cinq de devant", "avants (5)"])
        elif grp_clean in ["3ème ligne", "3eme ligne", "3e ligne"]:
            mask = df["Groupe"].astype(str).str.strip().str.lower().isin(["3ème ligne", "3eme ligne", "3e ligne", "troisième ligne"])
        elif grp_clean in ["trois-quarts", "trois quarts", "3/4", "arrieres", "arrières"]:
            mask = df["Groupe"].astype(str).str.strip().str.lower().isin(["trois-quarts", "trois quarts", "arrières", "arrieres", "3/4"])
        else:
            mask = df["Groupe"].astype(str).str.strip().str.lower() == grp_clean
        df_sub = df[mask]
        if not df_sub.empty:
            return df_sub

    col_p = "Poste_Detaille" if "Poste_Detaille" in df.columns else ("Poste" if "Poste" in df.columns else None)
    if col_p:
        if grp_clean in ["5 de devant", "devant"]:
            kw = ["pilier", "talonneur", "2ème ligne", "deuxième ligne", "première ligne", "5 de devant"]
        elif grp_clean in ["3ème ligne", "3eme ligne", "3e ligne"]:
            kw = ["3ème ligne", "troisième ligne", "flanker", "n8", "numéro 8"]
        else:
            kw = ["demi", "ouvreur", "mêlée", "centre", "ailier", "arrière", "trois-quarts"]
        mask = df[col_p].astype(str).str.strip().str.lower().apply(lambda s: any(k in s for k in kw))
        return df[mask]

    return pd.DataFrame()


def calculer_metriques_groupes_seance(df_s: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    """
    Calcule les moyennes exactes par ligne de poste pour le triptyque :
    - 5 de devant
    - 3ème Ligne
    - Trois-quarts
    Retourne {groupe: {"volume": m, "hsr": m, "vhsr": m, "rpe": float, "num_players": n}}
    """
    groupes = ["5 de devant", "3ème Ligne", "Trois-quarts"]
    res: Dict[str, Dict[str, Any]] = {
        g: {"volume": 0.0, "hsr": 0.0, "vhsr": 0.0, "rpe": None, "num_players": 0}
        for g in groupes
    }
    if df_s is None or not isinstance(df_s, pd.DataFrame) or df_s.empty:
        return res

    for grp in groupes:
        df_sub = _filter_df_by_group(df_s, grp)
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
) -> List[Dict[str, Any]]:
    """
    Extrait l'historique des N dernières séances enregistrées correspondant au même jour de la semaine
    (ex. si séance du vendredi -> les N derniers vendredis disponibles, séance du jour incluse).
    Gère automatiquement les semaines creuses sans entraînement.
    Ne dépend JAMAIS de la date système du jour.
    """
    dt_curr = parse_session_date(session_date_or_day)
    target_day = detect_day_of_week(session_date_or_day)

    if dt_curr is not None:
        target_weekday = dt_curr.weekday()
        if not target_day:
            JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
            target_day = JOURS[target_weekday]
        curr_iso_year, curr_iso_week, _ = dt_curr.isocalendar()
    else:
        # Si session_date_or_day est un nom de jour textuel (ex: "Vendredi")
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

    for w_info in saved_weeks:
        y = w_info["year"]
        w = w_info["week_num"]
        w_data = load_week_data(y, w)
        for s_id, s in w_data.get("sessions", {}).items():
            if not isinstance(s, dict):
                continue
            s_day_raw = s.get("day_type", "") + " " + s.get("date", "")
            s_day = detect_day_of_week(s_day_raw)
            s_dt = _parse_date_to_obj(s.get("date"), fallback_year=y, fallback_week=w, fallback_day_name=s_day or target_day)
            if s_dt is None:
                continue

            # Filtrage strict sur le même jour de la semaine
            is_same_weekday = (s_dt.weekday() == target_weekday)
            is_same_dayname = bool(s_day and s_day.lower() == target_day.lower())

            if is_same_weekday or is_same_dayname:
                date_key = s_dt.strftime("%Y-%m-%d")
                p_list = s.get("players_data", [])
                if p_list:
                    df_s = pd.DataFrame(p_list)
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

                    g_metrics = calculer_metriques_groupes_seance(df_s)
                    r_stats_s = s.get("rpe_stats", {})
                    for grp_key in ["5 de devant", "3ème Ligne", "Trois-quarts"]:
                        df_group = _filter_df_by_group(df_s, grp_key)
                        rpe_moyen = None
                        if not df_group.empty and "RPE" in df_group.columns:
                            s_rpe = pd.to_numeric(df_group["RPE"], errors="coerce").dropna()
                            s_rpe = s_rpe[(s_rpe >= 1.0) & (s_rpe <= 10.0)]
                            if not s_rpe.empty:
                                rpe_moyen = round(float(s_rpe.mean()), 1)
                        if rpe_moyen is None and r_stats_s and isinstance(r_stats_s, dict):
                            r_val = r_stats_s.get(grp_key) or r_stats_s.get(f"rpe_moyen_{grp_key.lower().replace(' ', '_').replace('è', 'e')}")
                            if r_val is not None:
                                try:
                                    f_r = float(r_val)
                                    if f_r > 0:
                                        rpe_moyen = round(f_r, 1)
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
                        "metrics": calculer_moyennes_seance(df_s),
                        "group_metrics": g_metrics,
                    }

    # Intégration de la séance actuelle (exclue ou écrasant l'historique du même jour)
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

    curr_g_metrics = calculer_metriques_groupes_seance(df_current)
    for grp_key in ["5 de devant", "3ème Ligne", "Trois-quarts"]:
        df_group = _filter_df_by_group(df_current, grp_key)
        rpe_moyen = None
        if not df_group.empty and "RPE" in df_group.columns:
            rpe_moyen = df_group['RPE'].dropna().mean()
        if (rpe_moyen is None or pd.isna(rpe_moyen)) and current_rpe_stats and isinstance(current_rpe_stats, dict):
            r_val = current_rpe_stats.get(grp_key) or current_rpe_stats.get(f"rpe_moyen_{grp_key.lower().replace(' ', '_').replace('è', 'e')}")
            if r_val is not None:
                try:
                    f_r = float(r_val)
                    if f_r > 0:
                        rpe_moyen = f_r
                except Exception:
                    pass
        curr_g_metrics[grp_key]["rpe"] = round(rpe_moyen, 1) if pd.notna(rpe_moyen) else None

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
    # Conserver les 3 ou 4 dernières séances (séance actuelle incluse)
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
) -> Dict[str, Optional[float]]:
    """Calcule les moyennes RPE de la semaine pour l'équipe et les 3 groupes de postes."""
    res: Dict[str, Optional[float]] = {
        "Total": None,
        "5 de devant": None,
        "3ème Ligne": None,
        "Trois-quarts": None,
    }
    # 1. Depuis les sessions via rpe_connector
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

    # 2. Complément depuis df_week si présent
    if df_week is not None and not df_week.empty and "RPE" in df_week.columns:
        r_all = pd.to_numeric(df_week["RPE"], errors="coerce").dropna()
        if not r_all.empty and res["Total"] is None:
            res["Total"] = round(float(r_all.mean()), 1)
        for g_k, g_aliases in [
            ("5 de devant", ["5 de devant", "devant"]),
            ("3ème Ligne", ["3ème ligne", "3eme ligne"]),
            ("Trois-quarts", ["trois-quarts", "trois quarts", "3/4"]),
        ]:
            if res[g_k] is None and "Groupe" in df_week.columns:
                m_g = df_week["Groupe"].astype(str).str.strip().str.lower().apply(
                    lambda s: any(a in s for a in g_aliases)
                )
                sub_r = pd.to_numeric(df_week.loc[m_g, "RPE"], errors="coerce").dropna()
                if not sub_r.empty:
                    res[g_k] = round(float(sub_r.mean()), 1)

    return res


def recuperer_historique_semaines(
    year_curr: int,
    week_curr: int,
    df_curr_hebdo: pd.DataFrame,
    n_max: int = 4,
) -> List[Dict[str, Any]]:
    """
    Extrait l'historique des N dernières semaines enregistrées (semaine en cours incluse).
    """
    saved_weeks = list_saved_weeks()
    saved_weeks_chrono = sorted(saved_weeks, key=lambda x: (x["year"], x["week_num"]))

    weeks_dict: Dict[str, Dict[str, Any]] = {}
    for w_info in saved_weeks_chrono:
        y = w_info["year"]
        w = w_info["week_num"]
        if (y < year_curr) or (y == year_curr and w <= week_curr):
            w_key = f"{y}-W{w:02d}"
            w_data = load_week_data(y, w)
            df_w, _ = aggregate_weekly_sessions(w_data)
            label_str = f"Semaine {w:02d}"
            short_str = f"S{w:02d}"
            metrics_w = calculer_moyennes_seance(df_w)
            rpe_grp_w = _calculate_rpe_groups_for_week(df_w, w_data.get("sessions", {}))
            
            # Enrichissement des métriques de RPE par groupe pour le tableau comparatif
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
                "is_current": False,
                "num_players": len(df_w),
                "metrics": metrics_w,
                "rpe_groups": rpe_grp_w,
            }

    curr_key = f"{year_curr}-W{week_curr:02d}"
    curr_label = f"Semaine {week_curr:02d}"
    curr_short = f"S{week_curr:02d}"
    curr_metrics = calculer_moyennes_seance(df_curr_hebdo)
    curr_rpe_grp = _calculate_rpe_groups_for_week(df_curr_hebdo, None)
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
        "is_current": True,
        "num_players": len(df_curr_hebdo) if df_curr_hebdo is not None else 0,
        "metrics": curr_metrics,
        "rpe_groups": curr_rpe_grp,
    }

    sorted_weeks = sorted(weeks_dict.values(), key=lambda x: (x["year"], x["week_num"]))
    n_clamped = max(1, min(8, n_max))
    final_weeks = sorted_weeks[-n_clamped:]

    if final_weeks:
        final_weeks[-1]["is_current"] = True
        for it in final_weeks[:-1]:
            it["is_current"] = False

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

    colors_pl = ["#D4AF37" if it.get("is_current") else "#1F4E79" for it in items_comparatif]
    colors_hsr = ["#D4AF37" if it.get("is_current") else "#A30000" for it in items_comparatif]

    # --- AX1 : Player Load ---
    ax1.set_facecolor("#F8FAFC")
    bars1 = ax1.bar(x, pl_tot, width=0.42, color=colors_pl, edgecolor="#0F2B48", linewidth=0.8, zorder=3)
    ax1.plot(x, pl_av, color="#A30000", linestyle="--", marker="o", markersize=4, label="Avants (Pack)", zorder=5)
    ax1.plot(x, pl_ar, color="#2E5A88", linestyle=":", marker="s", markersize=4, label="Arrières (3/4)", zorder=5)

    ax1.set_title(f"Évolution Player Load ({type_label}s)", fontsize=9.2, fontweight="bold", color="#1F4E79", pad=6)
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
    ax2.plot(x, hsr_ar, color="#2E5A88", linestyle=":", marker="s", markersize=4, label="Arrières (3/4)", zorder=5)

    ax2.set_title(f"Évolution HSR (> 20 km/h) ({type_label}s)", fontsize=9.2, fontweight="bold", color="#1F4E79", pad=6)
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
                     fontsize=7.2, fontweight="bold", color="#D97706" if it.get("is_current") else "#A30000")

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
            (rpe_5dev, "#C0392B", "o", "5 de devant"),
            (rpe_3eme, "#D97706", "s", "3ème Ligne"),
            (rpe_3quarts, "#2980B9", "^", "Trois-quarts"),
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
            ax3.plot(valid_eq_x, valid_eq_y, color="#1E293B", linestyle="--", linewidth=2.0, marker="D", markersize=4.0, label="Équipe", zorder=5)

        ax3.legend(loc="upper right", frameon=False, fontsize=6.2)
    else:
        ax3.text(0.5, 0.5, "Aucune saisie RPE disponible", ha="center", va="center", transform=ax3.transAxes,
                 fontsize=8.5, color="#94A3B8", fontstyle="italic")

    ax3.set_title(f"Évolution RPE Moyen par Groupe ({type_label}s)", fontsize=9.2, fontweight="bold", color="#1F4E79", pad=6)
    ax3.set_xticks(x)
    ax3.set_xticklabels(labels, fontsize=8.2, fontweight="bold", color="#334155")
    ax3.set_ylim(0, 10.5)
    ax3.set_ylabel("RPE (1-10)", fontsize=7.8, fontweight="bold", color="#1F4E79")
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
) -> io.BytesIO:
    """
    Génère le triptyque graphique officiel Staff RCA (3 colonnes côte à côte) :
    - Colonne 1 : 5 de devant
    - Colonne 2 : 3ème Ligne
    - Colonne 3 : Trois-quarts (3/4)
    Histogramme à barres groupées pour chaque séance historique :
    - Barre Bleue (#3498db / Bleu roi) : Volume / Distance totale (m) (Axe Y gauche)
    - Barre Rouge / Bordeaux (#c0392b) : HSR (> 20 km/h ou Z3) (m) (Axe Y gauche)
    - Barre Verte (#2ecc71 / Vert flash) : VHSR / Sprint (> 25 km/h ou Z5) (m) (Axe Y gauche)
    - Barre Orange / Ambre (#E67E22) : RPE Moyen de séance (Axe Y secondaire droit twinx, échelle 1-10)
    Étiquettes exactes de valeurs en gras au-dessus de chaque barre (ex: 2238, 188, 119, 7.2).
    Axe X avec date et adversaire (ex: vendredi 10/10 (RHEU), vendredi 06/11 (VERSAILLES)).
    Légende commune en haut : ● Volume  ● HSR  ● VHSR  ■ RPE.
    """
    fig, axes = plt.subplots(1, 3, figsize=(14.5, 6.8), dpi=200)
    fig.patch.set_facecolor("#FFFFFF")

    group_defs = [
        ("5 de devant", "5 DE DEVANT"),
        ("3ème Ligne", "3ÈME LIGNE"),
        ("Trois-quarts", "TROIS-QUARTS (3/4)"),
    ]

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

        # 3 barres de charge externe sur l'axe Y principal (mètres)
        b_vol = ax.bar(x - 1.5 * w, v_vol, width=w, color="#3498db", edgecolor="none", zorder=3)
        b_hsr = ax.bar(x - 0.5 * w, v_hsr, width=w, color="#c0392b", edgecolor="none", zorder=3)
        b_vhsr = ax.bar(x + 0.5 * w, v_vhsr, width=w, color="#2ecc71", edgecolor="none", zorder=3)

        # 1 barre de charge interne RPE sur l'axe Y secondaire (1 à 10)
        rpe_bar_heights = [float(v) if (v is not None and pd.notna(v) and float(v) > 0) else 0.0 for v in v_rpe]
        b_rpe = ax_rpe.bar(
            x + 1.5 * w,
            rpe_bar_heights,
            width=w,
            color="#E67E22",
            edgecolor="#C2410C",
            linewidth=0.8,
            alpha=0.92,
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
                        fontsize=7.5,
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

        ax.set_title(grp_title, fontsize=11.5, fontweight="bold", color="#1F4E79", pad=12)
        ax.set_xticks(x)
        x_labels = [it.get("label_x", it.get("label", "")) for it in items_comparatif]
        ax.set_xticklabels([lbl.replace(" (", "\n(") for lbl in x_labels], fontsize=8.2, fontweight="bold", color="#334155")
        ax.set_xlim(-0.55, max(float(K) - 0.45, 0.55))

        max_v = max(max(v_vol or [100.0]), 100.0)
        ax.set_ylim(0, max_v * 1.25)
        ax.set_ylabel("Mètres (m)", fontsize=7.8, fontweight="bold", color="#1F4E79")
        ax.tick_params(axis="x", length=0)
        ax.tick_params(axis="y", labelsize=7.5, color="#94A3B8")

    leg_elements = [
        Line2D([0], [0], marker="o", color="w", markerfacecolor="#3498db", markersize=9, label="Volume (m)"),
        Line2D([0], [0], marker="o", color="w", markerfacecolor="#c0392b", markersize=9, label="HSR (> 20 km/h) (m)"),
        Line2D([0], [0], marker="o", color="w", markerfacecolor="#2ecc71", markersize=9, label="VHSR (> 25 km/h) (m)"),
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
) -> None:
    """
    Rend la page de comparaison historique multi-séances au format triptyque Staff RCA :
    - En-tête de section (Titre + Sous-titre dynamique)
    - 3 colonnes graphiques côte à côte (5 de devant, 3ème Ligne, Trois-quarts)
    - Histogrammes groupés Volume, HSR, VHSR
    - Barre orange RPE découplée (échelle 1 à 10)
    - Pleine largeur A4 Paysage sans tableau compact
    """
    if not items_comparatif:
        pdf.set_xy(12.0, 30.0)
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(273.0, 10.0, safe_latin1("Aucune donnée comparative disponible pour ce jour."), align="C")
        return

    # Garantie absolue d'injection du RPE par groupe
    for it in items_comparatif:
        g_metrics = it.setdefault("group_metrics", {})
        is_curr = it.get("is_current", False)
        for grp in ["5 de devant", "3ème Ligne", "Trois-quarts"]:
            grp_dict = g_metrics.setdefault(grp, {})
            if grp_dict.get("rpe") is None or pd.isna(grp_dict.get("rpe")):
                df_src = None
                if is_curr and df_current is not None and not df_current.empty:
                    df_src = df_current
                elif "players_data" in it and it["players_data"]:
                    df_src = pd.DataFrame(it["players_data"])

                if df_src is not None and not df_src.empty:
                    df_group = _filter_df_by_group(df_src, grp)
                    if not df_group.empty and "RPE" in df_group.columns:
                        s_rpe = pd.to_numeric(df_group["RPE"], errors="coerce").dropna()
                        rpe_moyen = s_rpe.mean() if not s_rpe.empty else None
                        grp_dict["rpe"] = round(rpe_moyen, 1) if pd.notna(rpe_moyen) else None

    y_top = 19.5
    has_priors = len(items_comparatif) >= 2
    if not has_priors:
        pdf.set_fill_color(254, 243, 199)
        pdf.set_draw_color(245, 158, 11)
        pdf.rect(12.0, y_top, 273.0, 5.5, style="FD")
        pdf.set_xy(12.0, y_top)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.set_text_color(146, 64, 14)
        info_txt = f"HISTORIQUE : Aucune séance antérieure enregistrée pour un {target_day.lower()}. Données de la séance affichées seules."
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
        )
        pdf.image(buf_chart, x=12.0, y=y_chart, w=273.0, h=h_chart)
    except Exception as e:
        print(f"Erreur graphique triptyque comparatif PDF : {e}")


def render_comparatif_historique_pdf_page(
    pdf: Any,
    items_comparatif: List[Dict[str, Any]],
    type_label: str = "Séance",
    target_day: str = "",
) -> None:
    """
    Rend la page dédiée 'Comparatif Historique' au format A4 Paysage :
    - En-tête / Bannière informative si historique insuffisant (< 2 éléments).
    - Haut de page : Graphique d'évolution Matplotlib (Player Load & HSR).
    - Bas de page : Tableau de synthèse comparative multi-séances / multi-semaines avec colonnes dynamiques (2 à 8)
      et calcul du Delta (%) de la séance actuelle par rapport à la moyenne des séances antérieures.
    """
    if not items_comparatif:
        pdf.set_xy(12.0, 30.0)
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(273.0, 10.0, safe_latin1("Aucune donnee comparative disponible."), align="C")
        return

    y_top = 19.5
    has_priors = len(items_comparatif) >= 2
    if not has_priors:
        pdf.set_fill_color(254, 243, 199)
        pdf.set_draw_color(245, 158, 11)
        pdf.rect(12.0, y_top, 273.0, 5.0, style="FD")
        pdf.set_xy(12.0, y_top)
        pdf.set_font("Helvetica", "B", 7.2)
        pdf.set_text_color(146, 64, 14)
        info_txt = f"HISTORIQUE INSUFFISANT : Aucune {type_label.lower()} anterieure enregistree{f' pour un {target_day}' if target_day else ''}. Données actuelles affichées sans delta."
        pdf.cell(273.0, 5.0, safe_latin1(info_txt), align="C")
        y_chart_start = y_top + 6.0
    else:
        y_chart_start = y_top

    # 1. Section 1 : Graphique d'évolution Matplotlib
    pdf.set_xy(12.0, y_chart_start)
    pdf.set_font("Helvetica", "B", 8.8)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(273.0, 4.0, safe_latin1("1. ÉVOLUTION DE LA CHARGE (PLAYER LOAD), HAUTE INTENSITÉ (HSR) ET FATIGUE RPE :"), align="L")

    try:
        buf_chart = generer_graphique_comparatif_historique_matplotlib(items_comparatif, type_label=type_label)
        chart_h = 65.0
        pdf.image(buf_chart, x=12.0, y=y_chart_start + 4.5, w=273.0, h=chart_h)
    except Exception as e:
        print(f"Erreur graphique comparatif historique PDF : {e}")
        chart_h = 20.0

    # 2. Section 2 : Tableau de synthèse comparative
    y_table_sec = y_chart_start + 4.5 + chart_h + 2.0
    pdf.set_xy(12.0, y_table_sec)
    pdf.set_font("Helvetica", "B", 8.8)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(273.0, 4.0, safe_latin1("2. TABLEAU DE SYNTHÈSE COMPARATIVE (MOYENNES COLLECTIVES & PAR LIGNE) :"), align="L")

    num_items = len(items_comparatif)
    w_label = 55.0
    w_delta = 28.0
    w_col = (273.0 - w_label - w_delta) / max(num_items, 1)

    y_tbl_hdr = y_table_sec + 4.2
    pdf.set_xy(12.0, y_tbl_hdr)
    pdf.set_fill_color(31, 78, 121)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 7.5)
    pdf.cell(w_label, 5.0, safe_latin1("Métrique & Groupe"), 1, 0, "L", fill=True)

    for it in items_comparatif:
        is_curr = it.get("is_current", False)
        if is_curr:
            pdf.set_fill_color(21, 52, 82)
            pdf.set_text_color(212, 175, 55)
            h_text = f"{it.get('short_label', '')} (Act.)"
        else:
            pdf.set_fill_color(31, 78, 121)
            pdf.set_text_color(255, 255, 255)
            h_text = it.get("short_label", it.get("label", ""))
        pdf.set_font("Helvetica", "B", 7.0 if w_col < 30 else 7.5)
        pdf.cell(w_col, 5.0, safe_latin1(h_text[:14]), 1, 0, "C", fill=True)

    pdf.set_fill_color(31, 78, 121)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 7.2)
    pdf.cell(w_delta, 5.0, safe_latin1("Écart vs Moy."), 1, 1, "C", fill=True)

    metric_blocks = [
        {
            "title": "1. DISTANCE TOTALE (m)",
            "key": "Distance Totale (m)",
            "fmt": lambda v: f"{int(round(v)):,d} m".replace(",", " "),
        },
        {
            "title": "2. HIGH SPEED RUNNING (> 20 km/h) (m)",
            "key": "HSR (> 20 km/h) (m)",
            "fmt": lambda v: f"{int(round(v)):,d} m".replace(",", " "),
        },
        {
            "title": "3. ACCÉLÉRATIONS & DÉCÉLÉRATIONS INTENSES (> 3 m/s²)",
            "key": "Acc/Dec (> 3 m/s²)",
            "fmt": lambda v: f"{v:.1f}",
        },
        {
            "title": "4. INTENSITÉ RELATIVE (m/min)",
            "key": "m/min",
            "fmt": lambda v: f"{v:.1f} m/min",
        },
        {
            "title": "5. PLAYER LOAD (CHARGE GLOBALE)",
            "key": "Player Load",
            "fmt": lambda v: f"{v:.1f}",
        },
        {
            "title": "6. CHARGE INTERNE & FATIGUE (RPE 1-10)",
            "key": "RPE Moyen (1-10)",
            "fmt": lambda v: f"{v:.1f}/10" if (v is not None and v > 0) else "NR",
        },
    ]

    row_h = 3.2
    sub_levels = [
        ("Total (Collectif)", "Total"),
        ("Avants (Pack)", "Avants"),
        ("Arrières (Trois-quarts)", "Arrières"),
    ]

    for block in metric_blocks:
        pdf.set_x(12.0)
        pdf.set_fill_color(241, 245, 249)
        pdf.set_text_color(31, 78, 121)
        pdf.set_font("Helvetica", "B", 6.8)
        pdf.cell(273.0, 3.4, safe_latin1(block["title"]), 1, 1, "L", fill=True)

        for sub_label, sub_key in sub_levels:
            pdf.set_x(12.0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_draw_color(226, 232, 240)
            pdf.set_font("Helvetica", "", 6.6)
            pdf.set_text_color(30, 41, 59)
            pdf.cell(w_label, row_h, safe_latin1(f"   {sub_label}"), 1, 0, "L", fill=True)

            curr_val = items_comparatif[-1]["metrics"][block["key"]][sub_key] if items_comparatif else 0.0
            prior_vals = [it["metrics"][block["key"]][sub_key] for it in items_comparatif[:-1]] if len(items_comparatif) > 1 else []

            for it in items_comparatif:
                is_curr = it.get("is_current", False)
                v = it["metrics"][block["key"]][sub_key]
                v_txt = block["fmt"](v)
                pdf.set_fill_color(248, 250, 252) if is_curr else (255, 255, 255)
                pdf.set_font("Helvetica", "B" if is_curr else "", 6.6 if w_col < 30 else 7.0)
                pdf.set_text_color(31, 78, 121) if is_curr else (30, 41, 59)
                pdf.cell(w_col, row_h, safe_latin1(v_txt), 1, 0, "C", fill=True)

            if prior_vals:
                valid_priors = [pv for pv in prior_vals if pv is not None and pv > 0]
                if valid_priors and curr_val is not None and curr_val > 0:
                    moy_prior = sum(valid_priors) / len(valid_priors)
                    delta_pct = ((curr_val - moy_prior) / moy_prior) * 100.0
                    delta_txt = f"{delta_pct:+.1f} %"
                    d_color = (22, 101, 52) if delta_pct >= 0 else (153, 27, 27)
                elif curr_val is not None and curr_val > 0 and not valid_priors:
                    delta_txt = "-"
                    d_color = (100, 116, 139)
                else:
                    delta_txt = "-"
                    d_color = (100, 116, 139)
            else:
                delta_txt = "-"
                d_color = (100, 116, 139)

            pdf.set_font("Helvetica", "B", 6.6)
            pdf.set_text_color(*d_color)
            pdf.set_fill_color(255, 255, 255)
            pdf.cell(w_delta, row_h, safe_latin1(delta_txt), 1, 1, "C", fill=True)


def _render_podium_section_pdf(
    pdf: Any,
    podium_data: Dict[str, Dict[str, Any]],
    x_pod: float,
    y_pod: float,
    w_pod: float,
    title_header: str,
    subtitle_map: Optional[Dict[str, str]] = None,
) -> None:
    """Rend un bloc podium élégant de 3 catégories (Top Speed, % Marche, Haute Intensité)."""
    pdf.set_fill_color(31, 78, 121)
    pdf.rect(x_pod, y_pod, w_pod, 6.0, style="F")
    pdf.set_fill_color(212, 175, 55)  # Liseré doré #D4AF37
    pdf.rect(x_pod, y_pod + 5.5, w_pod, 0.5, style="F")

    pdf.set_xy(x_pod, y_pod)
    pdf.set_font("Helvetica", "B", 9.0)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(w_pod, 5.5, safe_latin1(title_header), align="C")

    categories_config = [
        {
            "key": "Top speed (km/h)",
            "title": "TOP SPEED (km/h)",
            "def_sub": "Vitesse Maximale",
            "color": (31, 78, 121),
        },
        {
            "key": "% de marche",
            "title": "% DE MARCHE",
            "def_sub": "Moins marche (> 0%)",
            "color": (46, 90, 136),
        },
        {
            "key": "Haute intensité (m)",
            "title": "HAUTE INTENSITÉ (m)",
            "def_sub": "Cumul Z3 + Z4 + Z5 > 14.4 km/h",
            "color": (163, 0, 0),
        },
    ]

    rank_badges = {
        1: {"label": "1er", "bg": (217, 119, 6), "fg": (255, 255, 255)},
        2: {"label": "2e", "bg": (100, 116, 139), "fg": (255, 255, 255)},
        3: {"label": "3e", "bg": (180, 83, 9), "fg": (255, 255, 255)},
    }

    y_card_start = y_pod + 6.0 + 2.0
    card_h = 42.0
    card_gap = 2.5

    for i, cat_cfg in enumerate(categories_config):
        y_card = y_card_start + i * (card_h + card_gap)

        pdf.set_fill_color(248, 250, 252)
        pdf.set_draw_color(203, 213, 225)
        pdf.rect(x_pod, y_card, w_pod, card_h, style="FD")

        accent_color = cat_cfg["color"]
        pdf.set_fill_color(*accent_color)
        pdf.rect(x_pod, y_card, w_pod, 5.8, style="F")

        pdf.set_xy(x_pod + 3.5, y_card)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(65, 5.8, safe_latin1(cat_cfg["title"]), align="L")

        sub_val = subtitle_map.get(cat_cfg["key"]) if isinstance(subtitle_map, dict) else None
        sub_txt = str(sub_val).strip() if sub_val is not None else cat_cfg["def_sub"]
        pdf.set_xy(x_pod + 65, y_card)
        pdf.set_font("Helvetica", "I", 7.5)
        pdf.set_text_color(226, 232, 240)
        pdf.cell(w_pod - 69, 5.8, safe_latin1(sub_txt), align="R")

        cat_data = podium_data.get(cat_cfg["key"], {}) if isinstance(podium_data, dict) else {}
        players = cat_data.get("players", []) if isinstance(cat_data, dict) else []

        row_w = w_pod - 4
        row_h = 10.0
        row_gap = 1.3
        y_rows_start = y_card + 5.8 + 1.2

        for rank_idx in range(1, 4):
            y_row = y_rows_start + (rank_idx - 1) * (row_h + row_gap)
            x_row = x_pod + 2

            p_entry = players[rank_idx - 1] if len(players) >= rank_idx and isinstance(players[rank_idx - 1], dict) else {}
            p_name = safe_latin1(p_entry.get("name", "-"))
            p_stat = safe_latin1(p_entry.get("stat", "-"))

            pdf.set_fill_color(255, 255, 255)
            pdf.set_draw_color(226, 232, 240)
            pdf.rect(x_row, y_row, row_w, row_h, style="FD")

            badge_info = rank_badges.get(rank_idx, {"label": f"{rank_idx}e", "bg": (100, 116, 139), "fg": (255, 255, 255)})
            photo_path = photo_manager.get_player_photo_path(p_name) if p_name and p_name != "-" else None
            has_photo = bool(photo_path and os.path.isfile(photo_path) and os.path.basename(photo_path) not in ["avatar_par_defaut.jpg", "avatar_par_defaut.png"])

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


def _render_podium_page_aere(
    pdf: Any,
    podium_data: Dict[str, Dict[str, Any]],
    group_name: str,
    subtitle_map: Optional[Dict[str, str]] = None,
) -> None:
    """
    Rend une page podium dédiée, ultra-aérée et prestigieuse pour le RCA (Général ou par ligne de poste).
    Organise les 3 catégories (Top Speed, % Marche, Haute Intensité) en 3 grandes colonnes verticales
    côte-à-côte avec bandeau supérieur, badges de podium, photos de profil et statistiques en grand format.
    Gère gracieusement le cas où un groupe positionnel compte strictement moins de 3 joueurs.
    """
    categories_config = [
        {
            "key": "Top speed (km/h)",
            "title": "TOP SPEED (km/h)",
            "def_sub": "Vitesse Maximale",
            "color": (31, 78, 121),
        },
        {
            "key": "% de marche",
            "title": "% DE MARCHE",
            "def_sub": "Moins marche (> 0%)",
            "color": (46, 90, 136),
        },
        {
            "key": "Haute intensité (m)",
            "title": "HAUTE INTENSITÉ (m)",
            "def_sub": "Cumul Z3 + Z4 + Z5 > 14.4 km/h",
            "color": (163, 0, 0),
        },
    ]

    rank_badges = {
        1: {"label": "1er", "bg": (217, 119, 6), "fg": (255, 255, 255)},
        2: {"label": "2e", "bg": (100, 116, 139), "fg": (255, 255, 255)},
        3: {"label": "3e", "bg": (180, 83, 9), "fg": (255, 255, 255)},
    }

    col_w = 87.0
    col_gap = 6.0
    x_start = 12.0
    y_start = 23.0
    banner_h = 14.0
    card_h = 47.0
    card_gap = 3.5

    for c_idx, cat in enumerate(categories_config):
        x_c = x_start + c_idx * (col_w + col_gap)
        accent_color = cat["color"]

        # 1. En-tête de catégorie
        pdf.set_fill_color(*accent_color)
        pdf.rect(x_c, y_start, col_w, banner_h, style="F")

        # Liseré doré décoratif
        pdf.set_fill_color(212, 175, 55)
        pdf.rect(x_c, y_start + banner_h - 1.0, col_w, 1.0, style="F")

        pdf.set_xy(x_c + 2.0, y_start + 1.2)
        pdf.set_font("Helvetica", "B", 11.0)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(col_w - 4.0, 5.5, safe_latin1(cat["title"]), align="C")

        sub_val = subtitle_map.get(cat["key"]) if isinstance(subtitle_map, dict) else None
        sub_txt = str(sub_val).strip() if sub_val is not None else cat["def_sub"]
        pdf.set_xy(x_c + 2.0, y_start + 7.2)
        pdf.set_font("Helvetica", "I", 8.5)
        pdf.set_text_color(226, 232, 240)
        pdf.cell(col_w - 4.0, 4.8, safe_latin1(sub_txt), align="C")

        # 2. Les 3 fiches joueurs
        cat_data = podium_data.get(cat["key"], {}) if isinstance(podium_data, dict) else {}
        players = cat_data.get("players", []) if isinstance(cat_data, dict) else []

        y_cards_start = y_start + banner_h + 3.0

        for r_idx in range(1, 4):
            y_card = y_cards_start + (r_idx - 1) * (card_h + card_gap)
            p_entry = players[r_idx - 1] if len(players) >= r_idx and isinstance(players[r_idx - 1], dict) else {}
            p_name = safe_latin1(p_entry.get("name", "-"))
            p_stat = safe_latin1(p_entry.get("stat", "-"))

            # Carte fond blanc avec contour fin
            pdf.set_fill_color(255, 255, 255)
            pdf.set_draw_color(203, 213, 225)
            pdf.rect(x_c, y_card, col_w, card_h, style="FD")

            # Petit liseré d'accent à gauche de la carte
            pdf.set_fill_color(*accent_color)
            pdf.rect(x_c, y_card, 2.5, card_h, style="F")

            # Badge de rang (1er, 2e, 3e) en haut à gauche
            badge = rank_badges.get(r_idx, {"label": f"{r_idx}e", "bg": (100, 116, 139), "fg": (255, 255, 255)})
            pdf.set_fill_color(*badge["bg"])
            pdf.rect(x_c + 5.0, y_card + 3.0, 15.0, 7.5, style="F")
            pdf.set_xy(x_c + 5.0, y_card + 3.0)
            pdf.set_font("Helvetica", "B", 8.5)
            pdf.set_text_color(*badge["fg"])
            pdf.cell(15.0, 7.5, badge["label"], align="C")

            # Statistique en haut à droite
            pdf.set_xy(x_c + 22.0, y_card + 3.0)
            pdf.set_font("Helvetica", "B", 10.0)
            pdf.set_text_color(*accent_color)
            pdf.cell(col_w - 26.0, 7.5, safe_latin1(p_stat), align="R")

            # Photo du joueur au centre
            photo_path = photo_manager.get_player_photo_path(p_name) if p_name and p_name != "-" else None
            has_photo = bool(photo_path and os.path.isfile(photo_path) and os.path.basename(photo_path) not in ["avatar_par_defaut.jpg", "avatar_par_defaut.png"])
            photo_size = 24.0
            photo_x = x_c + (col_w - photo_size) / 2.0
            photo_y = y_card + 11.5

            if has_photo:
                try:
                    pdf.image(photo_path, x=photo_x, y=photo_y, w=photo_size, h=photo_size)
                    pdf.set_draw_color(203, 213, 225)
                    pdf.rect(photo_x, photo_y, photo_size, photo_size, style="D")
                except Exception:
                    has_photo = False

            if not has_photo:
                # Avatar placeholder propre
                pdf.set_fill_color(241, 245, 249)
                pdf.set_draw_color(203, 213, 225)
                pdf.rect(photo_x, photo_y, photo_size, photo_size, style="FD")
                pdf.set_xy(photo_x, photo_y + 6.0)
                pdf.set_font("Helvetica", "B", 12)
                pdf.set_text_color(148, 163, 184)
                pdf.cell(photo_size, 10, "?", align="C")

            # Nom du joueur en bas
            name_disp = p_name if len(p_name) <= 22 else p_name[:20] + ".."
            pdf.set_xy(x_c + 3.0, y_card + 37.0)
            pdf.set_font("Helvetica", "B", 9.5)
            pdf.set_text_color(30, 41, 59)
            pdf.cell(col_w - 6.0, 7.0, safe_latin1(name_disp), align="C")


def _render_podiums_compact_horizontal(
    pdf: Any,
    podium_data: Dict[str, Dict[str, Any]],
    y_start: float,
    col_w: float = 87.5,
    col_gap: float = 5.25,
    subtitle_map: Optional[Dict[str, str]] = None,
    is_large: bool = False,
    x_start: float = 12.0,
) -> float:
    """
    Rend un bloc podium horizontal élégant de 3 catégories côte-à-côte :
    - Top Speed (km/h)
    - % de marche
    - Haute intensité (m)

    Deux formats supportés :
    - is_large=True (Page 1 - Podium Général) : cartes aérées avec photos 14x14 mm, nom, poste et stats grand format.
    - is_large=False (Pages 2, 3, 4 - Lignes de poste) : format compact haute densité avec photos 8.4x8.4 mm.

    Gère gracieusement le cas où un groupe compte strictement moins de 3 joueurs ou 0 joueur sans exception.
    Retourne la hauteur totale occupée (mm).
    """
    categories_config = [
        {
            "key": "Top speed (km/h)",
            "title": "TOP SPEED (km/h)",
            "def_sub": "Vitesse Maximale",
            "color": (31, 78, 121),
        },
        {
            "key": "% de marche",
            "title": "% DE MARCHE",
            "def_sub": "Moins marche (> 0%)",
            "color": (46, 90, 136),
        },
        {
            "key": "Haute intensité (m)",
            "title": "HAUTE INTENSITÉ (m)",
            "def_sub": "Cumul Z3+Z4+Z5 > 14.4 km/h",
            "color": (163, 0, 0),
        },
    ]

    rank_badges = {
        1: {"label": "1er", "bg": (217, 119, 6), "fg": (255, 255, 255)},
        2: {"label": "2e", "bg": (100, 116, 139), "fg": (255, 255, 255)},
        3: {"label": "3e", "bg": (180, 83, 9), "fg": (255, 255, 255)},
    }

    if is_large:
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

    y_rows_start = y_start + banner_h + (1.2 if is_large else 1.0)

    for c_idx, cat in enumerate(categories_config):
        x_c = x_start + c_idx * (col_w + col_gap)
        accent_color = cat["color"]

        # 1. En-tête de catégorie
        pdf.set_fill_color(*accent_color)
        pdf.rect(x_c, y_start, col_w, banner_h, style="F")

        # Liseré doré décoratif
        pdf.set_fill_color(212, 175, 55)
        pdf.rect(x_c, y_start + banner_h - 0.5, col_w, 0.5, style="F")

        pdf.set_xy(x_c + 2.0, y_start)
        pdf.set_font("Helvetica", "B", title_font)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(col_w * 0.55, banner_h - 0.5, safe_latin1(cat["title"]), align="L")

        sub_val = subtitle_map.get(cat["key"]) if isinstance(subtitle_map, dict) else None
        sub_txt = str(sub_val).strip() if sub_val is not None else cat["def_sub"]
        pdf.set_xy(x_c + col_w * 0.55, y_start)
        pdf.set_font("Helvetica", "I", sub_font)
        pdf.set_text_color(241, 245, 249)
        pdf.cell(col_w * 0.45 - 2.0, banner_h - 0.5, safe_latin1(sub_txt), align="R")

        # 2. Les 3 lignes joueurs
        cat_data = podium_data.get(cat["key"], {}) if isinstance(podium_data, dict) else {}
        players = cat_data.get("players", []) if isinstance(cat_data, dict) else []

        for r_idx in range(1, 4):
            y_row = y_rows_start + (r_idx - 1) * (row_h + row_gap)
            p_entry = players[r_idx - 1] if len(players) >= r_idx and isinstance(players[r_idx - 1], dict) else {}
            p_name = safe_latin1(p_entry.get("name", "-")).strip()
            p_poste = safe_latin1(p_entry.get("poste", "")).strip()
            p_stat = safe_latin1(p_entry.get("stat", "-")).strip()

            # Fond de carte blanc
            pdf.set_fill_color(255, 255, 255)
            pdf.set_draw_color(226, 232, 240)
            pdf.rect(x_c, y_row, col_w, row_h, style="FD")

            # Liseré d'accent vertical à gauche
            accent_w = 2.0 if is_large else 1.5
            pdf.set_fill_color(*accent_color)
            pdf.rect(x_c, y_row, accent_w, row_h, style="F")

            # Badge de rang (1er, 2e, 3e)
            badge = rank_badges.get(r_idx, {"label": f"{r_idx}e", "bg": (100, 116, 139), "fg": (255, 255, 255)})
            b_w = 10.0 if is_large else 7.5
            b_h = 8.0 if is_large else 7.5
            b_x = x_c + accent_w + 1.2
            b_y = y_row + (row_h - b_h) / 2.0
            pdf.set_fill_color(*badge["bg"])
            pdf.rect(b_x, b_y, b_w, b_h, style="F")
            pdf.set_xy(b_x, b_y)
            pdf.set_font("Helvetica", "B", 7.0 if is_large else 6.0)
            pdf.set_text_color(*badge["fg"])
            pdf.cell(b_w, b_h, badge["label"], align="C")

            # Photo / Avatar
            photo_x = b_x + b_w + 1.5
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

            # Textes et statistiques
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
                x_text = photo_x + photo_size + 1.8
                x_stat = x_c + col_w - 29.0
                w_name = x_stat - x_text - 1.0

                disp_name = p_name if len(p_name) <= 18 else p_name[:16] + ".."
                pdf.set_xy(x_text, y_row + 1.0)
                pdf.set_font("Helvetica", "B", 6.8)
                pdf.set_text_color(30, 41, 59)
                pdf.cell(w_name, 4.2, safe_latin1(disp_name), align="L")

                disp_poste = p_poste if p_poste else "RCA"
                if len(disp_poste) > 15:
                    disp_poste = disp_poste[:13] + ".."
                pdf.set_xy(x_text, y_row + 5.2)
                pdf.set_font("Helvetica", "", 5.8)
                pdf.set_text_color(100, 116, 139)
                pdf.cell(w_name, 3.8, safe_latin1(disp_poste), align="L")

                pdf.set_xy(x_stat, y_row)
                pdf.set_font("Helvetica", "B", 7.8)
                pdf.set_text_color(*accent_color)
                pdf.cell(27.5, row_h, safe_latin1(p_stat), align="R")

    total_h = banner_h + (1.2 if is_large else 1.0) + 3 * row_h + 2 * row_gap
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
        # 1. Bandeau officiel RCA (#1F4E79 Bleu nuit officiel)
        self.set_fill_color(31, 78, 121)
        self.rect(0, 0, 297, 16.5, "F")

        # 2. Liseré doré officiel (#D4AF37)
        self.set_fill_color(212, 175, 55)
        self.rect(0, 16.5, 297, 1.2, "F")

        # 3. Logo officiel du RCA
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_rca.png")
        if not os.path.exists(logo_path):
            logo_path = "logo_rca.png"
        x_title = 12
        if os.path.exists(logo_path):
            try:
                self.image(logo_path, x=12, y=2.2, w=12, h=12)
                x_title = 27
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
            self.set_text_color(212, 175, 55)
            self.cell(150, 6, safe_latin1(self.current_subtitle), align="L")
        else:
            self.set_xy(x_title, 3.2)
            self.set_font("Helvetica", "B", 12.0)
            self.set_text_color(255, 255, 255)
            self.cell(150, 10, safe_latin1("RUGBY CLUB AMIÉNOIS | Rapport Hebdomadaire de Charge GPS"), align="L")

        # 5. Semaine et dates
        self.set_font("Helvetica", "B", 9.5)
        self.set_text_color(212, 175, 55)
        self.set_xy(175, 3.2)
        dates_disp = f" - {self.dates_label}" if self.dates_label else ""
        self.cell(110, 10, safe_latin1(f"{self.week_label.upper()}{dates_disp}"), align="R")

    def footer(self):
        self.set_y(-9)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(100, 116, 139)
        self.cell(140, 6, safe_latin1("Cellule Performance & Préparation Physique RCA Amiens - Confidentiel Staff"), align="L")
        self.cell(133, 6, safe_latin1(f"Page {self.page_no()}/{self.total_pages}"), align="R")


def generate_weekly_pdf_report(
    df_hebdo: Optional[pd.DataFrame] = None,
    year: int = 2026,
    week_num: int = 1,
    synthese_jours: Optional[Union[Dict[str, Any], List[Dict[str, Any]]]] = None,
    notes_coach: str = "",
    objectifs_config: Optional[Dict[str, Dict[str, float]]] = None,
    inclure_comparatif: bool = False,
    nb_comparatif: int = 4,
    **kwargs: Any,
) -> bytes:
    """Génère le rapport hebdomadaire PDF complet officiel RCA au format Paysage (A4)."""
    if df_hebdo is None:
        df_hebdo = kwargs.get("df_agg", pd.DataFrame())

    if synthese_jours is None:
        synthese_jours = kwargs.get("selected_sessions_meta", {})

    if isinstance(synthese_jours, list):
        synthese_dict = {f"s_{i}": s for i, s in enumerate(synthese_jours) if isinstance(s, dict)}
    elif isinstance(synthese_jours, dict):
        synthese_dict = synthese_jours
    else:
        synthese_dict = {}

    week_label = f"Semaine {week_num:02d} - {year}"
    raw_dates = [format_readable_date(s.get("date", "")) for s in synthese_dict.values() if isinstance(s, dict) and s.get("date")]
    raw_dates = [d for d in raw_dates if d]
    dates_str = f"Du {raw_dates[0]} au {raw_dates[-1]}" if len(raw_dates) >= 2 else (raw_dates[0] if raw_dates else "")

    total_pages = 5 if inclure_comparatif else 4
    pdf = RCAWeeklyPDFReport(week_label=week_label, dates_label=dates_str, total_pages=total_pages)

    if df_hebdo is None or df_hebdo.empty:
        pdf.add_page()
        pdf.set_font("Helvetica", "", 12)
        pdf.cell(0, 10, safe_latin1("Aucune donnee disponible pour cette semaine."), align="L")
        return bytes(pdf.output())

    # =========================================================================
    # PAGE 1 : SYNTHÈSE GLOBALE DU CLUB & PODIUM GÉNÉRAL (FORMAT STRICT 4 PAGES)
    # =========================================================================
    pdf.current_subtitle = "SYNTHÈSE GLOBALE DU CLUB & CHARGE HEBDOMADAIRE"
    pdf.add_page()
    pdf.set_y(20.0)

    # 1. KPIs Globaux
    tot_km = float(df_hebdo["Distance Totale (km)"].sum()) if not df_hebdo.empty and "Distance Totale (km)" in df_hebdo.columns else 0.0
    tot_hi = float(df_hebdo["Haute intensité (m)"].sum()) if not df_hebdo.empty and "Haute intensité (m)" in df_hebdo.columns else 0.0
    nb_joueurs_tot = len(df_hebdo)
    max_speed = float(df_hebdo["Top Speed (km/h)"].max()) if not df_hebdo.empty and "Top Speed (km/h)" in df_hebdo.columns else 0.0
    nb_seances = len(synthese_dict)

    weekly_rpe = {}
    if rpe_connector:
        try:
            weekly_rpe = rpe_connector.calculate_weekly_rpe_from_sessions(synthese_dict)
        except Exception:
            pass

    kpis = [
        ("Volume Cumulé Équipe", f"{tot_km:.1f} km"),
        ("Haute Intensité Cumulée", f"{int(round(tot_hi)):,d} m".replace(",", " ")),
        ("Joueurs Actifs Semaine", f"{nb_joueurs_tot} joueurs"),
        ("Vitesse Max Semaine", f"{max_speed:.1f} km/h"),
        ("Séances Fusionnées", f"{nb_seances} séances"),
    ]

    if weekly_rpe.get("rpe_moyen_equipe") is not None:
        kpis.append(("RPE Moyen Semaine", f"{weekly_rpe['rpe_moyen_equipe']:.1f}/10"))
    elif "RPE" in df_hebdo.columns and pd.to_numeric(df_hebdo["RPE"], errors="coerce").notna().any():
        r_mean = float(pd.to_numeric(df_hebdo["RPE"], errors="coerce").dropna().mean())
        kpis.append(("RPE Moyen Semaine", f"{r_mean:.1f}/10"))

    col_w = (273.0 - (len(kpis) - 1) * 2.25) / len(kpis)
    for idx, (label, val) in enumerate(kpis):
        x = 12 + idx * (col_w + 2.25)
        pdf.set_fill_color(248, 250, 252)
        pdf.set_draw_color(226, 232, 240)
        pdf.rect(x, 20.0, col_w, 13.0, "FD")
        pdf.set_xy(x, 21.0)
        pdf.set_font("Helvetica", "B", 7.0)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(col_w, 3.8, label.upper(), align="C")
        pdf.set_xy(x, 25.0)
        pdf.set_font("Helvetica", "B", 10.5)
        pdf.set_text_color(31, 78, 121)
        pdf.cell(col_w, 6.5, val, align="C")

    # 2. Section intermédiaire (Tableau des séances + Graphiques)
    y_mid = 35.0

    # Colonne Gauche : Tableau des séances + Graphique Répartition de charge
    pdf.set_xy(12, y_mid)
    pdf.set_font("Helvetica", "B", 8.8)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(133, 4.0, safe_latin1("1. SÉANCES ENREGISTRÉES & RÉPARTITION DE CHARGE :"), align="L")

    y_table = y_mid + 4.5
    pdf.set_xy(12, y_table)
    pdf.set_fill_color(31, 78, 121)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 7.5)
    pdf.cell(45, 4.8, "Séance / Type", 1, 0, "L", fill=True)
    pdf.cell(24, 4.8, "Date", 1, 0, "C", fill=True)
    pdf.cell(20, 4.8, "Effectif", 1, 0, "C", fill=True)
    pdf.cell(22, 4.8, "Dist. (km)", 1, 0, "C", fill=True)
    pdf.cell(22, 4.8, "H.I. (m)", 1, 1, "C", fill=True)

    pdf.set_font("Helvetica", "", 7.2)
    pdf.set_text_color(15, 23, 42)
    sessions_to_show = list(synthese_dict.items())[:4]
    for s_id, s in sessions_to_show:
        pdf.set_x(12)
        pdf.cell(45, 4.2, safe_latin1(str(s.get("day_type", ""))[:20]), 1, 0, "L")
        pdf.cell(24, 4.2, safe_latin1(format_readable_date(s.get("date", ""))), 1, 0, "C")
        pdf.cell(20, 4.2, safe_latin1(f"{s.get('num_players', 0)}"), 1, 0, "C")
        pdf.cell(22, 4.2, safe_latin1(f"{float(s.get('total_distance_km', 0.0)):.1f} km"), 1, 0, "C")
        pdf.cell(22, 4.2, safe_latin1(f"{int(round(float(s.get('total_hi_m', 0.0)))):,d}".replace(",", " ")), 1, 1, "C")

    # Graphique Répartition de charge par séance sous le tableau
    y_chart_days = pdf.get_y() + 2.0
    chart_days_h = min(46.0, max(24.0, 106.0 - y_chart_days))
    try:
        buf_days = generer_graphique_hebdo_repartition_jours_matplotlib(synthese_dict)
        pdf.image(buf_days, x=12, y=y_chart_days, w=133, h=chart_days_h)
    except Exception as e:
        print(f"Erreur graphique repartition jours PDF : {e}")

    # Colonne Droite : Graphique Charge moyenne par Ligne
    pdf.set_xy(151, y_mid)
    pdf.set_font("Helvetica", "B", 8.8)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(133, 4.0, safe_latin1("2. CHARGE MOYENNE PAR LIGNE DE POSTE (SEMAINE) :"), align="L")

    try:
        buf_groups = generer_graphique_hebdo_groupes_matplotlib(df_hebdo)
        pdf.image(buf_groups, x=151, y=y_mid + 4.5, w=134, h=66.5)
    except Exception as e:
        print(f"Erreur graphique groupes hebdo PDF : {e}")

    # 3. Section 3 : Podium Général de l'effectif complet
    y_pod_gen = 108.0
    pdf.set_xy(12, y_pod_gen)
    pdf.set_font("Helvetica", "B", 9.2)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(273, 4.0, safe_latin1("3. PODIUM GÉNÉRAL HEBDOMADAIRE — EFFECTIF COMPLET :"), align="L")

    podium_general = generer_podiums_hebdomadaires(df_hebdo)
    sub_map_gen = calculer_sous_titres_podium(df_hebdo, label_prefix="Moy. semaine")
    _render_podiums_compact_horizontal(
        pdf=pdf,
        podium_data=podium_general,
        y_start=y_pod_gen + 4.8,
        subtitle_map=sub_map_gen,
        is_large=True,
    )

    # =========================================================================
    # PAGES 2, 3 ET 4 : BILANS DÉTAILLÉS PAR LIGNE DE POSTE (STRICTEMENT 4 PAGES TOTAL)
    # Page 2 : 5 de devant
    # Page 3 : 3ème Ligne
    # Page 4 : Trois-quarts
    # =========================================================================
    groupes_rca = ["5 de devant", "3ème Ligne", "Trois-quarts"]

    for grp in groupes_rca:
        df_sub = df_hebdo[df_hebdo["Groupe"] == grp] if not df_hebdo.empty and "Groupe" in df_hebdo.columns else pd.DataFrame()
        n_players_grp = len(df_sub)

        pdf.current_subtitle = f"BILAN HEBDOMADAIRE — {grp.upper()} ({n_players_grp} JOUEUR{'S' if n_players_grp > 1 else ''})"
        pdf.add_page()
        pdf.set_y(19.5)

        # 1. Objectifs Hebdomadaires vs Réalisé Cumulé
        pdf.set_xy(12, 19.5)
        pdf.set_font("Helvetica", "B", 9.0)
        pdf.set_text_color(11, 25, 44)
        pdf.cell(273, 4.0, safe_latin1(f"1. OBJECTIFS HEBDOMADAIRES VS RÉALISÉ DU GROUPE ({nb_seances} SÉANCES CUMULÉES) :"), align="L")

        try:
            buf_jauges = generer_graphique_hebdo_jauges_matplotlib(
                grp, df_sub, nb_seances=nb_seances, objectifs_config=objectifs_config
            )
            pdf.image(buf_jauges, x=12, y=24.0, w=273, h=38.0)
        except Exception as e:
            print(f"Erreur graphique jauges hebdo PDF pour {grp} : {e}")

        # 2. Podiums de la Ligne (Format Compact Horizontal)
        y_sec_pod = 63.5
        pdf.set_xy(12, y_sec_pod)
        pdf.set_font("Helvetica", "B", 9.0)
        pdf.set_text_color(11, 25, 44)
        pdf.cell(273, 4.0, safe_latin1(f"2. PODIUMS DU GROUPE — {grp.upper()} :"), align="L")

        podium_grp = generer_podiums_hebdomadaires(df_hebdo, groupe=grp)
        sub_map_grp = calculer_sous_titres_podium(df_sub, label_prefix="Moyenne")
        _render_podiums_compact_horizontal(
            pdf=pdf,
            podium_data=podium_grp,
            y_start=y_sec_pod + 4.5,
            subtitle_map=sub_map_grp,
            is_large=False,
        )

        # 3. Tableau individuel des statistiques cumulées
        y_sec_tbl = 109.5
        pdf.set_xy(12, y_sec_tbl)
        pdf.set_font("Helvetica", "B", 9.0)
        pdf.set_text_color(11, 25, 44)
        pdf.cell(273, 4.0, safe_latin1("3. STATISTIQUES INDIVIDUELLES CUMULÉES SUR LA SEMAINE :"), align="L")

        # En-tête du tableau
        y_tbl_hdr = y_sec_tbl + 4.5
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

        # Lignes du tableau
        if df_sub.empty:
            pdf.set_x(12)
            pdf.set_fill_color(248, 250, 252)
            pdf.set_font("Helvetica", "I", 8.0)
            pdf.set_text_color(100, 116, 139)
            pdf.cell(273, 7.0, safe_latin1("Aucun joueur enregistré pour ce groupe cette semaine."), 1, 1, "C", fill=True)
        else:
            avail_h = 196.0 - (y_tbl_hdr + 5.0)  # ~ 72.5 mm
            row_h = min(4.8, max(3.5, avail_h / max(n_players_grp, 1)))
            font_sz = 7.5 if row_h >= 4.5 else 6.5
            pdf.set_font("Helvetica", "", font_sz)
            pdf.set_text_color(15, 23, 42)
            fill_row = False

            def _safe_float(v: Any, default: float = 0.0) -> float:
                try:
                    if v is None or pd.isna(v):
                        return default
                    return float(v)
                except (ValueError, TypeError):
                    return default

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

    # =========================================================================
    # PAGE 5 : ANALYSE COMPARATIVE HISTORIQUE (OPTIONNELLE)
    # =========================================================================
    if inclure_comparatif:
        items_comp = recuperer_historique_semaines(
            year_curr=year,
            week_curr=week_num,
            df_curr_hebdo=df_hebdo,
            n_max=nb_comparatif,
        )
        pdf.current_subtitle = "ANALYSE COMPARATIVE HISTORIQUE — MULTI-SEMAINES"
        pdf.add_page()
        render_comparatif_historique_pdf_page(
            pdf=pdf,
            items_comparatif=items_comp,
            type_label="Semaine",
        )

    return bytes(pdf.output())
