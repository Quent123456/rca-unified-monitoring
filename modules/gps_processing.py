"""Module de traitement et d'analyse des données GPS Catapult One pour le Rugby Club Amiénois (RCA).

Gère :
1. Le chargement intelligent des fichiers Catapult One (.csv et .xlsx) avec détection d'en-tête.
2. Le mapping des boîtiers GPS vers l'effectif officiel RCA avec persistance JSON.
3. L'extraction et le filtrage des Splits / Ateliers d'entraînement.
4. Le calcul des métriques rugby : Distance, HSR (Z3), VHSR (Z4+Z5), Accélérations, Décélérations, PlayerLoad, m/min.
5. La jointure ultra-robuste avec les données RPE de suivi de charge (Nom + Prénom + Date, skipna=True, gestion 'NR').
6. L'injection dans le modèle Excel officiel STAPS de la préparatrice physique.
"""

from __future__ import annotations

import base64
import copy
import datetime
import io
import json
import os
import re
import textwrap
import urllib.parse
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

try:
    from modules import photo_manager, weekly_memory, rpe_connector
except ImportError:
    import photo_manager
    import weekly_memory
    import rpe_connector

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
def format_readable_date(val: Any) -> str:
    """
    Convertit n'importe quelle valeur de date (numéro de série Excel ex: 46315, 46316, 46318,
    timestamp, chaîne ISO, ou datetime) en date lisible au format français JJ/MM/AAAA.
    Définie directement dans app.py pour immuniser l'application contre les problèmes de cache de modules Streamlit Cloud.
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


# Rétro-compatibilité : s'assurer que weekly_memory expose aussi format_readable_date
if not hasattr(weekly_memory, "format_readable_date"):
    weekly_memory.format_readable_date = format_readable_date


# ==============================================================================
# 0. EFFECTIF SÉNIORS RCA & CONFIGURATION ATHLÉTIQUE
# ==============================================================================

TEMPLATE_FILENAME = "Modele_Vierge_STAPS.xlsx"
OUTPUT_DOWNLOAD_FILENAME = "Synthese_Complete_Seance.xlsx"

# Dossiers contenant les photos des joueurs (assets/players_photos/ normalisé)
ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
PLAYERS_PHOTOS_DIR = os.path.join(ASSETS_DIR, "players_photos")
os.makedirs(PLAYERS_PHOTOS_DIR, exist_ok=True)
PHOTOS_DIR = PLAYERS_PHOTOS_DIR
DEFAULT_AVATAR = os.path.join(PLAYERS_PHOTOS_DIR, "avatar_par_defaut.png")

# Vérification et extraction automatique des photos depuis le trombinoscope PDF si besoin
try:
    photo_manager.ensure_photos_extracted()
except Exception:
    pass

# Palette de couleurs officielles Rugby Club Amiénois / Style Canva
COLOR_OBJECTIF = "#2E5A88"  # Bleu roi / marine RCA
COLOR_REALISE = "#A30000"   # Rouge / Bordeaux RCA
COLOR_TEXT_MAIN = "#1F4E79"
COLOR_MUTED = "#555555"

# ------------------------------------------------------------------------------
# 1. BASE DE DONNÉES EFFECTIF SÉNIORS DU RUGBY CLUB AMIÉNOIS (EXTRACTION PDF)
# ------------------------------------------------------------------------------
joueurs_data: Dict[str, str] = {
    # Piliers (5 de devant)
    "SALUA Landry": "Pilier",
    "SALUA Noam": "Pilier",
    "FRION Quentin": "Pilier",
    "COZIC Adrien": "Pilier",
    "PECQUERY Thomas": "Pilier",
    "TRAMON Noam": "Pilier",
    "FOURNET Giovanni": "Pilier",
    "MAZE Nolan": "Pilier",
    "BICHARI Jérémy": "Pilier",
    "LEFEVER Thibault": "Pilier",
    "ROGIER Arnaud": "Pilier",
    "VASELLI Sébastien": "Pilier",
    "TANTER Edouard": "Pilier",
    "DAULT Bastien": "Pilier",
    "SCHNEERBERGER Julien": "Pilier",
    "DUGROPREZ Julien": "Pilier",
    # Talonneurs (5 de devant)
    "LESUEUR Jean-Étienne": "Talonneur",
    "Jean-Étienne ( TIJ)": "Talonneur",
    "GENERMONT Roméo": "Talonneur",
    "GUIBERT Samuel": "Talonneur",
    "GIET Axel": "Talonneur",
    "D. Anthymme": "Talonneur",
    # 2èmes lignes (5 de devant)
    "WAEYAERT Thomas": "2ème ligne",
    "GRENET Paul": "2ème ligne",
    "BOIDIN Donovan": "2ème ligne",
    "VDH Simon": "2ème ligne",
    "DENIER Gabin": "2ème ligne",
    "Gabin D.": "2ème ligne",
    "ROY Etienne": "2ème ligne",
    "FOUQUET Victorien": "2ème ligne",
    # 3èmes lignes (Alignement Préparatrice Physique RCA : SOTILLO Alejo et SAINTOMER Lucas inclus)
    "SOTILLO Alejo": "3ème ligne",
    "SAINTOMER Lucas": "3ème ligne",
    "SAITOMER Lucas": "3ème ligne",
    "SOTILLO Lucas": "3ème ligne",
    "VASELLI Enzo": "3ème ligne",
    "ORION Lucas": "3ème ligne",
    "OLIER Lucas": "3ème ligne",
    "CARPENTIER Simon": "3ème ligne",
    "LAFEUILLE Emilien": "3ème ligne",
    "MARTIN Benjamin": "3ème ligne",
    "MONTI Ciro": "3ème ligne",
    "WIOTTE Louis": "3ème ligne",
    "GENTIEN Tim": "3ème ligne",
    "CARLIER Louis": "3ème ligne",
    "COSSARD Léandre": "3ème ligne",
    "SAINTOMER Louis": "3ème ligne",
    "AUDIERE Adrien": "3ème ligne",
    "BINNINGER Alban": "3ème ligne",
    "BLANC Jules": "3ème ligne",
    "KAZUMBA Camille": "3ème ligne",
    # Demis de mêlée (Trois-quarts)
    "TESSON Gabriel": "Demi de mêlée",
    "BAMIERE Clément": "Demi de mêlée",
    "BAM'S": "Demi de mêlée",
    "SANTONI Stephan": "Demi de mêlée",
    "BOEUF Elliot": "Demi de mêlée",
    "BŒUF Eliot": "Demi de mêlée",
    "MAUGARD Nicolas": "Demi de mêlée",
    "BOUMENDIL Antoine": "Demi de mêlée",
    # Demis d'ouverture (Trois-quarts)
    "HARRY Max": "Demi d'ouverture",
    "HARRY Maxence": "Demi d'ouverture",
    "CARPENTIER Maël": "Demi d'ouverture",
    "ULMUCK Victor": "Demi d'ouverture",
    "USUBILLAGA Victor": "Demi d'ouverture",
    "GOMES Louis": "Demi d'ouverture",
    "SUEUR Xavier": "Demi d'ouverture",
    # Centres (Trois-quarts)
    "PRÉVOTÉ Oliver": "Centre",
    "PREVOTE Olivier": "Centre",
    "PELESEUMA Ope": "Centre",
    "OPE": "Centre",
    "TREMBLAY Brieuc": "Centre",
    "DE BRYUN Octave": "Centre",
    "DE BRUYN Octave": "Centre",
    "VENIN Benoît": "Centre",
    "HURTEL Oscar": "Centre",
    "Oscar": "Centre",
    "DELAVALLÉE Quentin": "Centre",
    "MANDIVENGA Brendon": "Centre",
    # Ailiers (Trois-quarts - dont Jules D. / DELACROIX Jules)
    "DELACROIX Jules": "Ailier",
    "CARPENTIER Joseph": "Ailier",
    "SALUA Maka": "Ailier",
    "SALUA Makatea": "Ailier",
    "SELTZER Antoine": "Ailier",
    "BATON Léonard": "Ailier",
    "BATTON Léonard": "Ailier",
    "MARZIANO Nathan": "Ailier",
    "SRIDI Youness": "Ailier",
    "LEROUX Mathis": "Ailier",
    "GANE Thomas": "Ailier",
    "DELATTRE Jacques": "Ailier",
    # Arrières (Trois-quarts)
    "LESPINASSE Lucien": "Arrière",
    "PRÉVOST Thomas": "Arrière",
    "PREVOST Thomas": "Arrière",
    "MOINE Arthur": "Arrière",
    "VILAIN Beryl": "Arrière",
}

# ------------------------------------------------------------------------------
# 1.1 GESTION DE L'EFFECTIF PERSONNALISÉ & PERSISTANCE DES NOUVEAUX JOUEURS
# ------------------------------------------------------------------------------

CONFIG_JOUEURS_CUSTOM_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_rca_joueurs_custom.json")

# Cache mémoire pour synchronisation rapide des joueurs personnalisés
JOUEURS_CUSTOM_CACHE: Dict[str, Dict[str, Any]] = {}


def charger_joueurs_personnalises() -> Dict[str, Dict[str, Any]]:
    """
    Charge les joueurs personnalisés créés dans le logiciel depuis config_rca_joueurs_custom.json.
    Format : { "Nom Prénom": { "poste": "...", "groupe": "...", "date_ajout": "..." } }
    """
    global JOUEURS_CUSTOM_CACHE
    if os.path.exists(CONFIG_JOUEURS_CUSTOM_FILE):
        try:
            with open(CONFIG_JOUEURS_CUSTOM_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    JOUEURS_CUSTOM_CACHE = data
                    return data
        except Exception:
            pass
    JOUEURS_CUSTOM_CACHE = {}
    return {}


def synchroniser_effectif_avec_custom() -> None:
    """Injecte les joueurs personnalisés dans le dictionnaire global joueurs_data."""
    custom = charger_joueurs_personnalises()
    for nom, info in custom.items():
        if isinstance(info, dict):
            joueurs_data[nom] = info.get("poste", "Non spécifié")
        elif isinstance(info, str):
            joueurs_data[nom] = info


# Synchronisation initiale au démarrage du module
synchroniser_effectif_avec_custom()


def sauvegarder_joueur_personnalise(
    nom: str,
    poste: str,
    groupe: Optional[str] = None,
    photo_bytes: Optional[bytes] = None,
    photo_ext: str = ".jpg",
) -> Tuple[bool, str]:
    """
    Enregistre un nouveau joueur dans config_rca_joueurs_custom.json,
    met à jour joueurs_data et enregistre la photo dans photos_joueurs/ si présente.
    """
    nom_clean = nom.strip()
    if not nom_clean:
        return False, "Le nom du joueur ne peut pas être vide."
    if nom_clean in ["— Sélectionner un joueur —", "➕ [Créer un nouveau joueur...]", "-", ""]:
        return False, "Le nom saisi n'est pas valide."

    poste_clean = poste.strip() if poste else "Inconnu"
    grp_final = (
        groupe
        if groupe in ["5 de devant", "3ème Ligne", "Trois-quarts"]
        else mapper_poste_vers_groupe_rca(poste_clean, nom_clean)
    )

    # Enregistrement de la photo si fournie dans assets/players_photos/
    if photo_bytes:
        photo_manager.save_manual_player_photo(nom_clean, photo_bytes, ext=photo_ext)

    custom = charger_joueurs_personnalises()
    custom[nom_clean] = {
        "poste": poste_clean,
        "groupe": grp_final,
        "date_ajout": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }

    try:
        with open(CONFIG_JOUEURS_CUSTOM_FILE, "w", encoding="utf-8") as f:
            json.dump(custom, f, ensure_ascii=False, indent=2)
        JOUEURS_CUSTOM_CACHE.clear()
        JOUEURS_CUSTOM_CACHE.update(custom)
        joueurs_data[nom_clean] = poste_clean

        try:
            try:

                import github_sync

            except (ImportError, ModuleNotFoundError):

                from modules import github_sync
            github_sync.push_file_to_github(
                CONFIG_JOUEURS_CUSTOM_FILE,
                commit_message=f"data: ajout/mise à jour du joueur {nom_clean}"
            )
            saved_p_path = photo_manager.get_player_photo_path(nom_clean)
            if saved_p_path and os.path.exists(saved_p_path):
                github_sync.push_file_to_github(
                    saved_p_path,
                    commit_message=f"assets: photo du joueur {nom_clean}"
                )
        except Exception:
            pass

        return True, f"Joueur « {nom_clean} » enregistré avec succès !"
    except Exception as e:
        return False, f"Erreur lors de l'enregistrement : {str(e)}"


def supprimer_joueur_personnalise(nom: str) -> Tuple[bool, str]:
    """Supprime un joueur personnalisé de la configuration, des mappings GPS et nettoie sa photo."""
    custom = charger_joueurs_personnalises()
    if nom in custom:
        del custom[nom]
        try:
            with open(CONFIG_JOUEURS_CUSTOM_FILE, "w", encoding="utf-8") as f:
                json.dump(custom, f, ensure_ascii=False, indent=2)
            JOUEURS_CUSTOM_CACHE.clear()
            JOUEURS_CUSTOM_CACHE.update(custom)
            if nom in joueurs_data:
                del joueurs_data[nom]

            # Nettoyage et suppression propre de la photo associée dans assets/players_photos/
            photo_path_to_del = photo_manager.get_player_photo_path(nom)
            photo_manager.delete_player_photo(nom)

            try:
                try:

                    import github_sync

                except (ImportError, ModuleNotFoundError):

                    from modules import github_sync
                github_sync.push_file_to_github(
                    CONFIG_JOUEURS_CUSTOM_FILE,
                    commit_message=f"data: suppression du joueur {nom}"
                )
                if photo_path_to_del:
                    github_sync.delete_file_from_github(
                        photo_path_to_del,
                        commit_message=f"assets: suppression photo {nom}"
                    )
            except Exception:
                pass

            # Nettoyage de toutes les présences du joueur dans la mémoire hebdomadaire (règle sans orphelin)
            try:
                weekly_memory.delete_player_records(nom)
            except Exception:
                pass

            # Nettoyage du mapping GPS si ce joueur y était associé
            current_gps_map = charger_mapping_gps_persistant()
            changed = False
            for k, v in list(current_gps_map.items()):
                if v == nom:
                    del current_gps_map[k]
                    changed = True
            if changed:
                sauvegarder_mapping_gps_persistant(current_gps_map)

            return True, f"Joueur « {nom} » supprimé de l'effectif avec succès."
        except Exception as e:
            return False, f"Erreur lors de la suppression : {str(e)}"
    return False, "Joueur introuvable dans l'effectif personnalisé."


def mapper_poste_vers_groupe_rca(poste_detaille: str, nom_joueur: str = "") -> str:
    """
    Convertit automatiquement et de façon robuste un poste détaillé du rugby en l'un
    des 3 grands groupes d'analyse officiels de la préparatrice physique du RCA :
    1. '5 de devant' : Pilier, Pilier droit, Pilier gauche, Talonneur, 2ème ligne, Deuxième ligne, Première ligne, 5 de devant.
    2. '3ème Ligne' : 3ème ligne, Troisième ligne, 3ème ligne aile, 3ème ligne centre, Flanker, Numéro 8.
    3. 'Trois-quarts' : Demis (mêlée/ouverture/9/10), Ouvrière, Demis, Centre, Ailier, Arrière, Trois-quarts, 3/4.

    Prend également en compte les alignements prioritaires demandés par la préparatrice physique :
    - Lucas S. et Alejo S. rattachés en '3ème Ligne'.
    - Bastien D., Thomas P., Edouard T., Jean-Étienne L. en '5 de devant'.
    - Enzo V., Lucas O., Simon C. en '3ème Ligne'.
    - Jules D. (DELACROIX Jules), Elliot B., Maël C., Victor U., Olivier P., Ope P., Maka S., Arthur M., Octave D., Lucien L. en 'Trois-quarts'.
    """
    nom_clean = str(nom_joueur).strip().lower() if nom_joueur else ""

    # 0. Vérification si joueur personnalisé avec un groupe explicite choisi lors de la création
    if nom_clean:
        for c_nom, c_val in JOUEURS_CUSTOM_CACHE.items():
            if c_nom.strip().lower() == nom_clean:
                if isinstance(c_val, dict) and c_val.get("groupe") in ["5 de devant", "3ème Ligne", "Trois-quarts"]:
                    return c_val["groupe"]

    # 1. Règle d'alignement prioritaire par joueur spécifique (Staff RCA)
    if nom_clean:
        # Lucas S. & Alejo S. -> 3ème Ligne (selon méthodologie préparatrice physique RCA)
        if "sotillo" in nom_clean:
            return "3ème Ligne"
        if ("saintomer" in nom_clean or "saitomer" in nom_clean) and "lucas" in nom_clean:
            return "3ème Ligne"

        # 5 de devant spécifiques
        if any(k in nom_clean for k in [
            "dault bastien", "bastien d", "dault",
            "pecquery thomas", "thomas p", "pecquery",
            "tanter edouard", "edouard t", "tanter",
            "lesueur jean-étienne", "lesueur jean-etienne", "jean-étienne l", "jean-etienne l", "lesueur", "tij",
            "cozic adrien", "cozic"
        ]):
            return "5 de devant"

        # 3ème Ligne spécifiques
        if any(k in nom_clean for k in [
            "vaselli enzo", "enzo v",
            "orion lucas", "olier lucas", "lucas o",
            "carpentier simon", "simon c",
            "blanc jules", "lafeuille emilien", "monti ciro"
        ]):
            return "3ème Ligne"

        # Trois-quarts spécifiques
        if any(k in nom_clean for k in [
            "delacroix jules", "jules d", "delacroix",
            "boeuf elliot", "bœuf eliot", "boeuf eliot", "bœuf elliot", "elliot b", "eliot b",
            "carpentier maël", "carpentier mael", "maël c", "mael c",
            "ulmuck victor", "usubillaga victor", "victor u", "ulmuck", "usubillaga",
            "prévoté oliver", "prevote olivier", "olivier p", "oliver p", "prevote", "prévoté",
            "peleseuma ope", "ope p", "ope",
            "salua maka", "salua makatea", "maka s",
            "moine arthur", "arthur m",
            "de bruyn octave", "de bryun octave", "octave d",
            "lespinasse lucien", "lucien l", "lespinasse",
            "bam's", "bamiere", "harry max", "harry maxence"
        ]):
            return "Trois-quarts"

    # 2. Conversion automatique par poste détaillé
    p_clean = str(poste_detaille).strip().lower()
    if not p_clean:
        return "Trois-quarts"

    # 5 de devant
    # Pilier, Pilier droit, Pilier gauche, Talonneur, 2ème ligne, Deuxième ligne, Première ligne, 5 de devant
    if any(k in p_clean for k in [
        "pilier", "prop", "talon", "hooker",
        "2ème ligne", "2eme ligne", "deuxième ligne", "deuxieme ligne", "2nd ligne", "2e ligne", "lock",
        "première ligne", "premiere ligne", "1ère ligne", "1ere ligne",
        "5 de devant", "cinq de devant", "front row", "tight five"
    ]):
        return "5 de devant"

    # 3ème Ligne
    # 3ème ligne, Troisième ligne, 3ème ligne aile, 3ème ligne centre, Flanker, Numéro 8
    if any(k in p_clean for k in [
        "3ème ligne", "3eme ligne", "troisième ligne", "troisieme ligne", "3e ligne",
        "3ème ligne aile", "3eme ligne aile", "3ème ligne centre", "3eme ligne centre",
        "flanker", "flank", "troisième ligne aile", "troisième ligne centre",
        "numéro 8", "numero 8", "num 8", "n°8", "no8", "n8",
        "third row", "back row", "loose forward"
    ]) or p_clean == "8":
        return "3ème Ligne"

    # Trois-quarts
    # Demi de mêlée (9), Demi d'ouverture (10), Ouvrière, Demis, Centre, Ailier, Arrière, Trois-quarts, 3/4
    if any(k in p_clean for k in [
        "demi de mêlée", "demi de melee", "demi melee", "scrum half", "scrum-half",
        "demi d'ouverture", "demi d ouverture", "demi ouverture", "fly half", "fly-half", "ouvreur", "ouvrière", "ouvriere",
        "demi", "demis",
        "centre", "center", "premier centre", "second centre",
        "ailier", "wing", "winger",
        "arrière", "arriere", "fullback", "full back",
        "trois-quarts", "trois quarts", "trois-quart", "trois quart", "3/4", "backs"
    ]) or p_clean in ["9", "10", "11", "12", "13", "14", "15"]:
        return "Trois-quarts"

    # 3. Règles heuristiques de repli
    if "5" in p_clean or "devant" in p_clean:
        return "5 de devant"
    if "3" in p_clean:
        return "3ème Ligne"

    return "Trois-quarts"


# Cartographie Poste -> Groupe athlétique officiel RCA (3 catégories)
POSTE_VERS_GROUPE: Dict[str, str] = {
    # 5 de devant
    "Pilier": "5 de devant",
    "Pilier droit": "5 de devant",
    "Pilier gauche": "5 de devant",
    "Talonneur": "5 de devant",
    "2ème ligne": "5 de devant",
    "Deuxième ligne": "5 de devant",
    "Première ligne": "5 de devant",
    "5 de devant": "5 de devant",
    # 3ème Ligne
    "3ème ligne": "3ème Ligne",
    "Troisième ligne": "3ème Ligne",
    "3ème ligne aile": "3ème Ligne",
    "3ème ligne centre": "3ème Ligne",
    "Flanker": "3ème Ligne",
    "Numéro 8": "3ème Ligne",
    # Trois-quarts
    "Demi de mêlée": "Trois-quarts",
    "Demi d'ouverture": "Trois-quarts",
    "Ouvrière": "Trois-quarts",
    "Demis": "Trois-quarts",
    "Centre": "Trois-quarts",
    "Centres": "Trois-quarts",
    "Ailier": "Trois-quarts",
    "Ailiers": "Trois-quarts",
    "Arrière": "Trois-quarts",
    "Arrières": "Trois-quarts",
    "Trois-quarts": "Trois-quarts",
}

# Objectifs athlétiques de référence par groupe (Moyenne individuelle cible par séance)
OBJECTIFS_CONFIG: Dict[str, Dict[str, float]] = {
    "5 de devant": {
        "Objectif Distance (m)": 4500.0,
        "Obj. HSR": 250.0,
        "Obj. VHSR": 50.0,
        "Obj. ACC/DEC": 35.0,
    },
    "3ème Ligne": {
        "Objectif Distance (m)": 5500.0,
        "Obj. HSR": 450.0,
        "Obj. VHSR": 100.0,
        "Obj. ACC/DEC": 45.0,
    },
    "Trois-quarts": {
        "Objectif Distance (m)": 6200.0,
        "Obj. HSR": 700.0,
        "Obj. VHSR": 250.0,
        "Obj. ACC/DEC": 50.0,
    },
    "Tous les groupes": {
        "Objectif Distance (m)": 5400.0,
        "Obj. HSR": 460.0,
        "Obj. VHSR": 130.0,
        "Obj. ACC/DEC": 43.0,
    },
}


def get_active_objectifs_config() -> Dict[str, Dict[str, float]]:
    """
    Récupère la configuration des objectifs actuellement active en session Streamlit,
    ou charge la configuration sauvegardée sur disque dans data/targets_config.json.
    """
    if "objectifs_config" in st.session_state and isinstance(st.session_state["objectifs_config"], dict):
        return st.session_state["objectifs_config"]

    cfg = weekly_memory.charger_objectifs_config()
    st.session_state["objectifs_config"] = cfg
    return cfg


def render_config_objectifs_ui() -> Dict[str, Dict[str, float]]:
    """
    Panneau interactif de configuration manuelle des objectifs de performance par groupe de poste.
    Permet :
    - De choisir un préréglage (Mardi, Mercredi, Vendredi, Match, Standard).
    - De modifier manuellement les 4 métriques clés pour les 3 groupes (5 de devant, 3ème Ligne, Trois-quarts).
    - De sauvegarder comme configuration par défaut dans data/targets_config.json.
    - Met à jour immédiatement st.session_state["objectifs_config"].
    """
    with st.expander("🎯 Paramétrer les Objectifs par Poste (5 de devant, 3ème Ligne, Trois-quarts)", expanded=False):
        st.markdown(
            "<div style='font-size:0.88rem; color:#475569; margin-bottom:12px;'>"
            "Définissez ici les valeurs cibles (objectifs) individuelles moyennes par séance. "
            "Ces cibles s'appliquent instantanément aux jauges de comparaison du Dashboard, "
            "au calcul du <b>% d'accomplissement</b> et au <b>Rapport PDF officiel</b>."
            "</div>",
            unsafe_allow_html=True,
        )

        current_cfg = get_active_objectifs_config()

        # 1. Préréglages et actions
        col_preset, col_btn_apply, col_btn_save = st.columns([3, 1.5, 2.2])
        with col_preset:
            preset_names = list(getattr(weekly_memory, "PREREGLAGES_OBJECTIFS", {}).keys())
            selected_preset = st.selectbox(
                "⚡ Préréglages par type de séance :",
                options=["-- Conserver les réglages actuels --"] + preset_names,
                index=0,
                key="select_targets_preset",
                help="Charge instantanément une grille d'objectifs adaptée au thème du jour.",
            )
        with col_btn_apply:
            st.write("")
            st.write("")
            if st.button("🔄 Appliquer", key="btn_apply_target_preset", use_container_width=True, help="Charge les valeurs de ce préréglage"):
                prereglages = getattr(weekly_memory, "PREREGLAGES_OBJECTIFS", {})
                if selected_preset in prereglages:
                    new_cfg = copy.deepcopy(prereglages[selected_preset])
                    new_cfg["Tous les groupes"] = weekly_memory._calculer_moyenne_tous_groupes(new_cfg)
                    st.session_state["objectifs_config"] = new_cfg
                    st.toast(f"Préréglage « {selected_preset} » appliqué !")
                    st.rerun()

        with col_btn_save:
            st.write("")
            st.write("")
            if st.button("💾 Sauvegarder par défaut", key="btn_save_target_defaults", use_container_width=True, help="Enregistre la configuration actuelle comme défaut permanent"):
                cfg_to_save = st.session_state.get("objectifs_config", current_cfg)
                weekly_memory.sauvegarder_objectifs_config(cfg_to_save)
                st.toast("✅ Objectifs sauvegardés par défaut dans data/targets_config.json !")

        # 2. Grille de 3 colonnes pour les 3 groupes de postes
        col_avants, col_3eme, col_arrieres = st.columns(3)

        cfg_5dev = current_cfg.get("5 de devant", {})
        cfg_3eme = current_cfg.get("3ème Ligne", {})
        cfg_tq = current_cfg.get("Trois-quarts", {})

        updated_cfg: Dict[str, Dict[str, float]] = {}

        with col_avants:
            st.markdown(
                "<div style='background-color:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:8px 12px; margin-bottom:10px;'>"
                "<b style='color:#1D4ED8; font-size:0.95rem;'>🏉 5 DE DEVANT</b><br/>"
                "<span style='font-size:0.75rem; color:#64748B;'>Piliers • Talonneurs • 2èmes lignes</span></div>",
                unsafe_allow_html=True,
            )
            v_dist_5 = st.number_input("Volume Total (m) :", min_value=500.0, max_value=12000.0, step=100.0, value=float(cfg_5dev.get("Objectif Distance (m)", 4500.0)), key="in_tgt_5dev_dist")
            v_hsr_5 = st.number_input("Distance HSR (>18 km/h) (m) :", min_value=0.0, max_value=3000.0, step=20.0, value=float(cfg_5dev.get("Obj. HSR", 250.0)), key="in_tgt_5dev_hsr")
            v_vhsr_5 = st.number_input("Distance VHSR (>21 km/h) (m) :", min_value=0.0, max_value=1500.0, step=10.0, value=float(cfg_5dev.get("Obj. VHSR", 50.0)), key="in_tgt_5dev_vhsr")
            v_acc_5 = st.number_input("ACC / DEC (nb) :", min_value=0.0, max_value=200.0, step=1.0, value=float(cfg_5dev.get("Obj. ACC/DEC", 35.0)), key="in_tgt_5dev_acc")
            updated_cfg["5 de devant"] = {
                "Objectif Distance (m)": v_dist_5,
                "Obj. HSR": v_hsr_5,
                "Obj. VHSR": v_vhsr_5,
                "Obj. ACC/DEC": v_acc_5,
            }

        with col_3eme:
            st.markdown(
                "<div style='background-color:#FEF2F2; border:1px solid #FECACA; border-radius:8px; padding:8px 12px; margin-bottom:10px;'>"
                "<b style='color:#B91C1C; font-size:0.95rem;'>🏉 3ÈME LIGNE</b><br/>"
                "<span style='font-size:0.75rem; color:#64748B;'>Flankers • N°8</span></div>",
                unsafe_allow_html=True,
            )
            v_dist_3 = st.number_input("Volume Total (m) :", min_value=500.0, max_value=12000.0, step=100.0, value=float(cfg_3eme.get("Objectif Distance (m)", 5500.0)), key="in_tgt_3eme_dist")
            v_hsr_3 = st.number_input("Distance HSR (>18 km/h) (m) :", min_value=0.0, max_value=3000.0, step=20.0, value=float(cfg_3eme.get("Obj. HSR", 450.0)), key="in_tgt_3eme_hsr")
            v_vhsr_3 = st.number_input("Distance VHSR (>21 km/h) (m) :", min_value=0.0, max_value=1500.0, step=10.0, value=float(cfg_3eme.get("Obj. VHSR", 100.0)), key="in_tgt_3eme_vhsr")
            v_acc_3 = st.number_input("ACC / DEC (nb) :", min_value=0.0, max_value=200.0, step=1.0, value=float(cfg_3eme.get("Obj. ACC/DEC", 45.0)), key="in_tgt_3eme_acc")
            updated_cfg["3ème Ligne"] = {
                "Objectif Distance (m)": v_dist_3,
                "Obj. HSR": v_hsr_3,
                "Obj. VHSR": v_vhsr_3,
                "Obj. ACC/DEC": v_acc_3,
            }

        with col_arrieres:
            st.markdown(
                "<div style='background-color:#F0FDF4; border:1px solid #BBF7D0; border-radius:8px; padding:8px 12px; margin-bottom:10px;'>"
                "<b style='color:#15803D; font-size:0.95rem;'>🏉 TROIS-QUARTS</b><br/>"
                "<span style='font-size:0.75rem; color:#64748B;'>Demis • Centres • Ailiers • Arrières</span></div>",
                unsafe_allow_html=True,
            )
            v_dist_tq = st.number_input("Volume Total (m) :", min_value=500.0, max_value=12000.0, step=100.0, value=float(cfg_tq.get("Objectif Distance (m)", 6200.0)), key="in_tgt_tq_dist")
            v_hsr_tq = st.number_input("Distance HSR (>18 km/h) (m) :", min_value=0.0, max_value=3000.0, step=20.0, value=float(cfg_tq.get("Obj. HSR", 700.0)), key="in_tgt_tq_hsr")
            v_vhsr_tq = st.number_input("Distance VHSR (>21 km/h) (m) :", min_value=0.0, max_value=1500.0, step=10.0, value=float(cfg_tq.get("Obj. VHSR", 250.0)), key="in_tgt_tq_vhsr")
            v_acc_tq = st.number_input("ACC / DEC (nb) :", min_value=0.0, max_value=200.0, step=1.0, value=float(cfg_tq.get("Obj. ACC/DEC", 50.0)), key="in_tgt_tq_acc")
            updated_cfg["Trois-quarts"] = {
                "Objectif Distance (m)": v_dist_tq,
                "Obj. HSR": v_hsr_tq,
                "Obj. VHSR": v_vhsr_tq,
                "Obj. ACC/DEC": v_acc_tq,
            }

        updated_cfg["Tous les groupes"] = weekly_memory._calculer_moyenne_tous_groupes(updated_cfg)
        st.session_state["objectifs_config"] = updated_cfg
        return updated_cfg


FINAL_COLUMNS: List[str] = [
    'Nom ; Prénom',
    'Temps de jeu ',
    'Distance (km)',
    ' Distance haute intensité (m)',
    'Top Speed (km/h)',
    'Distance Per Min (m/min)',
    'Distance in Speed Zone 1  (km)',
    'Distance in Speed Zone 2  (km)',
    'Distance in Speed Zone 3  (km)',
    'Distance in Speed Zone 4  (km)',
    'Distance in Speed Zone 5  (km)',
    'Max Deceleration (m/s/s)',
    'Max Acceleration (m/s/s)',
    'Accelerations Zone Count: 3 - 4 m/s/s',
    'Accelerations Zone Count: > 4 m/s/s',
    'Deceleration Zone Count: 3 - 4 m/s/s',
    'Deceleration Zone Count: > 4 m/s/s',
    'Accel/Dec',
    '%Marche ',
    '%course',
    '%HSR ',
    '%VHSR',
    'HSR m/min',
]

# Données de simulation avec joueurs réels du RCA
MOCK_DEMO_DATA: List[Dict[str, Any]] = [
    {"Player Name": "BAMIERE Clément", "Position": "Demi de mêlée", "Distance (km)": 6.3, "Zone 4 (km)": 0.48, "Zone 5 (km)": 0.24, "Top Speed": 32.4, "Duration": "01:22:00", "Z1 (km)": 2.5, "Z2 (km)": 2.2, "Z3 (km)": 0.88, "Acc34": 18, "Acc4": 8, "Dec34": 14, "Dec4": 6},
    {"Player Name": "PRÉVOST Thomas", "Position": "Arrière", "Distance (km)": 6.6, "Zone 4 (km)": 0.52, "Zone 5 (km)": 0.28, "Top Speed": 31.8, "Duration": "01:22:00", "Z1 (km)": 2.6, "Z2 (km)": 2.3, "Z3 (km)": 0.90, "Acc34": 20, "Acc4": 7, "Dec34": 15, "Dec4": 5},
    {"Player Name": "HARRY Max", "Position": "Demi d'ouverture", "Distance (km)": 6.0, "Zone 4 (km)": 0.44, "Zone 5 (km)": 0.20, "Top Speed": 31.1, "Duration": "01:22:00", "Z1 (km)": 2.4, "Z2 (km)": 2.1, "Z3 (km)": 0.86, "Acc34": 16, "Acc4": 6, "Dec34": 13, "Dec4": 4},
    {"Player Name": "VASELLI Enzo", "Position": "3ème ligne", "Distance (km)": 5.6, "Zone 4 (km)": 0.41, "Zone 5 (km)": 0.11, "Top Speed": 29.5, "Duration": "01:22:00", "Z1 (km)": 2.6, "Z2 (km)": 1.8, "Z3 (km)": 0.68, "Acc34": 16, "Acc4": 6, "Dec34": 13, "Dec4": 5},
    {"Player Name": "MARTIN Benjamin", "Position": "3ème ligne", "Distance (km)": 5.4, "Zone 4 (km)": 0.38, "Zone 5 (km)": 0.09, "Top Speed": 28.8, "Duration": "01:22:00", "Z1 (km)": 2.7, "Z2 (km)": 1.7, "Z3 (km)": 0.63, "Acc34": 15, "Acc4": 5, "Dec34": 12, "Dec4": 4},
    {"Player Name": "MONTI Ciro", "Position": "3ème ligne", "Distance (km)": 5.5, "Zone 4 (km)": 0.39, "Zone 5 (km)": 0.10, "Top Speed": 29.0, "Duration": "01:22:00", "Z1 (km)": 2.7, "Z2 (km)": 1.7, "Z3 (km)": 0.61, "Acc34": 15, "Acc4": 5, "Dec34": 12, "Dec4": 4},
    {"Player Name": "LESUEUR Jean-Étienne", "Position": "Talonneur", "Distance (km)": 4.6, "Zone 4 (km)": 0.17, "Zone 5 (km)": 0.04, "Top Speed": 26.1, "Duration": "01:18:00", "Z1 (km)": 2.9, "Z2 (km)": 1.3, "Z3 (km)": 0.23, "Acc34": 14, "Acc4": 4, "Dec34": 11, "Dec4": 3},
    {"Player Name": "COZIC Adrien", "Position": "Pilier", "Distance (km)": 4.4, "Zone 4 (km)": 0.15, "Zone 5 (km)": 0.03, "Top Speed": 25.2, "Duration": "01:15:00", "Z1 (km)": 2.8, "Z2 (km)": 1.2, "Z3 (km)": 0.22, "Acc34": 12, "Acc4": 3, "Dec34": 10, "Dec4": 2},
    {"Player Name": "SALUA Landry", "Position": "Pilier", "Distance (km)": 4.2, "Zone 4 (km)": 0.12, "Zone 5 (km)": 0.02, "Top Speed": 24.5, "Duration": "01:15:00", "Z1 (km)": 2.8, "Z2 (km)": 1.1, "Z3 (km)": 0.18, "Acc34": 11, "Acc4": 2, "Dec34": 9, "Dec4": 2},
]


def get_template_path() -> Optional[str]:
    """Recherche le fichier modèle Modele_Vierge_STAPS.xlsx de manière absolue et robuste."""
    possible_paths = [
        os.path.join(PROJECT_ROOT, "data", TEMPLATE_FILENAME),
        os.path.join(PROJECT_ROOT, TEMPLATE_FILENAME),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", TEMPLATE_FILENAME),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), TEMPLATE_FILENAME),
        os.path.join(os.getcwd(), "data", TEMPLATE_FILENAME),
        os.path.join(os.getcwd(), TEMPLATE_FILENAME),
        os.path.join("data", TEMPLATE_FILENAME),
        TEMPLATE_FILENAME,
    ]
    for p in possible_paths:
        if p and os.path.exists(p):
            return os.path.abspath(p)
    return None


# ==============================================================================
# 1. GESTION DES PHOTOS JOUEURS & ENCODAGE BASE64 (RCA AMIENS)
# ==============================================================================

def obtenir_chemin_photo(nom_joueur: str) -> Optional[str]:
    """
    Retourne le chemin absolu de la photo du joueur dans assets/players_photos/
    en s'appuyant sur l'algorithme robuste de photo_manager (normalisation,
    alias RCA, nom inversé et fuzzy matching).
    """
    if not nom_joueur or nom_joueur.strip() in ["", "-", "— Sélectionner un joueur —"]:
        return None
    return photo_manager.get_player_photo_path(nom_joueur)


def get_player_image_b64(player_name: str, size: int = 80) -> str:
    """
    Convertit la photo du joueur en Data URI Base64 (data:image/png;base64,...)
    pour injection directe dans le HTML/CSS des podiums et composants Streamlit.
    Si le joueur n'a pas de photo dans assets/players_photos/, génère un avatar SVG
    propre aux couleurs officielles du RCA Amiens (Bleu #003399 et Jaune/Or #FFD700)
    avec les initiales du joueur.
    """
    return photo_manager.get_player_image_b64(player_name, size=size)


def photo_to_base64_data_uri(nom_joueur: str) -> str:
    """Alias rétro-compatible pour get_player_image_b64."""
    return get_player_image_b64(nom_joueur)


# ==============================================================================
# 2. DATA VISUALISATION PLOTLY & PODIUM (STYLE CANVA ÉPURÉ)
# ==============================================================================

def creer_graphique_jauge(titre: str, val_objectif: float, val_realise: float) -> go.Figure:
    """
    Génère un graphique à barres stylisé inspiré du design Canva RCA :
    - Deux barres côte à côte : 'Objectif' (Bleu roi) et 'Réalisé' (Rouge RCA).
    - Valeurs et pourcentages d'accomplissement intégrés.
    """
    val_obj = float(val_objectif)
    val_real = float(val_realise)
    pct = (val_real / val_obj * 100) if val_obj > 0 else 0.0

    fmt_obj = f"{int(round(val_obj)):,d}".replace(",", " ") if abs(val_obj - round(val_obj)) < 0.01 else f"{val_obj:.1f}"
    fmt_real = f"{int(round(val_real)):,d}".replace(",", " ") if abs(val_real - round(val_real)) < 0.01 else f"{val_real:.1f}"

    max_val = max(val_obj, val_real, 1.0)
    y_max = max_val * 1.30

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=["Objectif", "Réalisé"],
        y=[val_obj, val_real],
        marker=dict(
            color=[COLOR_OBJECTIF, COLOR_REALISE],
            line=dict(width=0),
            cornerradius=8,
        ),
        text=[f"<b>{fmt_obj}</b>", f"<b>{fmt_real}</b>"],
        textposition="outside",
        textfont=dict(family="Segoe UI, Montserrat, Inter, sans-serif", size=13, color="#1A2B4C"),
        width=[0.42, 0.42],
        hovertemplate="<b>%{x}</b> : %{text}<extra></extra>",
        showlegend=False,
    ))

    if val_real > 0.14 * max_val:
        y_pct = val_real * 0.50
        font_color = "#FFFFFF"
        font_size = 19
    else:
        y_pct = val_real + (max_val * 0.08)
        font_color = COLOR_REALISE
        font_size = 14

    fig.add_annotation(
        x="Réalisé",
        y=y_pct,
        text=f"<b>{pct:.0f}%</b>",
        showarrow=False,
        font=dict(family="Segoe UI, Montserrat, Inter, sans-serif", size=font_size, color=font_color),
        align="center",
    )

    fig.update_layout(
        title=dict(
            text=f"<b>{titre}</b>",
            font=dict(family="Segoe UI, Montserrat, Inter, sans-serif", size=14, color=COLOR_TEXT_MAIN),
            x=0.5,
            xanchor="center",
            y=0.96,
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=10, t=46, b=26),
        height=275,
        xaxis=dict(
            showgrid=False,
            showline=False,
            zeroline=False,
            tickfont=dict(family="Segoe UI, Montserrat, Inter, sans-serif", size=12, color="#4A5568", weight="bold"),
        ),
        yaxis=dict(
            showgrid=False,
            showline=False,
            showticklabels=False,
            zeroline=False,
            range=[0, y_max],
        ),
    )
    return fig


def render_dashboard_objectifs(group_name: str, metrics: Dict[str, Tuple[float, float]]) -> None:
    """Affiche 4 graphiques stylisés sur une seule ligne (st.columns(4))."""
    col1, col2, col3, col4 = st.columns(4)
    metric_keys = list(metrics.keys())
    cols = [col1, col2, col3, col4]
    for idx, key in enumerate(metric_keys[:4]):
        val_obj, val_real = metrics[key]
        fig = creer_graphique_jauge(key, val_obj, val_real)
        with cols[idx]:
            st.plotly_chart(
                fig,
                use_container_width=True,
                config={"displayModeBar": False},
            )


def render_podium_section(
    podium_data: Dict[str, Dict[str, Any]],
    subtitle_map: Optional[Dict[str, str]] = None,
) -> None:
    """
    Affiche la section Podium avec les Tops de la séance :
    - 3 catégories : Top speed, % marche, Haute intensité.
    - Design 3 marches conforme aux exigences : 2e à gauche (Argent), 1er au centre surélevé (Or), 3e à droite (Bronze).
    - Médaillons circulaires avec photos réelles Base64 (80px 1er, 65px 2e/3e) ou avatar SVG officiel RCA.
    - Nom du joueur, Poste de rugby et valeur métrique en gras.
    - En-tête enrichi avec la moyenne exacte du groupe (subtitle_map).
    """
    st.divider()
    st.markdown('<div class="section-header">🏆 PODIUMS DE LA SÉANCE</div>', unsafe_allow_html=True)

    if not podium_data or not isinstance(podium_data, dict):
        st.info("Aucune donnée disponible pour le podium.")
        return

    col1, col2, col3 = st.columns(3)
    cols = [col1, col2, col3]
    categories = list(podium_data.keys())

    for idx, cat_name in enumerate(categories):
        cat_info = podium_data.get(cat_name) or {}
        if not isinstance(cat_info, dict):
            continue
        players = cat_info.get("players", [])
        if not isinstance(players, list):
            players = []
        icon = str(cat_info.get("icon", "🏉"))
        
        # Sécurisation du sous-titre
        sub_title = ""
        if isinstance(subtitle_map, dict):
            raw_sub = subtitle_map.get(cat_name)
            if raw_sub is not None:
                sub_title = str(raw_sub).strip()

        # Appel avec try/except fallback au cas où une ancienne signature serait en cache Streamlit Cloud
        try:
            card_html = photo_manager.render_podium_html_markup(
                category_title=str(cat_name),
                category_icon=icon,
                players=players,
                subtitle=sub_title,
            )
        except TypeError:
            try:
                card_html = photo_manager.render_podium_html_markup(
                    category_title=str(cat_name),
                    category_icon=icon,
                    players=players,
                )
            except Exception as e:
                st.error(f"Erreur d'affichage du podium {cat_name}: {e}")
                continue
        except Exception as e:
            st.error(f"Erreur d'affichage du podium {cat_name}: {e}")
            continue

        card_html_str = str(card_html) if card_html else ""
        clean_html = "\n".join([line.lstrip() for line in card_html_str.splitlines() if line.strip()])
        with cols[idx % 3]:
            if hasattr(st, "html"):
                st.html(clean_html)
            else:
                st.markdown(clean_html, unsafe_allow_html=True)


# ==============================================================================
# 3. PARSING ROBUSTE & CONVERSION NUMÉRIQUE / TEMPORELLE
# ==============================================================================

def parse_duration_to_minutes(val: Any) -> float:
    """Convertit une valeur de durée individuelle en minutes réelles (float)."""
    if pd.isna(val) or val is None or val == "":
        return 0.0

    if isinstance(val, (int, float)):
        num = float(val)
        if num > 100000:
            return num / 60000.0
        elif num > 240:
            return num / 60.0
        return num

    val_str = str(val).strip().replace(",", ".")

    if ":" in val_str:
        try:
            td = pd.to_timedelta(val_str)
            return float(td.total_seconds() / 60.0)
        except Exception:
            parts = val_str.split(":")
            try:
                if len(parts) == 3:
                    return float(parts[0]) * 60.0 + float(parts[1]) + float(parts[2]) / 60.0
                elif len(parts) == 2:
                    return float(parts[0]) + float(parts[1]) / 60.0
            except (ValueError, TypeError):
                return 0.0

    try:
        num = float(val_str)
        if num > 100000:
            return num / 60000.0
        elif num > 240:
            return num / 60.0
        return num
    except (ValueError, TypeError):
        return 0.0


def parse_duration_series(series: pd.Series, col_name: str = "") -> pd.Series:
    """Convertit une série complète de durées Catapult en minutes réelles (float)."""
    if series.empty:
        return pd.Series(dtype=float)

    col_lower = str(col_name).lower()
    raw_str = series.astype(str).str.strip().str.replace(",", ".")

    has_colon = raw_str.str.contains(":", na=False).any()
    if has_colon:
        return series.apply(parse_duration_to_minutes).astype(float)

    num_series = pd.to_numeric(raw_str, errors="coerce").fillna(0.0).astype(float)

    if any(u in col_lower for u in ["(ms)", " ms", "millisec"]):
        return (num_series / 60000.0).astype(float)
    if any(u in col_lower for u in ["(s)", "(sec)", "(secs)", " seconds", " sec"]):
        return (num_series / 60.0).astype(float)
    if any(u in col_lower for u in ["(min)", "(mins)", " minutes"]):
        return num_series.astype(float)

    max_val = num_series.max() if len(num_series) > 0 else 0.0
    positive_vals = num_series[num_series > 0]
    median_val = positive_vals.median() if len(positive_vals) > 0 else 0.0

    if max_val > 100000:
        return (num_series / 60000.0).astype(float)
    elif max_val > 240 or median_val > 180:
        return (num_series / 60.0).astype(float)
    else:
        return num_series.astype(float)


def clean_numeric_value(val: Any) -> float:
    """Convertit une valeur quelconque avec virgule ou espace en float."""
    if pd.isna(val) or val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace(" ", "").replace("\xa0", "").replace(",", ".")
    try:
        return float(val_str)
    except (ValueError, TypeError):
        return 0.0


def find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Recherche flexible d'une colonne insensible aux espaces et à la casse."""
    col_map = {re.sub(r"\s+", " ", str(c).strip().lower()): c for c in df.columns}

    for candidate in candidates:
        norm_candidate = re.sub(r"\s+", " ", candidate.strip().lower())
        if norm_candidate in col_map:
            return col_map[norm_candidate]

    for candidate in candidates:
        norm_candidate = re.sub(r"\s+", " ", candidate.strip().lower())
        for norm_col, orig_col in col_map.items():
            if norm_candidate in norm_col:
                return orig_col

    return None


def find_catapult_header_row(file_bytes: bytes, encoding: str = "utf-8", max_scan: int = 15) -> int:
    """Scanne les 15 premières lignes pour repérer dynamiquement l'en-tête réel."""
    try:
        text_preview = file_bytes[:16384].decode(encoding, errors="replace")
        lines = text_preview.splitlines()[:max_scan]
    except Exception:
        return 0

    for idx, raw_line in enumerate(lines):
        line = raw_line.strip().lower()
        if not line:
            continue
        if any(term in line for term in ["player name", "athlete name", "nom ; prénom", "nom;prénom", "player", "athlete"]):
            return idx

    for idx, raw_line in enumerate(lines):
        line = raw_line.strip().lower()
        if not line:
            continue
        if "date" in line and any(term in line for term in ["session", "split", "distance", "duration", "speed"]):
            return idx

    return 0


def load_catapult_csv(file_bytes: bytes) -> pd.DataFrame:
    """Lit un fichier CSV Catapult One avec détection automatique d'encodage et délimiteur."""
    encodings_to_try: List[str] = ["utf-8-sig", "utf-8", "latin1", "cp1252", "iso-8859-1"]
    separators_to_try: List[str] = [";", ",", "\t"]
    last_error: Optional[Exception] = None

    for enc in encodings_to_try:
        header_idx = find_catapult_header_row(file_bytes, encoding=enc, max_scan=15)
        for sep in separators_to_try:
            try:
                df = pd.read_csv(
                    io.BytesIO(file_bytes),
                    sep=sep,
                    encoding=enc,
                    header=header_idx,
                    low_memory=False,
                )
                if len(df.columns) >= 3:
                    df.columns = [str(c).strip() for c in df.columns]
                    lower_cols = [c.lower() for c in df.columns]
                    if any("player" in c or "athlete" in c or "nom" in c or "date" in c for c in lower_cols):
                        return df.dropna(how="all")
            except Exception as exc:
                last_error = exc
                continue

    if last_error is not None:
        raise last_error
    raise ValueError("Impossible de décoder le fichier CSV avec les encodages et délimiteurs standards.")


def load_catapult_file(uploaded_file: Any) -> pd.DataFrame:
    """Lit un export Catapult One (.csv, .xlsx, .xls) en sautant les métadonnées."""
    file_name = getattr(uploaded_file, "name", "file.xlsx").lower()
    file_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()

    if file_name.endswith((".xlsx", ".xls")):
        preview_df = pd.read_excel(io.BytesIO(file_bytes), header=None, nrows=15)
        header_idx = 0
        for idx, row in preview_df.iterrows():
            row_str = " ".join([str(v).lower() for v in row if pd.notna(v)])
            if any(term in row_str for term in ["player name", "athlete name", "nom ; prénom", "nom;prénom", "athlete", "player"]):
                header_idx = idx
                break
            elif "date" in row_str and any(term in row_str for term in ["session", "split", "distance", "duration"]):
                header_idx = idx
                break

        df = pd.read_excel(io.BytesIO(file_bytes), header=header_idx)
        df.columns = [str(c).strip() for c in df.columns]
        return df.dropna(how="all")
    else:
        return load_catapult_csv(file_bytes)


# ==============================================================================
# 4. MAPPING GPS DYNAMIQUE & INTELLIGENT (EFFECTIF DU CLUB)
# ==============================================================================

# Mapping standard de référence des 21 boîtiers GPS du club
DEFAULT_GPS_MAPPING: Dict[int, str] = {
    1: "PECQUERY Thomas",
    2: "Jean-Étienne ( TIJ)",
    3: "COZIC Adrien",
    4: "LAFEUILLE Emilien",
    5: "SOTILLO Alejo",
    6: "ORION Lucas",
    7: "VASELLI Enzo",
    8: "DAULT Bastien",
    9: "BAM'S",
    10: "CARPENTIER Maël",
    11: "SALUA Makatea",
    12: "HARRY Maxence",
    13: "ULMUCK Victor",
    14: "PREVOTE Olivier",
    15: "LESPINASSE Lucien",
    16: "PREVOST Thomas",
    17: "BLANC Jules",
    19: "DE BRUYN Octave",
    20: "TREMBLAY Brieuc",
    21: "BŒUF Eliot",
    22: "DELACROIX Jules",
}

# Fichier JSON local pour la persistance permanente des attributions GPS
CONFIG_GPS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_rca_gps.json")


def charger_mapping_gps_persistant() -> Dict[str, str]:
    """
    Charge le dictionnaire d'association GPS -> Joueur depuis config_rca_gps.json s'il existe.
    Retourne un dictionnaire normalisé { 'gps_id': 'Nom Joueur' }.
    """
    if os.path.exists(CONFIG_GPS_FILE):
        try:
            with open(CONFIG_GPS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    raw_map = data.get("mapping", data)
                    if isinstance(raw_map, dict):
                        return {
                            str(k).strip(): str(v).strip()
                            for k, v in raw_map.items()
                            if v and str(v).strip() not in ["— Sélectionner un joueur —", "-"]
                        }
        except Exception:
            pass
    return {}


def sauvegarder_mapping_gps_persistant(
    mapping: Dict[Any, str],
    last_file_sig: Optional[str] = None,
    overwrite: bool = False,
    splits_cache: Optional[Dict[Any, bool]] = None,
) -> bool:
    """
    Sauvegarde le dictionnaire d'association GPS -> Joueur dans config_rca_gps.json.
    Fusionne intelligemment avec les attributions déjà enregistrées pour ne jamais
    écraser les boîtiers non présents dans la séance active (sauf si overwrite=True).
    Mémorise également la signature du fichier validé et le cache de participation aux splits
    pour restaurer intégralement la session au rafraîchissement.
    """
    full_map: Dict[str, str] = {} if overwrite else charger_mapping_gps_persistant()
    for k, v in mapping.items():
        if v and str(v).strip() and str(v).strip() not in ["— Sélectionner un joueur —", "-"]:
            full_map[str(k).strip()] = str(v).strip()

    existing_sig = None
    existing_splits = {}
    if os.path.exists(CONFIG_GPS_FILE):
        try:
            with open(CONFIG_GPS_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
                if isinstance(d, dict):
                    existing_sig = d.get("last_validated_file_sig")
                    existing_splits = d.get("splits_participation", {})
        except Exception:
            pass

    # Conversion propre du cache de participation aux splits
    splits_dict_to_save = dict(existing_splits) if isinstance(existing_splits, dict) else {}
    if splits_cache:
        for k_sp, v_sp in splits_cache.items():
            if isinstance(k_sp, (tuple, list)) and len(k_sp) == 2:
                key_str = f"{k_sp[0]}:::{k_sp[1]}"
            else:
                key_str = str(k_sp)
            splits_dict_to_save[key_str] = bool(v_sp)

    payload = {
        "mapping": full_map,
        "derniere_mise_a_jour": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "nb_boitiers": len(full_map),
        "last_validated_file_sig": last_file_sig or existing_sig,
        "splits_participation": splits_dict_to_save,
    }
    try:
        with open(CONFIG_GPS_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        try:
            try:

                import github_sync

            except (ImportError, ModuleNotFoundError):

                from modules import github_sync
            github_sync.push_file_to_github(
                CONFIG_GPS_FILE,
                commit_message="config: mise à jour mapping boîtiers GPS"
            )
        except Exception:
            pass

        return True
    except Exception:
        return False


# ==============================================================================
# 4.1 MODAL DE CRÉATION & GESTION D'UN JOUEUR DANS L'EFFECTIF
# ==============================================================================

POSTES_RUGBY_STANDARDS: List[str] = [
    "Pilier",
    "Talonneur",
    "2ème ligne",
    "3ème ligne",
    "Demi de mêlée",
    "Demi d'ouverture",
    "Centre",
    "Ailier",
    "Arrière",
    "Autre (saisie libre)",
]


@st.dialog("👤 Créer / Ajouter un joueur dans l'effectif RCA", width="large")
def dialog_creer_nouveau_joueur(
    boitier_cible: Optional[Union[int, str]] = None,
    liste_boitiers: Optional[List[int]] = None,
) -> None:
    st.markdown(
        "<div style='font-size: 0.92rem; color: #475569; margin-bottom: 14px; line-height: 1.4;'>"
        "Ajoutez un joueur non répertorié à l'effectif officiel du <strong>Rugby Club Amiénois</strong>. "
        "Il sera immédiatement disponible dans toutes les listes de sélection GPS, graphiques et rapports officiels."
        "</div>",
        unsafe_allow_html=True,
    )

    nb_custom = len(JOUEURS_CUSTOM_CACHE)
    existing_photos = photo_manager.list_existing_player_photos()
    nb_photos = len(existing_photos)

    tab_creer, tab_liste, tab_trombi = st.tabs([
        "➕ Créer un nouveau joueur",
        f"📋 Joueurs personnalisés ({nb_custom})",
        f"📸 Trombinoscope & Photos ({nb_photos})",
    ])

    with tab_creer:
        col_nom, col_poste = st.columns([1.2, 1])
        with col_nom:
            nouveau_nom = st.text_input(
                "Nom et Prénom du joueur * :",
                placeholder="Ex: DUPONT Antoine",
                key="dlg_input_nom_joueur",
                help="Saisissez le Nom de famille suivi du Prénom (ex: DUPONT Antoine).",
            )
        with col_poste:
            choix_poste = st.selectbox(
                "Poste de rugby * :",
                options=POSTES_RUGBY_STANDARDS,
                index=0,
                key="dlg_input_poste_joueur",
            )

        poste_final = choix_poste
        if choix_poste == "Autre (saisie libre)":
            poste_autre = st.text_input(
                "Préciser le poste de rugby :",
                placeholder="Ex: Polyvalent 3/4",
                key="dlg_input_poste_autre",
            )
            poste_final = poste_autre.strip() if poste_autre.strip() else "Trois-quarts"

        # Groupe officiel RCA suggéré
        groupe_suggere = mapper_poste_vers_groupe_rca(poste_final, nouveau_nom)
        idx_grp = (
            ["5 de devant", "3ème Ligne", "Trois-quarts"].index(groupe_suggere)
            if groupe_suggere in ["5 de devant", "3ème Ligne", "Trois-quarts"]
            else 2
        )

        col_grp, col_boitier = st.columns(2)
        with col_grp:
            groupe_final = st.selectbox(
                "Groupe athlétique RCA officiel * :",
                options=["5 de devant", "3ème Ligne", "Trois-quarts"],
                index=idx_grp,
                help="Groupe d'appartenance utilisé pour les analyses collectives et objectifs de la préparatrice physique.",
                key="dlg_input_groupe_joueur",
            )

        with col_boitier:
            # Options d'attribution immédiate à un boîtier GPS
            opts_boitiers = ["— Ne pas attribuer immédiatement —"]
            boitier_default_idx = 0
            if liste_boitiers:
                for idx_b, b_num in enumerate(liste_boitiers, start=1):
                    opts_boitiers.append(f"Boîtier GPS n°{b_num}")
                    if boitier_cible is not None and str(b_num) == str(boitier_cible):
                        boitier_default_idx = idx_b

            boitier_choisi_str = st.selectbox(
                "Assigner directement à un boîtier GPS :",
                options=opts_boitiers,
                index=boitier_default_idx,
                help="Si sélectionné, le nouveau joueur sera directement attribué à ce boîtier GPS.",
                key="dlg_input_boitier_joueur",
            )

        col_photo, col_preview = st.columns([2, 1])
        with col_photo:
            photo_file = st.file_uploader(
                "Photo du joueur (optionnel) :",
                type=["jpg", "jpeg", "png", "webp"],
                help="Photo de profil pour le mapping GPS et les podiums de la séance.",
                key="dlg_upload_photo_joueur",
            )
        with col_preview:
            st.caption("Aperçu du groupe :")
            if groupe_final == "5 de devant":
                b_style = "background-color:#EFF6FF;color:#1D4ED8;border:1px solid #BFDBFE;"
            elif groupe_final == "3ème Ligne":
                b_style = "background-color:#F0FDF4;color:#15803D;border:1px solid #BBF7D0;"
            else:
                b_style = "background-color:#FEF2F2;color:#B91C1C;border:1px solid #FECACA;"
            st.markdown(
                f"<div style='{b_style}padding:6px 10px;border-radius:6px;font-weight:700;font-size:0.80rem;text-align:center;margin-top:2px;'>"
                f"🏉 {groupe_final}</div>",
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
        col_sub1, col_sub2 = st.columns([2, 1])
        with col_sub1:
            if st.button("💾 Enregistrer et créer le joueur", type="primary", use_container_width=True, key="dlg_btn_enregistrer"):
                nom_clean = nouveau_nom.strip()
                if not nom_clean:
                    st.error("⚠️ Veuillez renseigner le Nom et Prénom du joueur.")
                elif nom_clean in ["— Sélectionner un joueur —", "➕ [Créer un nouveau joueur...]", "-"]:
                    st.error("⚠️ Le nom saisi n'est pas autorisé.")
                else:
                    p_bytes = None
                    p_ext = ".jpg"
                    if photo_file is not None:
                        p_bytes = photo_file.getvalue()
                        p_ext = os.path.splitext(photo_file.name)[1]

                    ok, msg = sauvegarder_joueur_personnalise(
                        nom=nom_clean,
                        poste=poste_final,
                        groupe=groupe_final,
                        photo_bytes=p_bytes,
                        photo_ext=p_ext,
                    )
                    if ok:
                        # Si assignation directe à un boîtier GPS
                        if boitier_choisi_str and "Boîtier GPS n°" in boitier_choisi_str:
                            b_num = int(boitier_choisi_str.replace("Boîtier GPS n°", "").strip())
                            st.session_state[f"gps_sel_boitier_{b_num}"] = nom_clean
                            if "gps_mapping" in st.session_state:
                                st.session_state["gps_mapping"][b_num] = nom_clean
                                st.session_state["gps_mapping"][str(b_num)] = nom_clean
                            if "mapping_gps_joueurs" in st.session_state:
                                st.session_state["mapping_gps_joueurs"][b_num] = nom_clean
                                st.session_state["mapping_gps_joueurs"][str(b_num)] = nom_clean

                            c_map = charger_mapping_gps_persistant()
                            c_map[str(b_num)] = nom_clean
                            sauvegarder_mapping_gps_persistant(c_map)

                        st.toast(f"✅ {msg}")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")
        with col_sub2:
            if st.button("Fermer", use_container_width=True, key="dlg_btn_fermer"):
                st.rerun()

    with tab_liste:
        custom_players = charger_joueurs_personnalises()
        if not custom_players:
            st.info("Aucun joueur personnalisé n'a encore été créé dans le logiciel.")
        else:
            st.markdown(f"**{len(custom_players)} joueur(s) personnalisé(s) enregistré(s) :**")
            for c_nom, c_val in list(custom_players.items()):
                cj1, cj2, cj3, cj4 = st.columns([3, 2, 2, 0.8])
                with cj1:
                    photo_c = obtenir_chemin_photo(c_nom)
                    if photo_c and os.path.isfile(photo_c):
                        st.markdown(f"👤 **{c_nom}** *(photo 📸)*")
                    else:
                        st.markdown(f"👤 **{c_nom}**")
                with cj2:
                    st.caption(f"Poste : *{c_val.get('poste', 'Inconnu')}*")
                with cj3:
                    st.caption(f"Groupe : `{c_val.get('groupe', 'Non défini')}`")
                with cj4:
                    if st.button("🗑️", key=f"dlg_del_p_{c_nom}", help=f"Supprimer {c_nom}"):
                        ok_d, msg_d = supprimer_joueur_personnalise(c_nom)
                        if ok_d:
                            st.toast(f"✅ {msg_d}")
                            st.rerun()
                        else:
                            st.error(msg_d)

    with tab_trombi:
        st.markdown(
            "<div style='font-size: 0.90rem; color: #334155; margin-bottom: 12px;'>"
            "Gestion officielle du trombinoscope et des portraits de l'effectif sénior du <strong>RCA Amiens</strong>. "
            "Les photos sont enregistrées dans <code>assets/players_photos/</code> sous la forme normalisée <code>prenom_nom.png</code>."
            "</div>",
            unsafe_allow_html=True,
        )

        col_sync1, col_sync2 = st.columns([2.5, 1.5])
        with col_sync1:
            st.caption(f"📁 Dossier local : `assets/players_photos/` ({nb_photos} photo(s) chargée(s))")
        with col_sync2:
            if st.button("🔄 Ré-extraire depuis PDF", help="Extrait à nouveau les photos depuis EFFECTIF SÉNIORS (1).pdf", key="dlg_btn_reextract_trombi", use_container_width=True):
                with st.spinner("Extraction et appariement géométrique des photos depuis le trombinoscope PDF..."):
                    res_ext = photo_manager.extract_all_photos_from_pdf()
                    if res_ext.get("success"):
                        st.toast(f"✅ {res_ext.get('saved_count', 0)} photos extraites et synchronisées !")
                        st.rerun()
                    else:
                        st.error(f"❌ Échec de l'extraction : {res_ext.get('error', 'Erreur')}")

        st.divider()

        # Liste complète de tous les joueurs (PDF + Personnalisés)
        tous_les_joueurs = sorted(list(set(list(joueurs_data.keys()) + list(JOUEURS_CUSTOM_CACHE.keys()))))
        joueur_choisi_trombi = st.selectbox(
            "Sélectionner un joueur de l'effectif pour inspecter ou modifier sa photo :",
            options=tous_les_joueurs,
            key="dlg_select_joueur_trombi",
        )

        if joueur_choisi_trombi:
            col_t_photo, col_t_actions = st.columns([1.2, 2.8])
            photo_existante = photo_manager.get_player_photo_path(joueur_choisi_trombi)
            has_real_photo = bool(photo_existante and os.path.isfile(photo_existante))

            with col_t_photo:
                img_b64 = get_player_image_b64(joueur_choisi_trombi, size=90)
                statut_badge = (
                    "<span style='background:#DCFCE7;color:#166534;font-size:0.75rem;padding:2px 8px;border-radius:12px;font-weight:700;'>Photo officielle ✔</span>"
                    if has_real_photo
                    else "<span style='background:#FEF3C7;color:#92400E;font-size:0.75rem;padding:2px 8px;border-radius:12px;font-weight:700;'>Avatar par défaut</span>"
                )
                st.markdown(
                    f"""
                    <div style="text-align: center; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px;">
                        <img src="{img_b64}" style="width: 85px; height: 85px; border-radius: 50%; object-fit: cover; border: 2.5px solid #003399; margin-bottom: 8px;" />
                        <div>{statut_badge}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            with col_t_actions:
                poste_actuel = joueurs_data.get(joueur_choisi_trombi) or JOUEURS_CUSTOM_CACHE.get(joueur_choisi_trombi, {}).get("poste", "Inconnu")
                groupe_actuel = mapper_poste_vers_groupe_rca(poste_actuel, joueur_choisi_trombi)
                st.markdown(f"**Joueur :** {joueur_choisi_trombi} | **Poste :** `{poste_actuel}` | **Groupe :** `{groupe_actuel}`")

                new_photo_file = st.file_uploader(
                    "Uploader / Remplacer la photo de ce joueur :",
                    type=["png", "jpg", "jpeg", "webp"],
                    key=f"upload_photo_perso_{joueur_choisi_trombi}",
                    help="La photo sera enregistrée au format prenom_nom.png et apparaîtra instantanément sur les podiums et le mapping GPS.",
                )

                col_btn_sav, col_btn_del = st.columns(2)
                with col_btn_sav:
                    if new_photo_file is not None:
                        if st.button("💾 Enregistrer cette photo", type="primary", use_container_width=True, key=f"btn_save_photo_{joueur_choisi_trombi}"):
                            p_ext = os.path.splitext(new_photo_file.name)[1]
                            saved = photo_manager.save_manual_player_photo(joueur_choisi_trombi, new_photo_file.getvalue(), ext=p_ext)
                            if saved:
                                st.toast(f"✅ Photo enregistrée pour {joueur_choisi_trombi} !")
                                st.rerun()
                            else:
                                st.error("❌ Erreur lors de la sauvegarde de la photo.")
                with col_btn_del:
                    if has_real_photo:
                        if st.button("🗑️ Supprimer la photo", use_container_width=True, key=f"btn_del_photo_{joueur_choisi_trombi}"):
                            deleted = photo_manager.delete_player_photo(joueur_choisi_trombi)
                            if deleted:
                                st.toast(f"🗑️ Photo supprimée pour {joueur_choisi_trombi}.")
                                st.rerun()


# ==============================================================================
# 4.1b GESTION DES SPLITS D'ENTRAÎNEMENT & MATRICE DE PARTICIPATION
# ==============================================================================

GLOBAL_SPLIT_TAGS = {
    "session", "whole session", "full session", "all", "total",
    "seance", "séance", "globale", "global", "complete", "complète",
    "session summary", "all session", "overall", "nan"
}


def _natural_split_sort_key(s: str) -> List[Union[int, str]]:
    """Clé de tri naturel pour ordonner Split 1, Split 2 ... Split 10 proprement."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r"(\d+)", str(s).strip())]


def extraire_splits_entrainement(df_raw: Optional[pd.DataFrame]) -> Tuple[Optional[str], List[str]]:
    """
    Extrait et ordonne naturellement tous les splits d'entraînement uniques présents dans le fichier,
    en excluant la ligne globale 'all', 'session' et autres récapitulatifs globaux.
    Retourne (split_col, training_splits).
    """
    if df_raw is None or df_raw.empty:
        return None, []

    split_col = find_column(df_raw, [
        "Split Name", "Split", "Period Name", "Period", "Tags", "Tag", "Activity Name", "Session Name", "Session Title", "Segment"
    ])
    if not split_col or split_col not in df_raw.columns:
        return None, []

    raw_splits = df_raw[split_col].dropna().astype(str).str.strip().unique().tolist()
    training_splits = [
        s for s in raw_splits
        if s and s.lower() not in GLOBAL_SPLIT_TAGS and not s.lower().startswith("session ")
    ]

    sorted_splits = sorted(training_splits, key=_natural_split_sort_key)
    return split_col, sorted_splits


def detecter_anomalies_splits(
    df_raw: pd.DataFrame,
    mapping_gps: Dict[Union[int, str], str],
    split_col: str,
    training_splits: List[str],
) -> List[Dict[str, Any]]:
    """
    Détecte les anomalies de participation aux splits :
    Identifie si un joueur assigné a Distance == 0 (ou Durée == 0 ou split manquant) sur l'un des splits d'entraînement.
    Retourne une liste de dictionnaires décrivant chaque anomalie.
    """
    anomalies: List[Dict[str, Any]] = []
    if df_raw is None or df_raw.empty or not split_col or not training_splits:
        return anomalies

    player_col = find_column(df_raw, [
        "Nom ; Prénom", "Player Name", "Athlete Name", "Player", "Athlete", "Joueur", "Nom", "Name"
    ])
    if not player_col:
        return anomalies

    dist_col = find_column(df_raw, ["Distance (km)", "Total Distance (km)", "Distance (m)", "Total Distance (m)", "Distance"])
    dur_col = find_column(df_raw, ["Duration", "Temps de jeu", "Duration (mins)", "Duration (s)", "Durée", "Time"])

    df_mapped = appliquer_mapping(df_raw, mapping_gps)

    placeholder = "— Sélectionner un joueur —"
    trigger_new = "➕ [Créer un nouveau joueur...]"

    player_to_gps: Dict[str, Union[int, str]] = {}
    for g_id, p_name in mapping_gps.items():
        if p_name and p_name not in [placeholder, trigger_new, "-"]:
            player_to_gps[p_name] = g_id

    for p_name, g_id in player_to_gps.items():
        sub_p = df_mapped[df_mapped[player_col].astype(str).str.strip().str.lower() == p_name.strip().lower()]
        for sp in training_splits:
            sub_sp = sub_p[sub_p[split_col].astype(str).str.strip().str.lower() == sp.strip().lower()]
            if sub_sp.empty:
                anomalies.append({
                    "player": p_name,
                    "gps_id": g_id,
                    "split": sp,
                    "dist_m": 0.0,
                    "message": f"**{p_name}** (GPS {g_id}) n'a aucune donnée enregistrée sur le **{sp}**",
                })
            else:
                d_val = clean_numeric_value(sub_sp[dist_col].iloc[0]) if dist_col and dist_col in sub_sp.columns else 0.0
                if dist_col and "(km)" in dist_col.lower() and d_val < 100.0:
                    d_m = d_val * 1000.0
                else:
                    d_m = d_val

                dur_min = parse_duration_to_minutes(sub_sp[dur_col].iloc[0]) if dur_col and dur_col in sub_sp.columns else 1.0

                if d_m <= 0.0 or dur_min <= 0.0:
                    anomalies.append({
                        "player": p_name,
                        "gps_id": g_id,
                        "split": sp,
                        "dist_m": d_m,
                        "message": f"**{p_name}** (GPS {g_id}) a **0 m** enregistré sur le **{sp}**",
                    })

    return anomalies


def construire_matrice_splits(
    df_raw: pd.DataFrame,
    mapping_gps: Dict[Union[int, str], str],
    split_col: str,
    training_splits: List[str],
) -> pd.DataFrame:
    """
    Construit le DataFrame pour la matrice interactive Joueurs × Splits (st.data_editor) :
    - Colonnes fixes : GPS, Joueur, Groupe / Poste, Distance brute (m)
    - Colonnes dynamiques : une colonne booléenne par split d'entraînement (True par défaut)
    """
    rows = []
    if df_raw is None or df_raw.empty or not training_splits:
        return pd.DataFrame()

    player_col = find_column(df_raw, [
        "Nom ; Prénom", "Player Name", "Athlete Name", "Player", "Athlete", "Joueur", "Nom", "Name"
    ])
    dist_col = find_column(df_raw, ["Distance (km)", "Total Distance (km)", "Distance (m)", "Total Distance (m)", "Distance"])

    df_mapped = appliquer_mapping(df_raw, mapping_gps)

    placeholder = "— Sélectionner un joueur —"
    trigger_new = "➕ [Créer un nouveau joueur...]"

    assigned_items = []
    for g_id, p_name in mapping_gps.items():
        if p_name and p_name not in [placeholder, trigger_new, "-"]:
            try:
                g_num = int(g_id)
            except (ValueError, TypeError):
                g_num = 999
            assigned_items.append((g_num, g_id, p_name))

    seen_players = set()
    unique_assigned = []
    for g_num, g_id, p_name in sorted(assigned_items, key=lambda x: (x[0], str(x[1]))):
        if p_name not in seen_players:
            seen_players.add(p_name)
            unique_assigned.append((g_num, g_id, p_name))

    cache_choix = st.session_state.get("splits_participation_cache", {})

    for g_num, g_id, p_name in unique_assigned:
        poste_j = joueurs_data.get(p_name, "")
        if not poste_j:
            poste_j = st.session_state.get("session_player_postes", {}).get(p_name, "")
        grp_j = mapper_poste_vers_groupe_rca(poste_j, p_name)

        sub_p = df_mapped[df_mapped[player_col].astype(str).str.strip().str.lower() == p_name.strip().lower()] if player_col else pd.DataFrame()
        sub_training = sub_p[sub_p[split_col].astype(str).str.strip().isin(training_splits)] if not sub_p.empty and split_col else pd.DataFrame()

        dist_brute_m = 0.0
        if not sub_training.empty and dist_col and dist_col in sub_training.columns:
            v_dist = sub_training[dist_col].apply(clean_numeric_value)
            if "(km)" in dist_col.lower() and v_dist.median() < 100.0:
                dist_brute_m = float(v_dist.sum() * 1000.0)
            else:
                dist_brute_m = float(v_dist.sum())

        row_dict: Dict[str, Any] = {
            "GPS": g_num if g_num != 999 else str(g_id),
            "Joueur": p_name,
            "Groupe": f"{grp_j} ({poste_j})" if poste_j else grp_j,
            "Distance brute (m)": int(round(dist_brute_m)),
        }

        # Pour chaque split, cocher True par défaut, ou restaurer l'état utilisateur
        for sp in training_splits:
            default_val = cache_choix.get((p_name, sp), True)
            row_dict[sp] = bool(default_val)

        rows.append(row_dict)

    return pd.DataFrame(rows)


def filtrer_splits_selon_matrice(
    df_raw_mapped: pd.DataFrame,
    df_matrix: pd.DataFrame,
    split_col: str,
    training_splits: List[str],
) -> pd.DataFrame:
    """
    Filtre le DataFrame brut mappé pour ne conserver strictement que les lignes (Joueur, Split)
    qui sont cochées True dans la matrice de participation.
    Les lignes de métadonnées hors splits d'entraînement (ex: 'all') sont conservées intactes.
    """
    if df_raw_mapped is None or df_raw_mapped.empty:
        return df_raw_mapped
    if df_matrix is None or df_matrix.empty or not split_col or not training_splits:
        return df_raw_mapped.copy()

    player_col = find_column(df_raw_mapped, [
        "Nom ; Prénom", "Player Name", "Athlete Name", "Player", "Athlete", "Joueur", "Nom", "Name"
    ])
    if not player_col:
        return df_raw_mapped.copy()

    valid_pairs = set()
    for _, r in df_matrix.iterrows():
        p_name = str(r["Joueur"]).strip().lower()
        for sp in training_splits:
            if sp in r and bool(r[sp]):
                valid_pairs.add((p_name, str(sp).strip().lower()))

    training_splits_lower = {str(sp).strip().lower() for sp in training_splits}

    def _keep_row(row: pd.Series) -> bool:
        s_val = str(row[split_col]).strip() if pd.notna(row[split_col]) else ""
        s_lower = s_val.lower()

        if s_lower in training_splits_lower:
            p_val = str(row[player_col]).strip().lower() if pd.notna(row[player_col]) else ""
            return (p_val, s_lower) in valid_pairs

        return True

    mask = df_raw_mapped.apply(_keep_row, axis=1)
    return df_raw_mapped[mask].copy()


# ==============================================================================
# 4.2 INTERFACE D'ATTRIBUTION GPS DYNAMIQUE & SÉLECTION DES JOUEURS
# ==============================================================================

def ui_attribution_gps(df_raw: Optional[pd.DataFrame] = None) -> Tuple[Dict[Union[int, str], str], bool]:
    """
    Interface dynamique d'attribution des boîtiers GPS :
    - Filtre et extrait uniquement les numéros uniques de boîtiers GPS présents dans df_raw.
    - Convertit les identifiants en entiers (int) et les trie par ordre croissant (1, 2, 3...).
    - Mémorisation et pré-remplissage intelligent automatique via st.session_state.mapping_gps_joueurs et config_rca_gps.json.
    - Permet la création instantanée d'un nouveau joueur depuis la barre d'outils ou directement dans chaque sélecteur.
    - Affiche chaque boîtier avec sélecteur de joueur ('Nom Prénom - Poste'), photo et badge du groupe officiel RCA.
    - Convertit automatiquement le poste détaillé vers l'un des 3 groupes officiels RCA (5 de devant, 3ème Ligne, Trois-quarts).
    - Sécurité anti-doublon bloquante : interdit formellement d'attribuer un même joueur à 2 boîtiers.
    - Intègre le bouton de validation explicite « ✅ Valider l'effectif et l'association GPS ».
    Retourne (mapping, has_doublons).
    """
    # 1. Extraction et tri numérique strict des identifiants GPS détectés
    liste_gps_detectes: List[int] = []
    if df_raw is not None and not df_raw.empty:
        col_id = find_column(df_raw, [
            "Player Name", "Athlete Name", "Nom ; Prénom", "Player", "Athlete", "Joueur", "Nom", "Name"
        ])
        if col_id and col_id in df_raw.columns:
            nums_trouves: set[int] = set()
            for val in df_raw[col_id].dropna():
                val_str = str(val).strip()
                # Conversion directe si valeur numérique (int, float, ou str numérique)
                try:
                    num_float = float(val_str)
                    if num_float.is_integer() and num_float > 0:
                        nums_trouves.add(int(num_float))
                        continue
                except (ValueError, TypeError):
                    pass

                # Extraction si chaîne contenant un identifiant numérique (ex: "GPS 17")
                match = re.search(r"\b(\d+)\b", val_str)
                if match:
                    try:
                        n = int(match.group(1))
                        if n > 0:
                            nums_trouves.add(n)
                    except (ValueError, TypeError):
                        pass

            liste_gps_detectes = sorted(list(nums_trouves))

    # Si aucun boîtier numérique n'a été détecté dans le fichier (ex: noms déjà écrits en clair)
    if not liste_gps_detectes:
        names_found = []
        if df_raw is not None and not df_raw.empty:
            col_id = find_column(df_raw, [
                "Player Name", "Athlete Name", "Nom ; Prénom", "Player", "Athlete", "Joueur", "Nom", "Name"
            ])
            if col_id and col_id in df_raw.columns:
                names_found = [str(x).strip() for x in df_raw[col_id].dropna().unique() if str(x).strip()]

        if names_found:
            st.success(f"ℹ️ **{len(names_found)} joueur(s) identifié(s)** directement par leur nom dans le fichier :")
            c_e1, c_e2 = st.columns(2)
            half = (len(names_found) + 1) // 2
            for i_idx, n in enumerate(names_found):
                col_t = c_e1 if i_idx < half else c_e2
                p = joueurs_data.get(n, "Non spécifié")
                g = mapper_poste_vers_groupe_rca(p, n)
                col_t.markdown(f"• **{n}** — Poste : *{p}* ➔ `{g}`")

            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            c_b1, c_b_creer, c_b2, c_b3 = st.columns([0.8, 1.4, 2, 0.8])
            with c_b_creer:
                if st.button("➕ Nouveau joueur", key="btn_creer_joueur_no_units", use_container_width=True, help="Créer un joueur absent de la base"):
                    dialog_creer_nouveau_joueur()
            with c_b2:
                if st.button("✅ Valider l'effectif et l'association GPS", type="primary", use_container_width=True, key="btn_valider_sans_boitiers"):
                    st.session_state["step2_validated"] = True
                    map_direct = {n: n for n in names_found}
                    st.session_state["gps_mapping"] = map_direct
                    st.session_state["mapping_gps_joueurs"] = map_direct
                    sauvegarder_mapping_gps_persistant(map_direct)
                    st.toast("✅ Effectif validé et mémorisé avec succès !")
                    st.rerun()
            return {n: n for n in names_found}, False
        else:
            st.warning("⚠️ Aucun joueur ni numéro de boîtier GPS n'a été détecté dans le fichier.")
            return {}, False

    # 2. Liste des joueurs RCA disponibles (tri alphabétique) avec option d'ajout dynamique
    placeholder = "— Sélectionner un joueur —"
    TRIGGER_NOUVEAU_JOUEUR = "➕ [Créer un nouveau joueur...]"

    # Détection automatique si l'utilisateur a sélectionné l'option de création dans un selectbox
    boitier_declencheur: Optional[int] = None
    for g_id in liste_gps_detectes:
        k_w = f"gps_sel_boitier_{g_id}"
        if st.session_state.get(k_w) == TRIGGER_NOUVEAU_JOUEUR:
            boitier_declencheur = g_id
            # On réinitialise la clé avant que le widget ne soit ré-instancié
            st.session_state[k_w] = placeholder
            break

    if boitier_declencheur is not None:
        dialog_creer_nouveau_joueur(boitier_cible=boitier_declencheur, liste_boitiers=liste_gps_detectes)

    noms_joueurs_tries = sorted(list(joueurs_data.keys()))
    options = [placeholder, TRIGGER_NOUVEAU_JOUEUR] + noms_joueurs_tries

    def format_affichage_joueur(nom: str) -> str:
        """Affiche 'Nom Prénom - Poste' dans la liste déroulante et la barre de recherche."""
        if nom == placeholder:
            return placeholder
        if nom == TRIGGER_NOUVEAU_JOUEUR:
            return "➕ [Créer un nouveau joueur dans l'effectif...]"
        poste = joueurs_data.get(nom, "")
        if poste:
            return f"{nom} - {poste}"
        return nom

    # Initialisation & Synchronisation des dictionnaires de session
    if "mapping_gps_joueurs" not in st.session_state:
        st.session_state["mapping_gps_joueurs"] = charger_mapping_gps_persistant() or {str(k): v for k, v in DEFAULT_GPS_MAPPING.items()}
    if "gps_mapping" not in st.session_state:
        st.session_state["gps_mapping"] = dict(st.session_state["mapping_gps_joueurs"])

    saved_disk = charger_mapping_gps_persistant()
    saved_session = st.session_state.get("mapping_gps_joueurs") or st.session_state.get("gps_mapping") or {}

    # Dictionnaire de référence combiné pour pré-remplir instantanément
    reference_mapping: Dict[str, str] = {str(k): v for k, v in DEFAULT_GPS_MAPPING.items()}
    reference_mapping.update(saved_disk)
    reference_mapping.update({str(k): v for k, v in saved_session.items() if v and v != placeholder})

    # Calcul du nombre de boîtiers pré-remplis
    nb_prefilled = sum(
        1 for g_id in liste_gps_detectes
        if reference_mapping.get(str(g_id)) in options and reference_mapping.get(str(g_id)) not in [placeholder, TRIGGER_NOUVEAU_JOUEUR]
    )

    c_info1, c_btn_creer, c_btn_trombi, c_info2 = st.columns([2.6, 1.3, 1.4, 1.2])
    with c_info1:
        if nb_prefilled > 0:
            st.success(
                f"💡 **Mémorisation & Pré-remplissage actif :** {nb_prefilled} boîtier(s) sur {len(liste_gps_detectes)} "
                "ont été pré-remplis automatiquement selon vos configurations précédentes enregistrées."
            )
        else:
            st.caption(
                f"**{len(liste_gps_detectes)} boîtier(s) GPS actif(s) détecté(s)** sur cette séance. "
                "Associez chaque boîtier à son joueur :"
            )
    with c_btn_creer:
        if st.button("➕ Nouveau joueur", type="secondary", help="Créer et ajouter un joueur dans l'effectif RCA", key="btn_add_player_step2_top", use_container_width=True):
            dialog_creer_nouveau_joueur(liste_boitiers=liste_gps_detectes)
    with c_btn_trombi:
        if st.button("📸 Trombinoscope", type="secondary", help="Consulter et administrer les photos des joueurs", key="btn_trombi_step2_top", use_container_width=True):
            dialog_creer_nouveau_joueur(liste_boitiers=liste_gps_detectes)
    with c_info2:
        if st.button("🔄 Rétablir usine", help="Restaure les attributions initiales officielles du club RCA", key="btn_reset_default_factory", use_container_width=True):
            default_map = {str(k): v for k, v in DEFAULT_GPS_MAPPING.items()}
            st.session_state["mapping_gps_joueurs"] = default_map
            st.session_state["gps_mapping"] = default_map
            sauvegarder_mapping_gps_persistant(default_map, overwrite=True)
            for k in list(st.session_state.keys()):
                if k.startswith("gps_sel_boitier_"):
                    del st.session_state[k]
            st.toast("Attributions d'usine RCA restaurées.")
            st.rerun()

    mapping: Dict[Union[int, str], str] = {}
    col_gauche, col_droite = st.columns(2)
    split_index = (len(liste_gps_detectes) + 1) // 2

    # 3. Boucle STRICTE sur la liste des identifiants détectés (aucun range() utilisé)
    for idx, gps_id in enumerate(liste_gps_detectes):
        conteneur = col_gauche if idx < split_index else col_droite

        with conteneur:
            c_photo, c_select = st.columns([1, 3])

            prev = (
                saved_session.get(gps_id)
                or saved_session.get(str(gps_id))
                or reference_mapping.get(str(gps_id))
                or saved_disk.get(str(gps_id))
                or DEFAULT_GPS_MAPPING.get(gps_id)
            )
            default_idx = options.index(prev) if prev in options else 0

            def _enregistrer_changement_boitier(gps_id_target: int = gps_id):
                k_w = f"gps_sel_boitier_{gps_id_target}"
                val_w = st.session_state.get(k_w)
                if val_w and val_w not in [placeholder, TRIGGER_NOUVEAU_JOUEUR]:
                    sauvegarder_mapping_gps_persistant({str(gps_id_target): val_w})

            with c_select:
                choix = st.selectbox(
                    f"Boîtier GPS n°{gps_id}",
                    options=options,
                    index=default_idx,
                    format_func=format_affichage_joueur,
                    key=f"gps_sel_boitier_{gps_id}",
                    on_change=_enregistrer_changement_boitier,
                )

                if choix not in [placeholder, TRIGGER_NOUVEAU_JOUEUR]:
                    poste_d = joueurs_data.get(choix, "Inconnu")
                    grp_rca_defaut = mapper_poste_vers_groupe_rca(poste_d, choix)

                    # Initialisation de la mémoire de session pour les surcharges de postes
                    if "session_player_postes" not in st.session_state:
                        st.session_state["session_player_postes"] = {}

                    groupes_options = ["5 de devant", "3ème Ligne", "Trois-quarts"]
                    grp_actuel = st.session_state["session_player_postes"].get(choix, grp_rca_defaut)
                    if grp_actuel not in groupes_options:
                        grp_actuel = grp_rca_defaut

                    idx_grp = groupes_options.index(grp_actuel) if grp_actuel in groupes_options else 2

                    c_sel_grp, c_chk_persist = st.columns([1.7, 1.3])
                    with c_sel_grp:
                        nouveau_grp = st.selectbox(
                            f"Poste séance ({choix})",
                            options=groupes_options,
                            index=idx_grp,
                            key=f"session_grp_{gps_id}_{choix}",
                            help="Modifier en temps réel la catégorie de poste de ce joueur pour la séance en cours",
                            label_visibility="collapsed",
                        )
                        st.session_state["session_player_postes"][choix] = nouveau_grp
                    with c_chk_persist:
                        sauvegarder_defaut = st.checkbox(
                            "📌 Mémoriser",
                            key=f"chk_save_def_{gps_id}_{choix}",
                            help="Définir également ce groupe comme poste habituel par défaut dans l'effectif",
                        )
                        if sauvegarder_defaut and joueurs_data.get(choix) != nouveau_grp:
                            joueurs_data[choix] = nouveau_grp
                            sauvegarder_joueur_personnalise(choix, nouveau_grp, groupe=nouveau_grp)
                            st.toast(f"Poste par défaut de {choix} enregistré : {nouveau_grp}")

                    if nouveau_grp == "5 de devant":
                        badge_html = '<span style="background-color:#EFF6FF;color:#1D4ED8;border:1px solid #BFDBFE;padding:2px 7px;border-radius:5px;font-weight:700;font-size:0.75rem;">🏉 5 de devant</span>'
                    elif nouveau_grp == "3ème Ligne":
                        badge_html = '<span style="background-color:#F0FDF4;color:#15803D;border:1px solid #BBF7D0;padding:2px 7px;border-radius:5px;font-weight:700;font-size:0.75rem;">🏉 3ème Ligne</span>'
                    else:
                        badge_html = '<span style="background-color:#FEF2F2;color:#B91C1C;border:1px solid #FECACA;padding:2px 7px;border-radius:5px;font-weight:700;font-size:0.75rem;">🏉 Trois-quarts</span>'

                    st.markdown(
                        f"<div style='font-size:0.78rem;color:#475569;margin-top:-4px;margin-bottom:8px;'>"
                        f"Profil : <em>{poste_d}</em> &nbsp;➔&nbsp; Séance : {badge_html}</div>",
                        unsafe_allow_html=True,
                    )

            with c_photo:
                if choix not in [placeholder, TRIGGER_NOUVEAU_JOUEUR]:
                    photo = obtenir_chemin_photo(choix)
                    if photo and os.path.isfile(photo):
                        st.image(photo, width=48)
                    else:
                        avatar_b64 = get_player_image_b64(choix, size=48)
                        st.markdown(
                            f"<div style='width:48px;height:48px;border-radius:50%;overflow:hidden;border:2px solid #003399;display:flex;align-items:center;justify-content:center;box-shadow:0 2px 6px rgba(0,51,153,0.15);'>"
                            f"<img src='{avatar_b64}' style='width:48px;height:48px;object-fit:cover;' /></div>",
                            unsafe_allow_html=True,
                        )
                else:
                    st.markdown(
                        "<div style='width:48px;height:48px;border-radius:50%;"
                        "background:#E2E8F0;display:flex;align-items:center;"
                        "justify-content:center;color:#94A3B8;font-size:1.4rem;'>?</div>",
                        unsafe_allow_html=True,
                    )

            if choix not in [placeholder, TRIGGER_NOUVEAU_JOUEUR]:
                mapping[gps_id] = choix
                mapping[str(gps_id)] = choix
                st.session_state["gps_mapping"][gps_id] = choix
                st.session_state["gps_mapping"][str(gps_id)] = choix
                st.session_state["mapping_gps_joueurs"][gps_id] = choix
                st.session_state["mapping_gps_joueurs"][str(gps_id)] = choix

    # Synchronisation immédiate des attributions actives dans le stockage persistant
    if mapping:
        sauvegarder_mapping_gps_persistant(mapping)

    # 4. Sécurité anti-doublon stricte lors de l'attribution des GPS
    joueurs_assignes: Dict[str, List[int]] = {}
    for g_id in liste_gps_detectes:
        j_choisi = mapping.get(g_id)
        if j_choisi and j_choisi not in [placeholder, TRIGGER_NOUVEAU_JOUEUR]:
            joueurs_assignes.setdefault(j_choisi, []).append(g_id)

    doublons = {j: ids for j, ids in joueurs_assignes.items() if len(ids) > 1}
    has_doublons = len(doublons) > 0

    if has_doublons:
        details_doublons = ", ".join([
            f"« {j} » (boîtiers {', '.join(f'n°{i}' for i in ids)})"
            for j, ids in doublons.items()
        ])
        st.error(
            f"⛔ **Sécurité Anti-Doublon Active — Validation Impossible :**\n\n"
            f"Le(s) joueur(s) suivant(s) sont assignés à plusieurs boîtiers différents : {details_doublons}.\n\n"
            "Un même joueur ne peut pas être attribué à deux boîtiers GPS sur la même séance. "
            "Veuillez corriger ces doublons afin de pouvoir valider l'effectif."
        )

    # Récapitulatif d'attribution
    nb_assignes = len(joueurs_assignes)
    nb_total = len(liste_gps_detectes)
    if nb_assignes < nb_total and not has_doublons:
        st.warning(f"⚠️ **Attention :** {nb_total - nb_assignes} boîtier(s) sur {nb_total} ne sont pas encore attribués.")
    elif not has_doublons and nb_assignes > 0:
        st.success(f"✅ Tous les {nb_assignes} boîtiers GPS détectés sont attribués sans aucun doublon.")

    # --------------------------------------------------------------------------
    # 4.3 MATRICE INTERACTIVE DE PARTICIPATION AUX SPLITS (JOUEURS × ATELIERS)
    # --------------------------------------------------------------------------
    df_matrix_validated = None
    has_splits_detected = False
    training_splits_list: List[str] = []

    if not has_doublons and nb_assignes > 0 and df_raw is not None and not df_raw.empty:
        split_col_m, training_splits_list = extraire_splits_entrainement(df_raw)
        if training_splits_list:
            has_splits_detected = True
            st.markdown("<div style='height: 18px;'></div>", unsafe_allow_html=True)
            st.markdown(
                '<div style="background: linear-gradient(135deg, #0B192C 0%, #1F4E79 100%); padding: 12px 18px; border-radius: 10px; margin-bottom: 12px; border-left: 5px solid #D4AF37;">'
                '<div style="font-size: 1.12rem; font-weight: 800; color: #FFFFFF; display: flex; align-items: center; gap: 8px;">'
                '<span>⏱️</span> <span>Matrice de Participation aux Splits (Joueurs × Ateliers)</span>'
                '</div>'
                '<div style="font-size: 0.84rem; color: #E2E8F0; margin-top: 4px;">'
                'Par défaut, tous les joueurs assignés sont considérés comme participants à tous les ateliers (<strong>100% cochés</strong>). '
                'Si un joueur est sorti blessé, est arrivé en retard ou était ménagé sur un atelier, décochez sa case afin d\'aligner les calculs collectifs et individuels.'
                '</div>'
                '</div>',
                unsafe_allow_html=True,
            )

            # A. Système de détection automatique d'anomalies (0 m ou 0 s sur un split)
            anomalies = detecter_anomalies_splits(df_raw, mapping, split_col_m, training_splits_list)
            if anomalies:
                items_str = "\n".join([f"• {a['message']}. Décochez la case correspondante s'il n'a pas participé à cet exercice." for a in anomalies])
                st.warning(
                    f"⚠️ **Attention — Anomalie(s) de participation détectée(s) ({len(anomalies)}) :**\n\n"
                    f"{items_str}\n\n"
                    "💡 *Conseil : Vous pouvez utiliser le bouton « ⚡ Décocher automatiquement les splits à 0 m » ci-dessous pour ajuster la matrice en un clic.*"
                )
            else:
                st.success("✅ **Participation nominale :** Tous les joueurs assignés ont enregistré de la distance (> 0 m) sur chaque split d'entraînement.")

            # B. Barre d'outils et actions rapides
            c_btn_a, c_btn_b, c_info = st.columns([2.5, 2.5, 3])
            with c_btn_a:
                if st.button(
                    "⚡ Décocher automatiquement les splits à 0 m",
                    use_container_width=True,
                    key="btn_auto_uncheck_zero_splits",
                    help="Décoche automatiquement les splits où le joueur a 0 m ou n'a pas couru (anomalies).",
                ):
                    cache_choix = st.session_state.setdefault("splits_participation_cache", {})
                    ano_set = {(a["player"].strip().lower(), a["split"].strip().lower()) for a in anomalies}
                    for g_id, p_name in mapping.items():
                        if p_name and p_name not in [placeholder, TRIGGER_NOUVEAU_JOUEUR, "-"]:
                            for sp in training_splits_list:
                                if (p_name.strip().lower(), sp.strip().lower()) in ano_set:
                                    cache_choix[(p_name, sp)] = False
                                elif (p_name, sp) not in cache_choix:
                                    cache_choix[(p_name, sp)] = True
                    st.session_state["splits_matrix_version"] = st.session_state.get("splits_matrix_version", 0) + 1
                    st.toast("⚡ Splits à 0 m décochés automatiquement !")
                    st.rerun()

            with c_btn_b:
                if st.button(
                    "🔄 Tout cocher (100% de participation)",
                    use_container_width=True,
                    key="btn_check_all_splits",
                    help="Réinitialise toutes les cases à Vrai (100% de participation sur tous les exercices).",
                ):
                    cache_choix = st.session_state.setdefault("splits_participation_cache", {})
                    for g_id, p_name in mapping.items():
                        if p_name and p_name not in [placeholder, TRIGGER_NOUVEAU_JOUEUR, "-"]:
                            for sp in training_splits_list:
                                cache_choix[(p_name, sp)] = True
                    st.session_state["splits_matrix_version"] = st.session_state.get("splits_matrix_version", 0) + 1
                    st.toast("🔄 Tous les splits ont été cochés (100% de participation) !")
                    st.rerun()

            with c_info:
                st.markdown(
                    f"<div style='background:#F1F5F9;border:1px solid #CBD5E1;border-radius:8px;padding:8px 12px;text-align:center;font-size:0.85rem;font-weight:700;color:#1E293B;'>"
                    f"📊 {nb_assignes} joueur(s) • {len(training_splits_list)} splits détectés"
                    f"</div>",
                    unsafe_allow_html=True,
                )

            # C. Construction du DataFrame de la matrice et affichage avec st.data_editor
            df_mat = construire_matrice_splits(df_raw, mapping, split_col_m, training_splits_list)

            col_cfg = {
                "GPS": st.column_config.NumberColumn("Boîtier GPS", disabled=True, format="%d", help="Numéro du boîtier GPS"),
                "Joueur": st.column_config.TextColumn("Joueur", disabled=True, help="Nom du joueur officiel"),
                "Groupe": st.column_config.TextColumn("Groupe / Poste", disabled=True, help="Ligne officielle RCA"),
                "Distance brute (m)": st.column_config.NumberColumn("Distance brute (m)", disabled=True, format="%d m", help="Cumul brut enregistré sur les splits d'entraînement"),
            }
            for sp in training_splits_list:
                col_cfg[sp] = st.column_config.CheckboxColumn(
                    label=sp,
                    help=f"Participation de ce joueur au {sp}",
                    default=True,
                )

            mat_version = st.session_state.get("splits_matrix_version", 0)
            edited_mat = st.data_editor(
                df_mat,
                key=f"splits_data_editor_{mat_version}",
                column_config=col_cfg,
                disabled=["GPS", "Joueur", "Groupe", "Distance brute (m)"],
                hide_index=True,
                use_container_width=True,
            )

            # Synchronisation continue du cache de choix
            cache_choix = st.session_state.setdefault("splits_participation_cache", {})
            for _, r in edited_mat.iterrows():
                p_n = str(r["Joueur"])
                for sp in training_splits_list:
                    if sp in r:
                        cache_choix[(p_n, sp)] = bool(r[sp])

            df_matrix_validated = edited_mat
            st.session_state["splits_matrix_validated"] = edited_mat

    # 5. Bouton de validation explicite de l'Étape 2 (Sauvegarde dans st.session_state et fichier JSON)
    st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
    c_btn1, c_btn2, c_btn3 = st.columns([1, 2.5, 1])
    with c_btn2:
        btn_label = "✅ Valider l'effectif et la participation aux Splits" if has_splits_detected else "✅ Valider l'effectif et l'association GPS"
        btn_valider = st.button(
            btn_label,
            type="primary",
            disabled=has_doublons or nb_assignes == 0,
            use_container_width=True,
            key="btn_valider_effectif_step2",
            help="Valide définitivement l'attribution des boîtiers GPS et la matrice de participation aux splits, les enregistre en session et dans la configuration persistante.",
        )
        if btn_valider:
            if has_doublons:
                st.error("⛔ Validation bloquée : veuillez résoudre les doublons avant de continuer.")
            elif nb_assignes == 0:
                st.warning("⚠️ Veuillez attribuer au moins un boîtier GPS avant de valider.")
            else:
                st.session_state["step2_validated"] = True
                st.session_state["gps_mapping"] = mapping
                st.session_state["mapping_gps_joueurs"] = mapping
                if df_matrix_validated is not None:
                    st.session_state["splits_matrix_validated"] = df_matrix_validated
                cache_splits = st.session_state.get("splits_participation_cache", {})
                sauvegarder_mapping_gps_persistant(
                    mapping,
                    last_file_sig=st.session_state.get("uploaded_file_sig"),
                    splits_cache=cache_splits,
                )
                st.toast("✅ Effectif et participation aux splits validés avec succès !")
                st.rerun()

    return mapping, has_doublons


def appliquer_mapping(df_catapult: pd.DataFrame, dictionnaire_gps: Dict[Union[int, str], str]) -> pd.DataFrame:
    """
    Remplace les numéros de boîtiers GPS dans la colonne joueur
    par les noms officiels sélectionnés de l'effectif.
    Gère de façon robuste les types int, float (1.0), et str ('1', '1.0').
    """
    if not dictionnaire_gps:
        return df_catapult

    df = df_catapult.copy()
    player_col = find_column(df, [
        "Player Name", "Athlete Name", "Nom ; Prénom", "Player", "Athlete", "Joueur", "Nom", "Name"
    ])
    if not player_col:
        return df

    # Normalisation du dictionnaire avec toutes les variantes de clés
    dict_map: Dict[Any, str] = {}
    for k, v in dictionnaire_gps.items():
        if not v:
            continue
        dict_map[k] = v
        dict_map[str(k).strip()] = v
        try:
            k_num = float(str(k).strip())
            if k_num.is_integer():
                k_int = int(k_num)
                dict_map[k_int] = v
                dict_map[float(k_int)] = v
                dict_map[str(k_int)] = v
                dict_map[f"{k_int}.0"] = v
        except (ValueError, TypeError):
            pass

    def _remplacer_id(val: Any) -> Any:
        if pd.isna(val) or val is None:
            return val
        if val in dict_map:
            return dict_map[val]
        val_str = str(val).strip()
        if val_str in dict_map:
            return dict_map[val_str]
        val_clean = re.sub(r"\.0$", "", val_str)
        if val_clean in dict_map:
            return dict_map[val_clean]
        try:
            n = float(val_str)
            if n.is_integer() and int(n) in dict_map:
                return dict_map[int(n)]
        except (ValueError, TypeError):
            pass
        return val

    df[player_col] = df[player_col].apply(_remplacer_id)
    return df


# ==============================================================================
# 5. CALCULS STAPS & INGÉNIERIE DES CARACTÉRISTIQUES (GROUPBY JOUEUR)
# ==============================================================================

def load_template_positions() -> Dict[str, str]:
    """Extrait les postes du classeur modèle officiel si présent."""
    pos_map: Dict[str, str] = {}
    template_path = get_template_path()
    if not template_path:
        return pos_map

    try:
        wb = openpyxl.load_workbook(template_path, data_only=True)
        if 'Liste joueurs' in wb.sheetnames:
            ws = wb['Liste joueurs']
            cur_grp: Optional[str] = None
            for row in ws.iter_rows(min_row=2, max_row=120):
                c0 = str(row[0].value or '').strip().lower()
                c1 = str(row[1].value or '').strip()
                if c0 and c0 != 'moyenne':
                    cur_grp = mapper_poste_vers_groupe_rca(c0)
                if c1 and c1.lower() != '1er mi temps ' and cur_grp:
                    pos_map[c1.lower()] = cur_grp
    except Exception:
        pass
    return pos_map


def assign_player_groups(df: pd.DataFrame, df_raw: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """
    Assigne le groupe athlétique officiel ('5 de devant', '3ème Ligne', 'Trois-quarts') :
    1. En priorité absolue via le dictionnaire officiel 'joueurs_data' (effectif sénior RCA).
    2. Puis via les métadonnées brutes Catapult.
    3. Puis via le modèle Excel.
    4. Répartition de secours équilibrée.
    Conserve également le poste rugby précis dans la colonne 'Poste_Detaille'.
    """
    df = df.copy()
    player_col = "Nom ; Prénom"
    GROUPS = ["5 de devant", "3ème Ligne", "Trois-quarts"]
    template_positions = load_template_positions()
    unique_players = sorted(df[player_col].dropna().unique().tolist())

    group_list = []
    poste_detaille_list = []
    for idx, p in enumerate(df[player_col]):
        p_clean = str(p).strip()
        p_lower = p_clean.lower()

        # 0. Surcharge en temps réel spécifique à la séance (Étape 2 - Attribution GPS)
        session_overrides = st.session_state.get("session_player_postes", {}) if hasattr(st, "session_state") else {}
        matched_override = None
        for s_nom, s_grp in session_overrides.items():
            if s_nom and (s_nom.lower() == p_lower or s_nom.lower() in p_lower or p_lower in s_nom.lower()):
                matched_override = s_grp
                break

        if matched_override:
            group_list.append(matched_override)
            poste_detaille_list.append(matched_override)
            continue

        # 1. Correspondance avec la base de données effectif RCA
        matched_poste = None
        for j_nom, j_poste in joueurs_data.items():
            if j_nom.lower() == p_lower or j_nom.lower() in p_lower or p_lower in j_nom.lower():
                matched_poste = j_poste
                break

        if matched_poste:
            grp = mapper_poste_vers_groupe_rca(matched_poste, nom_joueur=p_clean)
            poste_detaille_list.append(matched_poste)
            group_list.append(grp)
            continue

        # 2. Métadonnées brutes Catapult (si présentes dans df_raw)
        raw_poste = None
        if df_raw is not None and not df_raw.empty:
            p_raw_col = find_column(df_raw, ["Position Name", "Position", "Poste", "Pos"])
            name_raw_col = find_column(df_raw, ["Nom ; Prénom", "Player Name", "Athlete Name", "Player", "Athlete", "Joueur", "Nom", "Name"])
            if p_raw_col and name_raw_col and p_raw_col in df_raw.columns and name_raw_col in df_raw.columns:
                sub_r = df_raw[df_raw[name_raw_col].astype(str).str.strip().str.lower() == p_lower]
                if not sub_r.empty and pd.notna(sub_r[p_raw_col].iloc[0]):
                    raw_poste = str(sub_r[p_raw_col].iloc[0]).strip()

        if raw_poste:
            grp = mapper_poste_vers_groupe_rca(raw_poste, nom_joueur=p_clean)
            poste_detaille_list.append(raw_poste)
            group_list.append(grp)
            continue

        # 3. Modèle Excel
        if p_lower in template_positions:
            t_pos = template_positions[p_lower]
            grp = mapper_poste_vers_groupe_rca(t_pos, nom_joueur=p_clean)
            poste_detaille_list.append(t_pos)
            group_list.append(grp)
            continue

        matched_grp = None
        for t_name, t_grp in template_positions.items():
            if t_name in p_lower or p_lower in t_name:
                matched_grp = t_grp
                break
        if matched_grp:
            grp = mapper_poste_vers_groupe_rca(matched_grp, nom_joueur=p_clean)
            poste_detaille_list.append(matched_grp)
            group_list.append(grp)
            continue

        # 4. Repli déterministe
        p_idx = unique_players.index(p) if p in unique_players else idx
        fallback_grp = GROUPS[p_idx % 3]
        poste_detaille_list.append(fallback_grp)
        group_list.append(fallback_grp)

    df["Poste_Detaille"] = poste_detaille_list
    df["Groupe"] = group_list
    return df


def process_catapult_data(
    df_raw: pd.DataFrame,
    split_mode: str = "splits",
    split_filter: Optional[Union[str, List[str]]] = None,
    exclude_warmup: bool = False,
) -> pd.DataFrame:
    """
    Traite le DataFrame brut Catapult One selon les méthodes STAPS de la préparatrice physique RCA :
    - Par défaut ('splits') : Exclut la ligne globale 'all' et somme uniquement les ateliers réels (Split 1..N).
    - Mode 'all' : Retient la ligne globale 'all' (séance brute complète avec temps morts).
    - HSR (m) : Correspond strictement à la Zone 3 (Z3_km * 1000.0) pour jauges et graphiques de groupe.
    - VHSR (m) : Correspond strictement au cumul Zone 4 + Zone 5 ((Z4_km + Z5_km) * 1000.0).
    - Haute intensité Podium (m) : Cumul total Zone 3 + Zone 4 + Zone 5 ((Z3_km + Z4_km + Z5_km) * 1000.0).
    - ACC / DEC n° (> 3 m/s²) : Somme des 4 zones d'accélérations et décélérations intenses (> 3 m/s²).
    - Distance Per Min (m/min) : Moyenne des splits d'ateliers.
    - Convertit le Temps de jeu en minutes réelles.
    """
    player_col = find_column(df_raw, [
        "Nom ; Prénom", "Player Name", "Athlete Name", "Player", "Athlete", "Joueur", "Nom", "Name"
    ])
    if not player_col:
        raise ValueError("Colonne Joueur introuvable dans le fichier. Colonnes attendues : 'Player Name', 'Athlete Name', etc.")

    # 1. Gestion des lignes de session et des Splits (Méthode Staff RCA vs Ligne all)
    split_col = find_column(df_raw, [
        "Split Name", "Split", "Period Name", "Period", "Tags", "Tag", "Activity Name", "Session Name", "Session Title", "Segment"
    ])

    df_working = df_raw.copy()

    if split_col and split_col in df_working.columns:
        s_series = df_working[split_col].astype(str).str.strip()
        s_lower = s_series.str.lower()

        # A. Filtrage explicite si demandé (ex: multiselect d'ateliers spécifiques)
        if split_filter is not None:
            if isinstance(split_filter, str):
                mask = s_lower.str.contains(split_filter.strip().lower(), na=False)
                if mask.any():
                    df_working = df_working[mask].copy()
            elif isinstance(split_filter, (list, tuple, set)):
                targets = [str(x).strip().lower() for x in split_filter]
                mask = s_lower.isin(targets)
                if mask.any():
                    df_working = df_working[mask].copy()

        # B. Filtrage selon le mode choisi (Splits ateliers par défaut vs Ligne all)
        else:
            GLOBAL_TAGS = {
                "session", "whole session", "full session", "all", "total",
                "seance", "séance", "globale", "global", "complete", "complète",
                "session summary", "all session", "overall"
            }
            is_global = s_lower.isin(GLOBAL_TAGS) | s_lower.str.startswith("session ") | (s_lower == "session")
            has_global_rows = is_global.any()
            has_split_rows = (~is_global & (s_series != "") & (s_series.str.lower() != "nan")).any()

            norm_mode = str(split_mode).lower().strip()
            use_splits_aggregation = ("split" in norm_mode or "atelier" in norm_mode or norm_mode == "default") and has_split_rows

            if use_splits_aggregation and has_global_rows:
                # Mode Préparatrice Physique : Exclure la ligne globale "all" pour agréger uniquement les ateliers réels
                players_with_splits = set(df_working.loc[~is_global, player_col].astype(str).str.strip().unique())
                mask_safe = (~is_global) | (~df_working[player_col].astype(str).str.strip().isin(players_with_splits))
                df_working = df_working[mask_safe].copy()
            elif "all" in norm_mode and has_global_rows:
                # Mode Séance Brute Complète : Retenir la ligne globale "all"
                players_with_global = set(df_working.loc[is_global, player_col].astype(str).str.strip().unique())
                mask_safe = is_global | (~df_working[player_col].astype(str).str.strip().isin(players_with_global))
                df_working = df_working[mask_safe].copy()

        # C. Exclusion optionnelle de l'échauffement
        if exclude_warmup:
            warmup_mask = df_working[split_col].astype(str).str.lower().str.contains("echauff|échauff|warmup|warm-up|warm_up", na=False)
            if (~warmup_mask).any():
                df_working = df_working[~warmup_mask].copy()

    dur_col = find_column(df_working, [
        "Duration", "Temps de jeu", "Duration (mins)", "Duration (s)", "Duration (sec)", "Durée", "Time", "Session Duration"
    ])
    dist_col = find_column(df_working, ["Distance (km)", "Total Distance (km)", "Distance"])
    dist_m_col = find_column(df_working, ["Distance (m)", "Total Distance (m)"])

    hsr_col = find_column(df_working, [
        "Distance in Speed Zone 3 (km)", "Distance in Speed Zone 3  (km)",
        "Distance haute intensité (m)", "High Speed Running (m)", "HSR Distance (m)", "HSR Distance (>5m/s)"
    ])
    vhsr_col = find_column(df_working, [
        "Distance in Speed Zone 4 (km)", "Distance in Speed Zone 4  (km)",
        "Distance in Speed Zone 5 (km)", "Distance in Speed Zone 5  (km)",
        "Sprint Distance (m)", "VHSR Distance (m)", "Very High Speed Running (m)"
    ])

    top_speed_col = find_column(df_working, ["Top Speed (km/h)", "Max Velocity (km/h)", "Top Speed", "Max Velocity", "Vmax"])
    hr_max_col = find_column(df_working, ["Hr Max (bpm)", "Max Heart Rate (bpm)", "HR Max", "Max HR", "Heart Rate Max", "Hr Max"])
    dpm_col = find_column(df_working, ["Distance Per Min (m/min)", "Meterage Per Minute", "m/min", "Distance per min", "Dist/min"])

    max_accel_col = find_column(df_working, ["Max Acceleration (m/s/s)", "Max Acceleration (m/s2)", "Max Acceleration", "Max Accel"])
    max_decel_col = find_column(df_working, ["Max Deceleration (m/s/s)", "Max Deceleration (m/s2)", "Max Deceleration", "Max Decel"])

    z1_col = find_column(df_working, ["Distance in Speed Zone 1 (km)", "Distance in Speed Zone 1  (km)", "Speed Zone 1", "Zone 1"])
    z2_col = find_column(df_working, ["Distance in Speed Zone 2 (km)", "Distance in Speed Zone 2  (km)", "Speed Zone 2", "Zone 2"])
    z3_col = find_column(df_working, ["Distance in Speed Zone 3 (km)", "Distance in Speed Zone 3  (km)", "Speed Zone 3", "Zone 3"])
    z4_col = find_column(df_working, ["Distance in Speed Zone 4 (km)", "Distance in Speed Zone 4  (km)", "Speed Zone 4", "Zone 4"])
    z5_col = find_column(df_working, ["Distance in Speed Zone 5 (km)", "Distance in Speed Zone 5  (km)", "Speed Zone 5", "Zone 5"])

    acc_34_col = find_column(df_working, ["Accelerations Zone Count: 3 - 4 m/s/s", "Accelerations Zone Count: 3-4 m/s/s", "Accel 3-4", "Accelerations (>3m/s2)"])
    acc_gt4_col = find_column(df_working, ["Accelerations Zone Count: > 4 m/s/s", "Accelerations Zone Count: >4 m/s/s", "Accel >4"])
    dec_34_col = find_column(df_working, ["Deceleration Zone Count: 3 - 4 m/s/s", "Deceleration Zone Count: 3-4 m/s/s", "Decel 3-4", "Décélérations (<-3m/s2)"])
    dec_gt4_col = find_column(df_working, ["Deceleration Zone Count: > 4 m/s/s", "Deceleration Zone Count: >4 m/s/s", "Decel >4"])
    impacts_col = find_column(df_working, ["Impacts", "Total Impacts"])
    player_load_col = find_column(df_working, ["Player Load", "PlayerLoad", "Total Player Load", "Player_Load", "Charge"])

    df_clean = pd.DataFrame()
    df_clean["Nom ; Prénom"] = df_working[player_col].astype(str).str.strip()

    if dur_col:
        df_clean["Duration_min"] = parse_duration_series(df_working[dur_col], col_name=dur_col)
    else:
        df_clean["Duration_min"] = 0.0

    if dist_col:
        vals = df_working[dist_col].apply(clean_numeric_value)
        if "(m)" in dist_col.lower() or vals.median() > 100.0:
            df_clean["Distance_km"] = vals / 1000.0
        else:
            df_clean["Distance_km"] = vals
    elif dist_m_col:
        df_clean["Distance_km"] = df_working[dist_m_col].apply(clean_numeric_value) / 1000.0
    else:
        df_clean["Distance_km"] = 0.0

    df_clean["Top_Speed"] = df_working[top_speed_col].apply(clean_numeric_value) if top_speed_col else 0.0
    df_clean["Hr_Max"] = df_working[hr_max_col].apply(clean_numeric_value) if hr_max_col else 0.0
    df_clean["Distance_Per_Min"] = df_working[dpm_col].apply(clean_numeric_value) if dpm_col else 0.0
    df_clean["Max_Accel"] = df_working[max_accel_col].apply(clean_numeric_value) if max_accel_col else 0.0
    df_clean["Max_Decel"] = df_working[max_decel_col].apply(lambda x: abs(clean_numeric_value(x))) if max_decel_col else 0.0
    df_clean["Impacts"] = df_working[impacts_col].apply(clean_numeric_value) if impacts_col else 0.0
    df_clean["Player_Load"] = df_working[player_load_col].apply(clean_numeric_value) if player_load_col else 0.0

    for z_name, c_source in [("Z1_km", z1_col), ("Z2_km", z2_col), ("Z3_km", z3_col), ("Z4_km", z4_col), ("Z5_km", z5_col)]:
        if c_source:
            vals = df_working[c_source].apply(clean_numeric_value)
            if "(m)" in c_source.lower() or vals.median() > 20.0:
                vals = vals / 1000.0
            df_clean[z_name] = vals
        else:
            df_clean[z_name] = 0.0

    # Si z3_col était absent mais hsr_col présent (ex: "HSR Distance (>5m/s)")
    if not z3_col and hsr_col:
        v_hsr = df_working[hsr_col].apply(clean_numeric_value)
        val_hsr_km = (v_hsr / 1000.0) if "(m)" in hsr_col.lower() or v_hsr.median() > 10.0 else v_hsr
        df_clean["Z3_km"] = val_hsr_km

    # Nomenclature officielle préparatrice physique RCA :
    # 1. HSR (m) correspond strictement à la Zone 3
    df_clean["HSR_m"] = df_clean["Z3_km"] * 1000.0

    # 2. VHSR (m) correspond strictement au cumul Zone 4 + Zone 5
    df_clean["Z4_m"] = df_clean["Z4_km"] * 1000.0
    df_clean["Z5_m"] = df_clean["Z5_km"] * 1000.0
    df_clean["VHSR_m"] = (df_clean["Z4_km"] + df_clean["Z5_km"]) * 1000.0

    # Si z4_col et z5_col étaient absents mais vhsr_col présent (ex: "Sprint Distance (m)")
    if not z4_col and not z5_col and vhsr_col:
        v_vhsr = df_working[vhsr_col].apply(clean_numeric_value)
        val_vhsr_m = v_vhsr if "(m)" in vhsr_col.lower() or v_vhsr.median() > 5.0 else v_vhsr * 1000.0
        df_clean["VHSR_m"] = val_vhsr_m
        df_clean["Z4_m"] = val_vhsr_m * 0.7
        df_clean["Z5_m"] = val_vhsr_m * 0.3
        df_clean["Z4_km"] = df_clean["Z4_m"] / 1000.0
        df_clean["Z5_km"] = df_clean["Z5_m"] / 1000.0

    df_clean["Acc_34"] = df_working[acc_34_col].apply(clean_numeric_value) if acc_34_col else 0.0
    df_clean["Acc_gt4"] = df_working[acc_gt4_col].apply(clean_numeric_value) if acc_gt4_col else 0.0
    df_clean["Dec_34"] = df_working[dec_34_col].apply(clean_numeric_value) if dec_34_col else 0.0
    df_clean["Dec_gt4"] = df_working[dec_gt4_col].apply(clean_numeric_value) if dec_gt4_col else 0.0

    # Groupby Joueur
    agg_dict = {
        "Duration_min": "sum",
        "Distance_km": "sum",
        "HSR_m": "sum",
        "VHSR_m": "sum",
        "Z4_m": "sum",
        "Z5_m": "sum",
        "Top_Speed": "max",
        "Hr_Max": "max",
        "Distance_Per_Min": "mean",
        "Z1_km": "sum",
        "Z2_km": "sum",
        "Z3_km": "sum",
        "Z4_km": "sum",
        "Z5_km": "sum",
        "Max_Decel": "max",
        "Max_Accel": "max",
        "Acc_34": "sum",
        "Acc_gt4": "sum",
        "Dec_34": "sum",
        "Dec_gt4": "sum",
        "Impacts": "sum",
        "Player_Load": "sum",
    }
    grouped = df_clean.groupby("Nom ; Prénom", as_index=False).agg(agg_dict)

    grouped["Duration_min"] = grouped["Duration_min"].astype(float)
    pos_dur = grouped["Duration_min"][grouped["Duration_min"] > 0]
    if len(pos_dur) > 0:
        med_dur = pos_dur.median()
        if med_dur > 100000:
            grouped["Duration_min"] = grouped["Duration_min"] / 60000.0
        elif med_dur > 240:
            grouped["Duration_min"] = grouped["Duration_min"] / 60.0

    temps_jeu = grouped["Duration_min"].astype(float)
    d_total_m = grouped["Distance_km"] * 1000.0
    dist_tot_km = grouped["Distance_km"]

    # Distance Per Min : moyenne des splits si disponible, sinon d_total_m / temps_jeu
    if grouped["Distance_Per_Min"].max() > 0:
        dist_per_min = grouped["Distance_Per_Min"].values
    else:
        dist_per_min = np.where(temps_jeu > 0, d_total_m / temps_jeu, 0.0)

    hsr_per_min = np.where(temps_jeu > 0, grouped["HSR_m"] / temps_jeu, 0.0)

    # ACC / DEC n° (> 3 m/s²) : somme exacte des 4 zones
    accel_decel_total = grouped["Acc_34"] + grouped["Acc_gt4"] + grouped["Dec_34"] + grouped["Dec_gt4"]

    # % de marche directement calculé sur les données agrégées selon le mode actif
    pct_marche = np.where(dist_tot_km > 0, (grouped["Z1_km"] / dist_tot_km) * 100.0, 0.0)
    pct_course = np.where(dist_tot_km > 0, (grouped["Z2_km"] / dist_tot_km) * 100.0, 0.0)
    pct_hsr = np.where(dist_tot_km > 0, (grouped["Z3_km"] / dist_tot_km) * 100.0, 0.0)
    pct_vhsr = np.where(dist_tot_km > 0, ((grouped["Z4_km"] + grouped["Z5_km"]) / dist_tot_km) * 100.0, 0.0)

    # Haute intensité Podium & Tableau : cumul total Zone 3 + Zone 4 + Zone 5 (> 14.4 km/h) (=SUM(I:K)*1000)
    haute_intensite_podium = (grouped["Z3_km"] + grouped["Z4_km"] + grouped["Z5_km"]) * 1000.0

    df_indiv = pd.DataFrame()
    df_indiv['Nom ; Prénom'] = grouped['Nom ; Prénom']
    df_indiv['Temps de jeu '] = temps_jeu.astype(float)
    df_indiv['Distance (km)'] = grouped['Distance_km']
    df_indiv[' Distance haute intensité (m)'] = haute_intensite_podium
    df_indiv['Haute intensité Podium (m)'] = haute_intensite_podium
    df_indiv['HSR Z3 (m)'] = grouped['HSR_m']
    df_indiv['HSR_m'] = grouped['HSR_m']
    df_indiv['Top Speed (km/h)'] = grouped['Top_Speed']
    df_indiv['Distance Per Min (m/min)'] = dist_per_min
    df_indiv['Distance in Speed Zone 1  (km)'] = grouped['Z1_km']
    df_indiv['Distance in Speed Zone 2  (km)'] = grouped['Z2_km']
    df_indiv['Distance in Speed Zone 3  (km)'] = grouped['Z3_km']
    df_indiv['Distance in Speed Zone 4  (km)'] = grouped['Z4_km']
    df_indiv['Distance in Speed Zone 5  (km)'] = grouped['Z5_km']
    df_indiv['Max Deceleration (m/s/s)'] = grouped['Max_Decel']
    df_indiv['Max Acceleration (m/s/s)'] = grouped['Max_Accel']
    df_indiv['Accelerations Zone Count: 3 - 4 m/s/s'] = grouped['Acc_34']
    df_indiv['Accelerations Zone Count: > 4 m/s/s'] = grouped['Acc_gt4']
    df_indiv['Deceleration Zone Count: 3 - 4 m/s/s'] = grouped['Dec_34']
    df_indiv['Deceleration Zone Count: > 4 m/s/s'] = grouped['Dec_gt4']
    df_indiv['Accel/Dec'] = accel_decel_total
    df_indiv['%Marche '] = pct_marche
    df_indiv['%course'] = pct_course
    df_indiv['%HSR '] = pct_hsr
    df_indiv['%VHSR'] = pct_vhsr
    df_indiv['HSR m/min'] = hsr_per_min
    df_indiv['VHSR (m)'] = grouped['VHSR_m']
    df_indiv['Player Load'] = grouped['Player_Load']
    if df_indiv['Player Load'].max() == 0:
        df_indiv['Player Load'] = (df_indiv['Distance (km)'] * 85.0 + df_indiv['Accel/Dec'] * 1.5).round(1)
    if hr_max_col:
        df_indiv['Hr Max (bpm)'] = grouped['Hr_Max']

    df_indiv = assign_player_groups(df_indiv, df_raw=df_working)
    return df_indiv


# ==============================================================================
# 6. MOTEUR DE CALCULS DYNAMIQUES (JAUGES ET PODIUM)
# ==============================================================================

def calculer_moyennes_groupes_rca(
    df_raw: pd.DataFrame,
    df_indiv: pd.DataFrame,
    split_mode: str = "splits",
    split_filter: Optional[Union[str, List[str]]] = None,
    exclude_warmup: bool = False,
) -> Dict[str, Dict[str, float]]:
    """
    Calcule les moyennes par groupe selon la méthodologie officielle de la préparatrice physique RCA :
    En mode 'Splits' :
    1. Filtre les splits inactifs où Distance (km) <= 0 (ou durée <= 0) pour chaque joueur.
    2. Pour chaque groupe ('5 de devant', '3ème Ligne', 'Trois-quarts') et chaque Split,
       calcule la moyenne des joueurs actifs sur ce split.
    3. Additionne ces moyennes sur l'ensemble des splits de la séance pour obtenir :
       - Distance Moyenne (m)
       - HSR Moyen (m) (strictement Zone 3)
       - VHSR Moyen (m) (Zone 4 + Zone 5)
       - Accel / Decel moyennes (> 3 m/s²)
       - Moyenne des m/min (Distance Per Min)
    """
    GROUPS = ["5 de devant", "3ème Ligne", "Trois-quarts"]
    res: Dict[str, Dict[str, float]] = {}

    player_col = find_column(df_raw, [
        "Nom ; Prénom", "Player Name", "Athlete Name", "Player", "Athlete", "Joueur", "Nom", "Name"
    ])
    split_col = find_column(df_raw, [
        "Split Name", "Split", "Period Name", "Period", "Tags", "Tag", "Activity Name", "Session Name", "Session Title", "Segment"
    ])

    player_group_map = {}
    if not df_indiv.empty and "Nom ; Prénom" in df_indiv.columns and "Groupe" in df_indiv.columns:
        player_group_map = dict(zip(
            df_indiv["Nom ; Prénom"].astype(str).str.strip(),
            df_indiv["Groupe"].astype(str).str.strip()
        ))

    norm_mode = str(split_mode).lower().strip()
    use_splits = ("split" in norm_mode or "atelier" in norm_mode or norm_mode == "default" or norm_mode == "manual") and split_col and split_col in df_raw.columns

    if use_splits:
        df_w = df_raw.copy()

        # Mapping du groupe pour chaque ligne
        def _get_group(p_val: Any) -> str:
            p_str = str(p_val).strip()
            if p_str in player_group_map:
                return player_group_map[p_str]
            for k, v in player_group_map.items():
                if k.lower() == p_str.lower():
                    return v
            return "Non défini"

        df_w["_Groupe"] = df_w[player_col].apply(_get_group) if player_col and player_col in df_w.columns else "Non défini"

        s_series = df_w[split_col].astype(str).str.strip()
        s_lower = s_series.str.lower()

        # Filtrage manuel ou exclusion de la ligne globale 'all'
        if split_filter is not None:
            if isinstance(split_filter, str):
                mask = s_lower.str.contains(split_filter.strip().lower(), na=False)
                if mask.any():
                    df_w = df_w[mask].copy()
            elif isinstance(split_filter, (list, tuple, set)):
                targets = [str(x).strip().lower() for x in split_filter]
                mask = s_lower.isin(targets)
                if mask.any():
                    df_w = df_w[mask].copy()
        else:
            GLOBAL_TAGS = {
                "session", "whole session", "full session", "all", "total",
                "seance", "séance", "globale", "global", "complete", "complète",
                "session summary", "all session", "overall"
            }
            is_global = s_lower.isin(GLOBAL_TAGS) | s_lower.str.startswith("session ") | (s_lower == "session")
            has_split_rows = (~is_global & (s_series != "") & (s_series != "nan")).any()
            if has_split_rows:
                df_w = df_w[~is_global].copy()

        if exclude_warmup:
            warmup_mask = df_w[split_col].astype(str).str.lower().str.contains("echauff|échauff|warmup|warm-up|warm_up", na=False)
            if (~warmup_mask).any():
                df_w = df_w[~warmup_mask].copy()

        dist_col = find_column(df_w, ["Distance (km)", "Total Distance (km)", "Distance"])
        dist_m_col = find_column(df_w, ["Distance (m)", "Total Distance (m)"])
        z3_col = find_column(df_w, ["Distance in Speed Zone 3 (km)", "Distance in Speed Zone 3  (km)", "Speed Zone 3", "Zone 3"])
        z4_col = find_column(df_w, ["Distance in Speed Zone 4 (km)", "Distance in Speed Zone 4  (km)", "Speed Zone 4", "Zone 4"])
        z5_col = find_column(df_w, ["Distance in Speed Zone 5 (km)", "Distance in Speed Zone 5  (km)", "Speed Zone 5", "Zone 5"])
        hsr_col = find_column(df_w, ["Distance haute intensité (m)", "High Speed Running (m)", "HSR Distance (m)", "HSR Distance (>5m/s)"])
        vhsr_col = find_column(df_w, ["Sprint Distance (m)", "VHSR Distance (m)", "Very High Speed Running (m)"])
        dpm_col = find_column(df_w, ["Distance Per Min (m/min)", "Meterage Per Minute", "m/min", "Distance per min", "Dist/min"])

        acc_34_col = find_column(df_w, ["Accelerations Zone Count: 3 - 4 m/s/s", "Accelerations Zone Count: 3-4 m/s/s", "Accel 3-4", "Accelerations (>3m/s2)"])
        acc_gt4_col = find_column(df_w, ["Accelerations Zone Count: > 4 m/s/s", "Accelerations Zone Count: >4 m/s/s", "Accel >4"])
        dec_34_col = find_column(df_w, ["Deceleration Zone Count: 3 - 4 m/s/s", "Deceleration Zone Count: 3-4 m/s/s", "Decel 3-4", "Décélérations (<-3m/s2)"])
        dec_gt4_col = find_column(df_w, ["Deceleration Zone Count: > 4 m/s/s", "Deceleration Zone Count: >4 m/s/s", "Decel >4"])

        if dist_col and dist_col in df_w.columns:
            vals = df_w[dist_col].apply(clean_numeric_value)
            df_w["_Dist_m"] = vals if ("(m)" in dist_col.lower() or vals.median() > 100.0) else vals * 1000.0
        elif dist_m_col and dist_m_col in df_w.columns:
            df_w["_Dist_m"] = df_w[dist_m_col].apply(clean_numeric_value)
        else:
            df_w["_Dist_m"] = 0.0

        if z3_col and z3_col in df_w.columns:
            vals = df_w[z3_col].apply(clean_numeric_value)
            df_w["_Z3_m"] = vals if ("(m)" in z3_col.lower() or vals.median() > 20.0) else vals * 1000.0
        elif hsr_col and hsr_col in df_w.columns:
            vals = df_w[hsr_col].apply(clean_numeric_value)
            df_w["_Z3_m"] = vals if ("(m)" in hsr_col.lower() or vals.median() > 10.0) else vals * 1000.0
        else:
            df_w["_Z3_m"] = 0.0

        z4_m = pd.Series(0.0, index=df_w.index)
        z5_m = pd.Series(0.0, index=df_w.index)
        if z4_col and z4_col in df_w.columns:
            v4 = df_w[z4_col].apply(clean_numeric_value)
            z4_m = v4 if ("(m)" in z4_col.lower() or v4.median() > 20.0) else v4 * 1000.0
        if z5_col and z5_col in df_w.columns:
            v5 = df_w[z5_col].apply(clean_numeric_value)
            z5_m = v5 if ("(m)" in z5_col.lower() or v5.median() > 20.0) else v5 * 1000.0

        if (not z4_col or not z5_col) and vhsr_col and vhsr_col in df_w.columns:
            v_vhsr = df_w[vhsr_col].apply(clean_numeric_value)
            df_w["_VHSR_m"] = v_vhsr if ("(m)" in vhsr_col.lower() or v_vhsr.median() > 5.0) else v_vhsr * 1000.0
        else:
            df_w["_VHSR_m"] = z4_m + z5_m

        acc_dec = pd.Series(0.0, index=df_w.index)
        for c in [acc_34_col, acc_gt4_col, dec_34_col, dec_gt4_col]:
            if c and c in df_w.columns:
                acc_dec += df_w[c].apply(clean_numeric_value)
        df_w["_Acc_Dec"] = acc_dec

        if dpm_col and dpm_col in df_w.columns:
            df_w["_Dpm"] = df_w[dpm_col].apply(clean_numeric_value)
        else:
            df_w["_Dpm"] = 0.0

        # RÈGLE STAFF RCA & MATRICE DES SPLITS :
        # Le DataFrame df_w contient strictement les participations validées dans la matrice des splits
        df_active = df_w.copy()

        for g in GROUPS:
            sub = df_active[df_active["_Groupe"] == g]
            if not sub.empty:
                split_means = sub.groupby(split_col)[["_Dist_m", "_Z3_m", "_VHSR_m", "_Acc_Dec", "_Dpm"]].mean()
                tot = split_means.sum()
                res[g] = {
                    "Dist_m": float(tot["_Dist_m"]),
                    "HSR_m": float(tot["_Z3_m"]),
                    "VHSR_m": float(tot["_VHSR_m"]),
                    "Acc_Dec": float(tot["_Acc_Dec"]),
                    "Dpm": float(split_means["_Dpm"].mean()) if not split_means["_Dpm"].empty else 0.0,
                }

    # Repli pour les groupes manquants ou si mode 'all'
    for g in GROUPS:
        if g not in res:
            sub = df_indiv[df_indiv["Groupe"] == g] if not df_indiv.empty and "Groupe" in df_indiv.columns else pd.DataFrame()
            if not sub.empty:
                res[g] = {
                    "Dist_m": float(sub["Distance (km)"].mean() * 1000.0) if "Distance (km)" in sub.columns else 0.0,
                    "HSR_m": float(sub[" Distance haute intensité (m)"].mean()) if " Distance haute intensité (m)" in sub.columns else 0.0,
                    "VHSR_m": float(sub["VHSR (m)"].mean()) if "VHSR (m)" in sub.columns else 0.0,
                    "Acc_Dec": float(sub["Accel/Dec"].mean()) if "Accel/Dec" in sub.columns else 0.0,
                    "Dpm": float(sub["Distance Per Min (m/min)"].mean()) if "Distance Per Min (m/min)" in sub.columns else 0.0,
                }
            else:
                res[g] = {"Dist_m": 0.0, "HSR_m": 0.0, "VHSR_m": 0.0, "Acc_Dec": 0.0, "Dpm": 0.0}

    return res


def calculer_metriques_jauges(
    df_subset: pd.DataFrame,
    group_name: str,
    mode_calc: str = "Moyenne / joueur",
    group_split_stats: Optional[Dict[str, float]] = None,
    config_objectifs: Optional[Dict[str, Dict[str, float]]] = None,
) -> Dict[str, Tuple[float, float]]:
    """Calcule dynamiquement les 4 métriques des jauges pour le groupe sélectionné."""
    if config_objectifs is None:
        config_objectifs = get_active_objectifs_config()
    targets = config_objectifs.get(group_name, config_objectifs.get("Tous les groupes", OBJECTIFS_CONFIG["Tous les groupes"]))

    if len(df_subset) == 0:
        return {k: (v, 0.0) for k, v in targets.items()}

    use_mean = (mode_calc == "Moyenne / joueur")

    if use_mean and group_split_stats is not None:
        val_dist = float(group_split_stats.get("Dist_m", 0.0))
        val_hsr = float(group_split_stats.get("HSR_m", 0.0))
        val_vhsr = float(group_split_stats.get("VHSR_m", 0.0))
        val_acc_dec = float(group_split_stats.get("Acc_Dec", 0.0))
    else:
        d_m = df_subset["Distance (km)"] * 1000.0
        val_dist = d_m.mean() if use_mean else d_m.sum()

        # HSR = Zone 3 (selon méthodologie STAPS RCA)
        if "Distance in Speed Zone 3  (km)" in df_subset.columns:
            hsr_m = df_subset["Distance in Speed Zone 3  (km)"] * 1000.0
        elif " Distance haute intensité (m)" in df_subset.columns:
            hsr_m = df_subset[" Distance haute intensité (m)"]
        else:
            hsr_m = pd.Series([0.0] * len(df_subset))
        val_hsr = hsr_m.mean() if use_mean else hsr_m.sum()

        # VHSR = Zone 4 + Zone 5
        if "VHSR (m)" in df_subset.columns:
            vhsr_m = df_subset["VHSR (m)"]
        elif "Distance in Speed Zone 4  (km)" in df_subset.columns and "Distance in Speed Zone 5  (km)" in df_subset.columns:
            vhsr_m = (df_subset["Distance in Speed Zone 4  (km)"] + df_subset["Distance in Speed Zone 5  (km)"]) * 1000.0
        else:
            vhsr_m = pd.Series([0.0] * len(df_subset))
        val_vhsr = vhsr_m.mean() if use_mean else vhsr_m.sum()

        acc_dec = df_subset["Accel/Dec"] if "Accel/Dec" in df_subset.columns else pd.Series([0.0] * len(df_subset))
        val_acc_dec = acc_dec.mean() if use_mean else acc_dec.sum()

    nb_joueurs = len(df_subset)
    multiplier = 1.0 if use_mean else max(float(nb_joueurs), 1.0)

    return {
        "Objectif Distance (m)": (targets["Objectif Distance (m)"] * multiplier, val_dist),
        "Obj. HSR": (targets["Obj. HSR"] * multiplier, val_hsr),
        "Obj. VHSR": (targets["Obj. VHSR"] * multiplier, val_vhsr),
        "Obj. ACC/DEC": (targets["Obj. ACC/DEC"] * multiplier, val_acc_dec),
    }


def generer_podium_dynamique(df_subset: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    """Génère le Podium des 3 tops à partir des données réelles avec photos Base64 et postes."""
    def _make_player_entry(rank: int, name: Any, stat_str: Any, poste: Any = "") -> Dict[str, Any]:
        p_name = str(name).strip() if pd.notna(name) and name is not None else "-"
        if p_name in ["nan", "None", ""]:
            p_name = "-"
        p_poste = str(poste).strip() if pd.notna(poste) and poste is not None else ""
        if p_poste in ["nan", "None"]:
            p_poste = ""
        if not p_poste and p_name and p_name != "-":
            p_poste = joueurs_data.get(p_name, "")
            if not p_poste:
                p_poste = JOUEURS_CUSTOM_CACHE.get(p_name, {}).get("poste", "")
            if not p_poste:
                parts = p_name.split()
                if len(parts) >= 2:
                    inv1 = f"{parts[-1]} {' '.join(parts[:-1])}"
                    p_poste = joueurs_data.get(inv1, "")
                    if not p_poste:
                        inv2 = f"{' '.join(parts[1:])} {parts[0]}"
                        p_poste = joueurs_data.get(inv2, "")
        return {
            "rank": rank,
            "name": p_name,
            "poste": p_poste if p_poste else "RCA",
            "stat": str(stat_str) if stat_str is not None else "-",
            "avatar_src": get_player_image_b64(p_name, size=80 if rank == 1 else 65) if p_name != "-" else get_player_image_b64("", size=80 if rank == 1 else 65),
        }

    def _placeholder(rank: int) -> Dict[str, Any]:
        return {
            "rank": rank,
            "name": "-",
            "poste": "",
            "stat": "-",
            "avatar_src": get_player_image_b64("", size=80 if rank == 1 else 65),
        }

    if df_subset is None or not isinstance(df_subset, pd.DataFrame) or len(df_subset) == 0:
        return {
            "Top speed (km/h)": {"icon": "⚡", "players": [_placeholder(1), _placeholder(2), _placeholder(3)]},
            "% de marche": {"icon": "🚶", "players": [_placeholder(1), _placeholder(2), _placeholder(3)]},
            "Haute intensité (m)": {"icon": "🔥", "players": [_placeholder(1), _placeholder(2), _placeholder(3)]},
        }

    podium_result: Dict[str, Dict[str, Any]] = {}

    # 1. TOP SPEED
    ts_col = "Top Speed (km/h)"
    if ts_col in df_subset.columns:
        s_ts = pd.to_numeric(df_subset[ts_col], errors="coerce").fillna(0.0)
        df_sorted_ts = df_subset.assign(_ts_sort=s_ts).sort_values(by="_ts_sort", ascending=False)
    else:
        df_sorted_ts = df_subset
    players_ts = []
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

    # 2. % DE MARCHE (Les 3 plus bas : ceux qui ont le moins marché, > 0%)
    walk_col = "%Marche "
    if walk_col in df_subset.columns:
        s_walk = pd.to_numeric(df_subset[walk_col], errors="coerce").fillna(0.0)
        df_valid_walk = df_subset[s_walk > 0.0]
        if not df_valid_walk.empty:
            df_sorted_walk = df_valid_walk.sort_values(by=walk_col, ascending=True)
        else:
            df_sorted_walk = df_subset.sort_values(by=walk_col, ascending=True)
    else:
        df_sorted_walk = df_subset
    players_walk = []
    for rank, (_, row) in enumerate(df_sorted_walk.head(3).iterrows(), start=1):
        name = str(row.get("Nom ; Prénom", "-")).strip()
        val_raw = row.get(walk_col, 0.0)
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

    # 3. HAUTE INTENSITÉ (m) - ALIGNEMENT SUR LA MÉTHODE STAFF RCA (Cumul Zones 3, 4 et 5 > 14.4 km/h)
    hi_col = "Haute intensité Podium (m)" if "Haute intensité Podium (m)" in df_subset.columns else " Distance haute intensité (m)"
    if hi_col in df_subset.columns:
        s_hi = pd.to_numeric(df_subset[hi_col], errors="coerce").fillna(0.0)
        df_sorted_hi = df_subset.assign(_hi_sort=s_hi).sort_values(by="_hi_sort", ascending=False)
    else:
        df_sorted_hi = df_subset
    players_hi = []
    for rank, (_, row) in enumerate(df_sorted_hi.head(3).iterrows(), start=1):
        name = str(row.get("Nom ; Prénom", "-")).strip()
        val_raw = row.get(hi_col, 0.0)
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
# 7. INJECTION DANS LE TEMPLATE EXCEL VIA OPENPYXL (data_only=False)
# ==============================================================================

def populate_template_excel(df_indiv: pd.DataFrame, df_raw: Optional[pd.DataFrame] = None) -> io.BytesIO:
    """Charge le classeur officiel (Modele_Vierge_STAPS.xlsx) et injecte les données."""
    template_path = get_template_path()
    if not template_path:
        raise FileNotFoundError(f"Le fichier modèle '{TEMPLATE_FILENAME}' est introuvable.")

    wb = openpyxl.load_workbook(template_path, data_only=False)

    if 'données indiv. joueurs' in wb.sheetnames:
        ws_indiv = wb['données indiv. joueurs']
        num_players = len(df_indiv)

        for i, row in df_indiv.iterrows():
            r = 3 + i
            for col_idx, col_name in enumerate(FINAL_COLUMNS[:22], start=1):
                val = row[col_name]
                # Formatage des pourcentages pour le format natif Excel "0%"
                if col_name in ['%Marche ', '%course', '%HSR ', '%VHSR'] and isinstance(val, (int, float)) and val > 1.0:
                    val = val / 100.0
                ws_indiv.cell(row=r, column=col_idx, value=val)

            # Formules et métriques de la préparatrice physique RCA
            ws_indiv.cell(row=r, column=4, value=f"=SUM(I{r}:K{r})*1000")
            ws_indiv.cell(row=r, column=18, value=f"=SUM(N{r}:Q{r})")
            ws_indiv.cell(row=r, column=23, value=f"=SUM(S{r}:V{r})")
            ws_indiv.cell(row=r, column=24, value=row['HSR m/min'])
            ws_indiv.cell(row=r, column=26, value=f"=I{r}*1000")
            ws_indiv.cell(row=r, column=27, value=f"=(SUM(J{r}:K{r}))*1000")

        last_injected_row = 3 + num_players - 1
        if ws_indiv.max_row > last_injected_row:
            for extra_r in range(last_injected_row + 1, ws_indiv.max_row + 1):
                for c in range(1, 35):
                    ws_indiv.cell(row=extra_r, column=c).value = None

    if df_raw is not None and 'données brutes' in wb.sheetnames:
        ws_brutes = wb['données brutes']
        for col_idx, col_name in enumerate(df_raw.columns, start=1):
            ws_brutes.cell(row=3, column=col_idx, value=col_name)

        for row_idx, row_vals in enumerate(df_raw.itertuples(index=False), start=4):
            for col_idx, cell_val in enumerate(row_vals, start=1):
                val_to_write = cell_val if not pd.isna(cell_val) else None
                ws_brutes.cell(row=row_idx, column=col_idx, value=val_to_write)

        last_brute_row = 4 + len(df_raw) - 1
        if ws_brutes.max_row > last_brute_row:
            for extra_r in range(last_brute_row + 1, ws_brutes.max_row + 1):
                for c in range(1, len(df_raw.columns) + 1):
                    ws_brutes.cell(row=extra_r, column=c).value = None

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def create_fallback_excel(df_indiv: pd.DataFrame) -> io.BytesIO:
    """Génération Excel de secours autonome si le modèle est absent."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_indiv.to_excel(writer, sheet_name="données indiv. joueurs", index=False)
        ws = writer.sheets["données indiv. joueurs"]
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "B2"
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        for col_idx in range(1, len(FINAL_COLUMNS) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            col_letter = get_column_letter(col_idx)
            if col_idx == 1:
                ws.column_dimensions[col_letter].width = 15
            else:
                ws.column_dimensions[col_letter].width = max(len(str(FINAL_COLUMNS[col_idx - 1]).strip()) + 3, 14)
    output.seek(0)
    return output




def calculate_acwr(df: pd.DataFrame, metric: str = "total_distance") -> pd.DataFrame:
    """Calcule le ratio charge aigue / chronique (ACWR) par joueur sur base glissante (7j / 28j)."""
    if df.empty or metric not in df.columns:
        return pd.DataFrame()

    df_calc = df.copy()
    date_col = "date" if "date" in df_calc.columns else ("Date" if "Date" in df_calc.columns else None)
    player_col = "player_name" if "player_name" in df_calc.columns else ("Nom" if "Nom" in df_calc.columns else "Player Name")
    
    if not date_col or not player_col:
        return pd.DataFrame()

    df_calc[date_col] = pd.to_datetime(df_calc[date_col], errors="coerce")
    df_calc = df_calc.dropna(subset=[date_col]).sort_values(by=[player_col, date_col])

    results = []
    for player, group in df_calc.groupby(player_col):
        daily = group.set_index(date_col).resample("D")[metric].sum().fillna(0)
        acute = daily.rolling(window=7, min_periods=1).mean()
        chronic = daily.rolling(window=28, min_periods=1).mean()
        acwr_val = acute / chronic.replace(0, np.nan)

        res_df = pd.DataFrame({
            player_col: player,
            date_col: daily.index,
            "daily_load": daily.values,
            "acute_load_7d": acute.values,
            "chronic_load_28d": chronic.values,
            "acwr": acwr_val.values.round(2)
        })
        results.append(res_df)

    if not results:
        return pd.DataFrame()

    return pd.concat(results, ignore_index=True)
