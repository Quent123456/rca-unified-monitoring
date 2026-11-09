"""
modules/gsheets.py
Intégration Google Sheets via gspread + Service Account.
Gère les 4 onglets du fichier "Test - Suivi Charge RCA Amiens".
"""

import gspread
from google.oauth2.service_account import Credentials
import streamlit as st
import pandas as pd
from datetime import datetime
from typing import Optional
import time

# ── Scopes Google API ──────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_NAME = "Test - Suivi Charge RCA Amiens"

# ── Noms des onglets ────────────────────────────────────────────────────────────
SHEET_SAISIES         = "Saisies"
SHEET_RPE_INDV        = "RPE INDV"
SHEET_RPE_LIGNE_SEM   = "RPE PAR LIGNE SEMAINE"
SHEET_RPE_LIGNE_MOIS  = "RPE PAR LIGNE MOIS"

# ── Mapping RPE chiffre → texte (Borg CR-10) ──────────────────────────────────
RPE_LABELS = {
    0:  "0 : Repos",
    1:  "1 : Très très facile",
    2:  "2 : Très facile",
    3:  "3 : Facile",
    4:  "4 : Effort modéré",
    5:  "5 : Effort moyen",
    6:  "6 : Effort un peu difficile",
    7:  "7 : Difficile",
    8:  "8 : Très difficile",
    9:  "9 : Très très difficile",
    10: "10 : Maximal",
}

# ── Postes rugby → catégories pour RPE PAR LIGNE ──────────────────────────────
POSTE_TO_LIGNE = {
    "Pilier gauche":         "1ER",
    "Pilier droit":          "1ER",
    "Talonneur":             "TAL.",
    "2ème ligne gauche":     "2EME",
    "2ème ligne droit":      "2EME",
    "3ème ligne aile gauche":"3ME",
    "3ème ligne aile droit": "3ME",
    "3ème ligne centre":     "3ME",
    "3ème ligne aile":       "3ME",
    "Demi de mêlée":         "DEMIS",
    "Ouverture":             "DEMIS",
    "Centre gauche":         "CENTRES",
    "Centre droit":          "CENTRES",
    "Centre":                "CENTRES",
    "Ailier gauche":         "AIL.",
    "Ailier droit":          "AIL.",
    "Ailier":                "AIL.",
    "Arrière":               "ARR.",
}

# ── Jours d'entraînement ───────────────────────────────────────────────────────
JOURS_ENTRAINEMENT = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi",
                       "Samedi", "Dimanche", "Jour de match"]

# Colonnes RPE INDV par jour (D=Mardi, G=Mercredi, J=Vendredi, M=Match)
# Chaque jour occupe 3 colonnes : RPE | DUREE | CHARGE
JOUR_COLUMN_MAP = {
    "Mardi":         {"rpe": 3,  "duree": 4,  "charge": 5},   # D, E, F (0-indexed)
    "Mercredi":      {"rpe": 6,  "duree": 7,  "charge": 8},   # G, H, I
    "Vendredi":      {"rpe": 9,  "duree": 10, "charge": 11},  # J, K, L
    "Jour de match": {"rpe": 12, "duree": 13, "charge": 14},  # M, N, O
}


# ── Authentification ───────────────────────────────────────────────────────────

@st.cache_resource(ttl=3600)
def get_gspread_client() -> Optional[gspread.Client]:
    """
    Retourne un client gspread authentifié via Service Account.
    Les credentials sont lus depuis st.secrets["gcp_service_account"].
    Retourne None si les secrets ne sont pas configurés.
    """
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        client = gspread.authorize(creds)
        return client
    except KeyError:
        return None
    except Exception as e:
        st.error(f"Erreur d'authentification Google : {e}")
        return None


def get_spreadsheet() -> Optional[gspread.Spreadsheet]:
    """Ouvre le fichier Google Sheets cible."""
    client = get_gspread_client()
    if client is None:
        return None
    try:
        return client.open(SPREADSHEET_NAME)
    except gspread.SpreadsheetNotFound:
        st.error(
            f"Fichier '{SPREADSHEET_NAME}' introuvable. "
            "Vérifiez qu'il est partagé avec le Service Account."
        )
        return None
    except Exception as e:
        st.error(f"Impossible d'ouvrir le fichier : {e}")
        return None


# ── Initialisation des onglets ─────────────────────────────────────────────────

def _get_or_create_sheet(spreadsheet: gspread.Spreadsheet,
                          title: str) -> gspread.Worksheet:
    """Retourne l'onglet s'il existe, le crée sinon."""
    try:
        return spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=title, rows=1000, cols=30)


def init_sheets() -> bool:
    """
    Crée les 4 onglets avec les en-têtes exacts si absents.
    Retourne True si succès, False sinon.
    """
    spreadsheet = get_spreadsheet()
    if spreadsheet is None:
        return False

    _init_saisies(spreadsheet)
    _init_rpe_indv(spreadsheet)
    _init_rpe_ligne_sem(spreadsheet)
    _init_rpe_ligne_mois(spreadsheet)
    return True


def _init_saisies(spreadsheet: gspread.Spreadsheet) -> None:
    """Initialise l'onglet Saisies avec les en-têtes de Form_Responses."""
    ws = _get_or_create_sheet(spreadsheet, SHEET_SAISIES)
    headers = ws.row_values(1)
    if not headers:
        ws.update(
            "A1:L1",
            [[
                "Clé de recherche",
                "Horodateur",
                "Difficulté séance (RPE texte)",
                "Ressenti / Notes",
                "Nom / Prénom",
                "Poste",
                "Numéro Semaine",
                "Jour semaine",
                "RPE (Chiffre)",
                "Durée (min)",
                "Charge (UA)",
                "Fatigue | Courbatures | Sommeil",
            ]],
        )
        # Formatage en-têtes (gras)
        ws.format("A1:L1", {"textFormat": {"bold": True}})


def _init_rpe_indv(spreadsheet: gspread.Spreadsheet) -> None:
    """Initialise l'onglet RPE INDV avec le double en-tête."""
    ws = _get_or_create_sheet(spreadsheet, SHEET_RPE_INDV)
    if ws.row_values(1):
        return  # Déjà initialisé

    # Ligne 1 : regroupements
    row1 = [
        "Poste", "NOM / Prénom", "Semaine N°",
        "MARDI", "", "",
        "MERCREDI", "", "",
        "VENDREDI", "", "",
        "Jour de match", "", "",
        "Charge moyenne", "Total charge E.", "Ecart Type",
        "Indice Monotonie (IM)", "Indice de Contrainte (IC)",
        "Indice de Forme (IF)", "Variation charge",
        "Charge Aigue (= Ch. moy.)", "Charge Chronique",
        "Ratio C.A et C.C", "Total charge E + Match",
    ]
    # Ligne 2 : sous-en-têtes
    row2 = [
        "", "", "",
        "RPE", "DUREE (min)", "CHARGE (UA)",
        "RPE", "DUREE (min)", "CHARGE (UA)",
        "RPE", "DUREE (min)", "CHARGE (UA)",
        "RPE", "DUREE (min)", "CHARGE (UA)",
        "", "", "", "", "", "", "", "", "", "", "",
    ]
    ws.update("A1:Z2", [row1, row2])
    ws.format("A1:Z2", {"textFormat": {"bold": True}})


def _init_rpe_ligne_sem(spreadsheet: gspread.Spreadsheet) -> None:
    """Initialise l'onglet RPE PAR LIGNE SEMAINE."""
    ws = _get_or_create_sheet(spreadsheet, SHEET_RPE_LIGNE_SEM)
    if ws.row_values(1):
        return

    row1 = ["RPE PAR LIGNE / SEMAINE"] + [""] * 10
    row2 = [
        "Semaine", "Poste",
        "Charge moyenne", "Total charge E.", "Ecart Type",
        "Indice Monotonie (IM)", "Indice de Contrainte (IC)",
        "Indice de Forme (IF)", "Variation charge",
        "Charge Aigue", "Charge Chronique",
        "Ratio C.A et C.C", "Total charge E + Match",
    ]
    ws.update("A1:M2", [row1, row2])
    ws.format("A1:M1", {
        "textFormat": {"bold": True, "fontSize": 14},
        "horizontalAlignment": "CENTER",
    })
    ws.format("A2:M2", {"textFormat": {"bold": True}})


def _init_rpe_ligne_mois(spreadsheet: gspread.Spreadsheet) -> None:
    """Initialise l'onglet RPE PAR LIGNE MOIS."""
    ws = _get_or_create_sheet(spreadsheet, SHEET_RPE_LIGNE_MOIS)
    if ws.row_values(1):
        return

    row1 = ["RPE PAR LIGNE / MOIS"] + [""] * 10
    row2 = [
        "Mois", "Poste",
        "Charge moyenne", "Total charge E.", "Ecart Type",
        "Indice Monotonie (IM)", "Indice de Contrainte (IC)",
        "Indice de Forme (IF)", "Variation charge",
        "Charge Aigue", "Charge Chronique",
        "Ratio C.A et C.C", "Total charge E + Match",
    ]
    ws.update("A1:M2", [row1, row2])
    ws.format("A1:M1", {
        "textFormat": {"bold": True, "fontSize": 14},
        "horizontalAlignment": "CENTER",
    })
    ws.format("A2:M2", {"textFormat": {"bold": True}})


# ── ÉCRITURE — Onglet Saisies ──────────────────────────────────────────────────

def append_saisie_row(
    nom_prenom: str,
    poste: str,
    semaine: int,
    jour: str,
    rpe: float,
    duree: float,
    charge: float,
    fatigue: int,
    courbatures: int,
    sommeil: int,
    notes: str = "",
    is_rest: bool = False,
) -> bool:
    """
    Ajoute une ligne dans l'onglet 'Saisies' lors de la validation du formulaire.
    Retourne True si succès, False sinon.

    Structure de la ligne (colonnes A→L) :
    A: Clé de recherche = semaine + nom (ex: "42CARPENTIER Maël")
    B: Horodateur
    C: Difficulté texte (ex: "4 : Effort modéré")
    D: Ressenti / Notes
    E: Nom / Prénom
    F: Poste
    G: Numéro Semaine
    H: Jour séance
    I: RPE (chiffre)
    J: Durée (min)
    K: Charge (UA)
    L: Fatigue | Courbatures | Sommeil
    """
    spreadsheet = get_spreadsheet()
    if spreadsheet is None:
        return False

    try:
        ws = spreadsheet.worksheet(SHEET_SAISIES)
    except gspread.WorksheetNotFound:
        init_sheets()
        ws = spreadsheet.worksheet(SHEET_SAISIES)

    horodateur = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    cle        = f"{semaine}{nom_prenom}"
    rpe_texte  = RPE_LABELS.get(int(rpe), str(rpe)) if not is_rest else "Repos"
    bien_etre  = f"Fatigue:{fatigue} | Courbatures:{courbatures} | Sommeil:{sommeil}"

    row = [
        cle,
        horodateur,
        rpe_texte,
        notes,
        nom_prenom,
        poste,
        semaine,
        jour,
        rpe if not is_rest else 0,
        duree if not is_rest else 0,
        charge if not is_rest else 0,
        bien_etre,
    ]

    try:
        ws.append_row(row, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        st.error(f"Erreur écriture Saisies : {e}")
        return False


# ── ÉCRITURE — Onglet RPE INDV ─────────────────────────────────────────────────

def update_rpe_indv(
    nom_prenom: str,
    poste: str,
    semaine: int,
    jour: str,
    rpe: float,
    duree: float,
    charge: float,
) -> bool:
    """
    Met à jour la cellule correspondante dans l'onglet RPE INDV.
    Trouve la ligne du joueur pour la semaine donnée, puis remplit
    la colonne du jour (Mardi / Mercredi / Vendredi / Jour de match).
    Crée un nouveau bloc joueur si inexistant.
    Retourne True si succès.
    """
    spreadsheet = get_spreadsheet()
    if spreadsheet is None:
        return False

    try:
        ws = spreadsheet.worksheet(SHEET_RPE_INDV)
    except gspread.WorksheetNotFound:
        init_sheets()
        ws = spreadsheet.worksheet(SHEET_RPE_INDV)

    if jour not in JOUR_COLUMN_MAP:
        return True  # Jour non mappé → on n'écrit pas (mais on n'échoue pas)

    col_map   = JOUR_COLUMN_MAP[jour]  # {rpe: col_idx, duree: col_idx, charge: col_idx}
    all_vals  = ws.get_all_values()
    target_row_idx = None  # Index 0-based dans all_vals

    # Recherche ligne = joueur + semaine correspondante
    for i, row in enumerate(all_vals):
        if len(row) >= 3:
            row_nom     = str(row[1]).strip()
            row_semaine = str(row[2]).strip()
            if row_nom == nom_prenom and row_semaine == str(semaine):
                target_row_idx = i
                break

    if target_row_idx is None:
        # Créer un nouveau bloc pour ce joueur
        target_row_idx = _create_player_block(ws, all_vals, nom_prenom, poste, semaine)

    if target_row_idx is None:
        return False

    sheet_row = target_row_idx + 1  # gspread est 1-based

    try:
        # Colonnes gspread (1-based) : A=1, B=2, C=3, D=4 ...
        ws.update_cell(sheet_row, col_map["rpe"]    + 1, rpe)
        ws.update_cell(sheet_row, col_map["duree"]  + 1, duree)
        ws.update_cell(sheet_row, col_map["charge"] + 1, charge)

        # Recalcul des métriques dans cette ligne
        _update_metrics_row(ws, sheet_row, all_vals, nom_prenom, semaine)
        return True
    except Exception as e:
        print(f"Erreur écriture RPE INDV : {e}")
        return False


def _create_player_block(
    ws: gspread.Worksheet,
    all_vals: list,
    nom_prenom: str,
    poste: str,
    semaine: int,
) -> Optional[int]:
    """
    Trouve ou crée la ligne de semaine pour le joueur.
    Retourne l'index 0-based de la ligne cible.
    """
    try:
        # Chercher si le joueur existe déjà
        existing_rows = []
        for i, row in enumerate(all_vals):
            if len(row) >= 2 and str(row[1]).strip() == nom_prenom:
                existing_rows.append(i)
                if len(row) >= 3 and str(row[2]).strip() == str(semaine):
                    return i

        if existing_rows:
            # Joueur existe → insérer la semaine manquante
            last_idx   = max(existing_rows)
            insert_pos = last_idx + 2  # gspread 1-based
            ws.insert_row(["", nom_prenom, semaine] + [""] * 23, insert_pos)
            return insert_pos - 1
        else:
            # Nouveau joueur → ajouter la ligne pour cette semaine
            new_row = [poste, nom_prenom, semaine] + [""] * 23
            ws.append_row(new_row, value_input_option="USER_ENTERED")
            return len(all_vals)
    except Exception as e:
        print(f"Erreur _create_player_block : {e}")
        return None


def _update_metrics_row(
    ws: gspread.Worksheet,
    sheet_row: int,
    all_vals: list,
    nom_prenom: str,
    semaine: int,
) -> None:
    """
    Recalcule les métriques (colonnes P→Z) pour la ligne mise à jour.
    Toutes les charges du joueur pour les semaines disponibles sont utilisées.
    Les calculs sont écrits comme valeurs numériques (pas de formules Sheet).
    """
    import numpy as np

    # Récupérer toutes les charges (F, I, L, O) de toutes les semaines du joueur
    charges_par_semaine = {}
    for i, row in enumerate(all_vals):
        if len(row) >= 2 and str(row[1]).strip() == nom_prenom:
            sem = str(row[2]).strip()
            try:
                s = int(sem)
                # Charges : col F=5, I=8, L=11, O=14 (0-indexed)
                c_vals = []
                for ci in [5, 8, 11, 14]:
                    v = row[ci] if ci < len(row) else ""
                    try:
                        c_vals.append(float(v))
                    except (ValueError, TypeError):
                        c_vals.append(0.0)
                charges_par_semaine[s] = c_vals
            except ValueError:
                pass

    # Récupérer la ligne courante fraîchement
    current_row = ws.row_values(sheet_row)
    while len(current_row) < 15:
        current_row.append("")

    charges_ligne = []
    for ci in [5, 8, 11, 14]:
        try:
            charges_ligne.append(float(current_row[ci]))
        except (ValueError, TypeError, IndexError):
            charges_ligne.append(0.0)

    charges_jour_entrainement = charges_ligne[:3]  # Mardi, Mercredi, Vendredi
    charge_match = charges_ligne[3]

    total_e   = sum(charges_jour_entrainement)
    moy       = total_e / 3 if total_e > 0 else 0
    ecart     = float(np.std(charges_jour_entrainement)) if total_e > 0 else 0
    im        = round(moy / ecart, 2) if ecart > 0 else 0
    ic        = round(total_e * im, 1)
    total_ematch = total_e + charge_match

    # Charge aiguë (semaine courante) et charge chronique (4 dernières semaines)
    charge_aigue = moy
    toutes_semaines = sorted(charges_par_semaine.keys())
    charges_chron_list = []
    for s in toutes_semaines:
        c = charges_par_semaine[s]
        charges_chron_list.append(sum(c[:3]) / 3 if sum(c[:3]) > 0 else 0)

    charge_chronique = round(float(np.mean(charges_chron_list[-4:])), 1) if charges_chron_list else 0
    ratio_ca_cc = round(charge_aigue / charge_chronique, 2) if charge_chronique > 0 else ""

    # Indice de Forme = chronique - aigue (variation tendance)
    if_ = round(charge_chronique - charge_aigue, 1)

    # Variation charge vs semaine précédente
    sem_prec = semaine - 1
    moy_prec = 0
    if sem_prec in charges_par_semaine:
        c_prec = charges_par_semaine[sem_prec]
        moy_prec = sum(c_prec[:3]) / 3 if sum(c_prec[:3]) > 0 else 0
    variation = round((moy - moy_prec) / moy_prec * 100, 2) if moy_prec > 0 else ""

    # Écriture colonnes P(16)→Z(26) [1-based]
    metrics = [
        round(moy, 1),          # P : Charge moyenne
        round(total_e, 1),      # Q : Total charge E.
        round(ecart, 1),        # R : Ecart Type
        im,                     # S : Indice Monotonie
        round(ic, 1),           # T : Indice Contrainte
        if_,                    # U : Indice de Forme
        variation,              # V : Variation charge
        round(charge_aigue, 1), # W : Charge Aiguë
        charge_chronique,       # X : Charge Chronique
        ratio_ca_cc,            # Y : Ratio C.A et C.C
        round(total_ematch, 1), # Z : Total charge E + Match
    ]
    # Mise à jour colonnes P→Z (indices 16→26)
    for offset, val in enumerate(metrics):
        ws.update_cell(sheet_row, 16 + offset, val)


# ── ÉCRITURE — Onglet RPE PAR LIGNE SEMAINE ───────────────────────────────────

def update_rpe_par_ligne_semaine(
    poste: str,
    semaine: int,
    nom_prenom: str,
    rpe: float,
    duree: float,
    charge: float,
) -> bool:
    """
    Met à jour les statistiques agrégées par poste et semaine.
    Lit toutes les données RPE INDV pour recalculer les moyennes du poste.
    """
    spreadsheet = get_spreadsheet()
    if spreadsheet is None:
        return False

    ligne_poste = POSTE_TO_LIGNE.get(poste)
    if not ligne_poste:
        return True  # Poste non reconnu, skip silencieux

    try:
        ws_indv = spreadsheet.worksheet(SHEET_RPE_INDV)
        ws_sem  = spreadsheet.worksheet(SHEET_RPE_LIGNE_SEM)
    except gspread.WorksheetNotFound:
        return False

    # Récupérer toutes les données RPE INDV pour ce poste et cette semaine
    all_indv = ws_indv.get_all_values()
    charges_poste = []

    for row in all_indv[2:]:  # Skip 2 lignes d'en-tête
        if len(row) < 15:
            continue
        row_poste   = str(row[0]).strip()
        row_nom     = str(row[1]).strip()
        row_semaine = str(row[2]).strip()

        # Chercher le poste dans la première colonne du bloc joueur
        if str(row_semaine) != str(semaine):
            continue

        # Vérifier si ce joueur appartient au bon groupe de ligne
        joueur_ligne = None
        # On doit retrouver le poste du joueur (peut être vide sur les lignes 2-5 du bloc)
        # On cherche dans les lignes précédentes
        for check_row in all_indv[2:]:
            if len(check_row) >= 2 and str(check_row[1]).strip() == row_nom:
                if str(check_row[0]).strip():
                    joueur_ligne = POSTE_TO_LIGNE.get(str(check_row[0]).strip())
                    break

        if joueur_ligne != ligne_poste:
            continue

        # Charges journalières
        try:
            c_vals = [float(row[ci]) if row[ci] else 0 for ci in [5, 8, 11]]
            total = sum(c_vals)
            moy   = total / 3 if total > 0 else 0
            charges_poste.append(moy)
        except (ValueError, IndexError):
            pass

    if not charges_poste:
        return True

    import numpy as np
    moy_poste   = round(float(np.mean(charges_poste)), 1)
    total_poste = round(float(np.sum(charges_poste)), 1)
    ecart_poste = round(float(np.std(charges_poste)), 1)
    im = round(moy_poste / ecart_poste, 2) if ecart_poste > 0 else 0
    ic = round(total_poste * im, 1)

    # Trouver ou créer la ligne dans RPE PAR LIGNE SEMAINE
    all_sem = ws_sem.get_all_values()
    target  = None
    for i, row in enumerate(all_sem):
        if len(row) >= 2 and str(row[0]) == str(semaine) and str(row[1]) == ligne_poste:
            target = i + 1  # 1-based
            break

    if target is None:
        # Créer la ligne
        ws_sem.append_row(
            [semaine, ligne_poste, moy_poste, total_poste, ecart_poste, im, ic, "", "", "", "", "", ""],
            value_input_option="USER_ENTERED",
        )
    else:
        ws_sem.update(
            f"A{target}:M{target}",
            [[semaine, ligne_poste, moy_poste, total_poste, ecart_poste, im, ic, "", "", "", "", "", ""]],
        )

    return True


# ── LECTURE — Données depuis Google Sheets ────────────────────────────────────

def read_saisies_as_df() -> pd.DataFrame:
    """Lit l'onglet Saisies et retourne un DataFrame."""
    spreadsheet = get_spreadsheet()
    if spreadsheet is None:
        return pd.DataFrame()
    try:
        ws  = spreadsheet.worksheet(SHEET_SAISIES)
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except Exception:
        return pd.DataFrame()


def is_configured() -> bool:
    """Vérifie si les secrets Google sont configurés."""
    try:
        _ = st.secrets["gcp_service_account"]
        return True
    except (KeyError, FileNotFoundError):
        return False
