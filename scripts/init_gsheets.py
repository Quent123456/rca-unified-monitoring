"""
Script d'initialisation des 4 onglets Google Sheets.
Structure exacte d'apres les images modele.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tomllib
import gspread
from google.oauth2.service_account import Credentials
import time

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SPREADSHEET_NAME = "Test - Suivi Charge RCA Amiens"

# Auth
secrets_path = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
with open(secrets_path, "rb") as f:
    secrets = tomllib.load(f)
creds = Credentials.from_service_account_info(secrets["gcp_service_account"], scopes=SCOPES)
client = gspread.authorize(creds)
spreadsheet = client.open(SPREADSHEET_NAME)
print(f"Connecte a : {spreadsheet.title}")

def get_or_create(title, rows=500, cols=30):
    try:
        ws = spreadsheet.worksheet(title)
        print(f"  Onglet '{title}' existe deja.")
        return ws
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)
        print(f"  Onglet '{title}' cree.")
        time.sleep(1)
        return ws

# ── Renommer l'onglet existant en "Saisies" ───────────────────────────────────
print("\n[1] Configuration onglet Saisies...")
existing = spreadsheet.worksheets()
if len(existing) == 1 and existing[0].title != "Saisies":
    existing[0].update_title("Saisies")
    print("  Onglet existant renomme en 'Saisies'.")
    time.sleep(1)

ws_saisies = get_or_create("Saisies")

# En-tetes Saisies (inspire de Form_Responses1 dans les images)
headers_saisies = [
    "Cle de recherche",
    "Horodateur",
    "Difficulte seance (RPE texte)",
    "Ressenti / Notes",
    "Nom / Prenom",
    "Poste",
    "Numero Semaine",
    "Jour seance",
    "RPE (Chiffre)",
    "Duree (min)",
    "Charge (UA)",
    "Fatigue | Courbatures | Sommeil",
]
existing_h = ws_saisies.row_values(1)
if not existing_h or existing_h[0] != "Cle de recherche":
    ws_saisies.update("A1:L1", [headers_saisies])
    ws_saisies.format("A1:L1", {
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
        "backgroundColor": {"red": 0.29, "green": 0.29, "blue": 0.29},
    })
    print("  En-tetes Saisies ecrits.")
time.sleep(1)

# ── Onglet RPE INDV ────────────────────────────────────────────────────────────
print("\n[2] Configuration onglet RPE INDV...")
ws_indv = get_or_create("RPE INDV", rows=500, cols=26)

row1 = [
    "Poste", "NOM / Prenom", "Semaine N",
    "MARDI", "", "",
    "MERCREDI", "", "",
    "VENDREDI", "", "",
    "Jour de match", "", "",
    "Charge moyenne", "Total charge E.", "Ecart Type",
    "Indice Monotonie (IM)", "Indice de Contrainte (IC)",
    "Indice de Forme (IF)", "Variation charge",
    "Charge Aigue", "Charge Chronique",
    "Ratio C.A et C.C", "Total charge E + Match",
]
row2 = [
    "", "", "",
    "RPE", "DUREE (min)", "CHARGE (UA)",
    "RPE", "DUREE (min)", "CHARGE (UA)",
    "RPE", "DUREE (min)", "CHARGE (UA)",
    "RPE", "DUREE (min)", "CHARGE (UA)",
    "", "", "", "", "", "", "", "", "", "", "",
]

existing_h = ws_indv.row_values(1)
if not existing_h or existing_h[0] != "Poste":
    ws_indv.update("A1:Z2", [row1, row2])
    # Mise en forme en-tetes groupes
    ws_indv.format("A1:C2", {"textFormat": {"bold": True}})
    ws_indv.format("D1:F2", {"backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}, "textFormat": {"bold": True}})
    ws_indv.format("G1:I2", {"backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.9}, "textFormat": {"bold": True}})
    ws_indv.format("J1:L2", {"backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.8}, "textFormat": {"bold": True}})
    ws_indv.format("M1:O2", {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 0.0}, "textFormat": {"bold": True}})
    # Colonnes metriques
    ws_indv.format("P1:P2", {"backgroundColor": {"red": 0.94, "green": 0.53, "blue": 0.13}, "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}}})
    ws_indv.format("Q1:Q2", {"backgroundColor": {"red": 0.94, "green": 0.53, "blue": 0.13}, "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}}})
    ws_indv.format("R1:R2", {"backgroundColor": {"red": 0.75, "green": 0.75, "blue": 0.75}, "textFormat": {"bold": True}})
    ws_indv.format("S1:S2", {"backgroundColor": {"red": 0.0, "green": 0.85, "blue": 0.85}, "textFormat": {"bold": True}})
    ws_indv.format("T1:T2", {"backgroundColor": {"red": 0.9, "green": 0.1, "blue": 0.1}, "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}}})
    ws_indv.format("U1:U2", {"backgroundColor": {"red": 0.0, "green": 0.8, "blue": 0.0}, "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}}})
    ws_indv.format("V1:V2", {"backgroundColor": {"red": 1.0, "green": 0.6, "blue": 0.0}, "textFormat": {"bold": True}})
    ws_indv.format("W1:W2", {"backgroundColor": {"red": 0.7, "green": 0.0, "blue": 0.9}, "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}}})
    ws_indv.format("X1:X2", {"backgroundColor": {"red": 1.0, "green": 0.7, "blue": 0.8}, "textFormat": {"bold": True}})
    ws_indv.format("Y1:Y2", {"backgroundColor": {"red": 1.0, "green": 0.8, "blue": 0.9}, "textFormat": {"bold": True}})
    ws_indv.format("Z1:Z2", {"backgroundColor": {"red": 0.4, "green": 0.0, "blue": 0.0}, "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}}})
    print("  En-tetes RPE INDV ecrits avec couleurs.")
time.sleep(2)

# ── Onglet RPE PAR LIGNE SEMAINE ──────────────────────────────────────────────
print("\n[3] Configuration onglet RPE PAR LIGNE SEMAINE...")
ws_sem = get_or_create("RPE PAR LIGNE SEMAINE", rows=200, cols=13)

headers_sem_r1 = ["RPE PAR LIGNE / SEMAINE"] + [""] * 12
headers_sem_r2 = [
    "Semaine", "Poste",
    "Charge moyenne", "Total charge E.", "Ecart Type",
    "Indice Monotonie (IM)", "Indice de Contrainte (IC)",
    "Indice de Forme (IF)", "Variation charge",
    "Charge Aigue", "Charge Chronique",
    "Ratio C.A et C.C", "Total charge E + Match",
]
existing_h = ws_sem.row_values(1)
if not existing_h or existing_h[0] != "RPE PAR LIGNE / SEMAINE":
    ws_sem.update("A1:M2", [headers_sem_r1, headers_sem_r2])
    ws_sem.format("A1:M1", {
        "textFormat": {"bold": True, "fontSize": 14, "foregroundColor": {"red":0.8,"green":0.0,"blue":0.0}},
        "horizontalAlignment": "CENTER",
    })
    ws_sem.merge_cells("A1:M1")
    ws_sem.format("A2:M2", {
        "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}},
        "backgroundColor": {"red": 0.15, "green": 0.15, "blue": 0.15},
    })
    ws_sem.format("C2:E2", {"backgroundColor": {"red": 0.94, "green": 0.53, "blue": 0.13}})
    ws_sem.format("F2:F2", {"backgroundColor": {"red": 0.0, "green": 0.85, "blue": 0.85}})
    ws_sem.format("G2:G2", {"backgroundColor": {"red": 0.9, "green": 0.1, "blue": 0.1}})
    ws_sem.format("H2:H2", {"backgroundColor": {"red": 0.0, "green": 0.8, "blue": 0.0}})
    ws_sem.format("I2:I2", {"backgroundColor": {"red": 1.0, "green": 0.6, "blue": 0.0}})
    ws_sem.format("J2:J2", {"backgroundColor": {"red": 0.7, "green": 0.0, "blue": 0.9}})
    ws_sem.format("K2:K2", {"backgroundColor": {"red": 1.0, "green": 0.7, "blue": 0.8}})
    ws_sem.format("M2:M2", {"backgroundColor": {"red": 0.4, "green": 0.0, "blue": 0.0}})
    print("  En-tetes RPE PAR LIGNE SEMAINE ecrits.")
time.sleep(2)

# ── Onglet RPE PAR LIGNE MOIS ─────────────────────────────────────────────────
print("\n[4] Configuration onglet RPE PAR LIGNE MOIS...")
ws_mois = get_or_create("RPE PAR LIGNE MOIS", rows=200, cols=13)

headers_mois_r1 = ["RPE PAR LIGNE"] + [""] * 12
headers_mois_r2 = [
    "Mois", "Poste",
    "Charge moyenne", "Total charge E.", "Ecart Type",
    "Indice Monotonie (IM)", "Indice de Contrainte (IC)",
    "Indice de Forme (IF)", "Variation charge",
    "Charge Aigue", "Charge Chronique",
    "Ratio C.A et C.C", "Total charge E + Match",
]
existing_h = ws_mois.row_values(1)
if not existing_h or existing_h[0] != "RPE PAR LIGNE":
    ws_mois.update("A1:M2", [headers_mois_r1, headers_mois_r2])
    ws_mois.format("A1:M1", {
        "textFormat": {"bold": True, "fontSize": 14, "foregroundColor": {"red":0.8,"green":0.0,"blue":0.0}},
        "horizontalAlignment": "CENTER",
    })
    ws_mois.merge_cells("A1:M1")
    ws_mois.format("A2:M2", {
        "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}},
        "backgroundColor": {"red": 0.15, "green": 0.15, "blue": 0.15},
    })
    print("  En-tetes RPE PAR LIGNE MOIS ecrits.")
time.sleep(1)

print("\nInitialisation terminee. Onglets dans le fichier :")
for ws in spreadsheet.worksheets():
    print(f"  - {ws.title}")
