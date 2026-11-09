"""
Test end-to-end : ecriture d'une ligne de test dans Saisies.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tomllib
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SPREADSHEET_NAME = "Test - Suivi Charge RCA Amiens"

secrets_path = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
with open(secrets_path, "rb") as f:
    secrets = tomllib.load(f)

creds = Credentials.from_service_account_info(secrets["gcp_service_account"], scopes=SCOPES)
client = gspread.authorize(creds)
spreadsheet = client.open(SPREADSHEET_NAME)

ws = spreadsheet.worksheet("Saisies")
print(f"Onglet Saisies : {ws.row_count} lignes x {ws.col_count} colonnes")

# Verifier que les en-tetes sont la
headers = ws.row_values(1)
print(f"En-tetes (ligne 1) : {headers}")

# Ajouter une ligne de test
test_row = [
    "TEST_CONNEXION",
    datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    "5 : Effort moyen",
    "Test automatique connexion",
    "DUPONT Antoine",
    "Demi de melee",
    37,
    "Mercredi",
    5,
    90,
    450,
    "Fatigue:3 | Courbatures:3 | Sommeil:4",
]
ws.append_row(test_row, value_input_option="USER_ENTERED")

# Verifier que la ligne a ete ajoutee
all_rows = ws.get_all_values()
print(f"\nNombre de lignes apres ecriture : {len(all_rows)}")
print(f"Derniere ligne : {all_rows[-1]}")
print("\nTest d'ecriture Google Sheets : OK")
print(f"URL du fichier : {spreadsheet.url}")
