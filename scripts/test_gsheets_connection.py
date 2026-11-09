"""
Script de test de connexion Google Sheets.
Vérifie : auth, ouverture du fichier, listage des onglets.
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tomllib
import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SPREADSHEET_NAME = "Test - Suivi Charge RCA Amiens"

# Lire les secrets depuis le fichier TOML
secrets_path = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
with open(secrets_path, "rb") as f:
    secrets = tomllib.load(f)

creds_dict = secrets["gcp_service_account"]
print(f"[1] Credentials chargees pour : {creds_dict['client_email']}")

# Authentification
creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
client = gspread.authorize(creds)
print("[2] Authentification Google OK")

# Ouverture du fichier
try:
    spreadsheet = client.open(SPREADSHEET_NAME)
    print(f"[3] Fichier ouvert : '{spreadsheet.title}'")
    print(f"    URL : {spreadsheet.url}")
    sheets = spreadsheet.worksheets()
    print(f"[4] Onglets existants ({len(sheets)}) :")
    for ws in sheets:
        print(f"    - {ws.title} ({ws.row_count} lignes x {ws.col_count} colonnes)")
except gspread.SpreadsheetNotFound:
    print(f"[ERREUR] Fichier '{SPREADSHEET_NAME}' introuvable.")
    print("  -> Verifiez que le fichier est partage avec :")
    print(f"     {creds_dict['client_email']}")
    sys.exit(1)

print("\nConnexion Google Sheets operationnelle !")
