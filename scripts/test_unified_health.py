"""
Script de test de santé et de diagnostic global pour le projet RCA Unified Monitoring.
"""

import sys
import os

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

def test_imports():
    print("=== TEST DES MODULES METIER ===")
    import modules.alerts as alerts
    import modules.calculations as calculations
    import modules.database as database
    import modules.gps_processing as gps_processing
    import modules.gsheets as gsheets
    import modules.gsheet_sync as gsheet_sync
    import modules.pdf_reports as pdf_reports
    import modules.photo_manager as photo_manager
    import modules.rpe_connector as rpe_connector
    import modules.visualizations as visualizations
    import modules.weekly_memory as weekly_memory
    import modules.github_sync as github_sync
    import modules.ui_theme as ui_theme
    print("[OK] Tous les 13 modules de modules/ ont ete importes avec succes.")

    print("\n=== TEST DE RECHERCHE ET DATA ===")
    assert os.path.exists("assets/logo_rca.png"), "assets/logo_rca.png manquant"
    assert os.path.exists("logo_rca.png"), "logo_rca.png manquant"
    assert len(ui_theme.get_rca_logo_base64()) > 1000, "Base64 logo invalide ou vide"
    assert os.path.exists("data/Modele_Vierge_STAPS.xlsx"), "Modele Excel manquant"
    assert os.path.exists("data/weekly_memory/week_2026_W43.json"), "Archive W43 manquante"
    assert os.path.exists("data/weekly_memory/week_2026_W45.json"), "Archive W45 manquante"
    print("[OK] Tous les assets et fichiers de donnees sont presents.")

    print("\n=== TEST DES 5 PAGES STREAMLIT ===")
    pages = [
        "pages/1_Saisie_Quotidienne.py",
        "pages/2_Tableau_Individuel.py",
        "pages/3_Tableau_Collectif.py",
        "pages/4_Analyse_GPS.py",
        "pages/5_Gestion_Joueurs.py"
    ]
    for p in pages:
        assert os.path.exists(p), f"Page manquante: {p}"
        print(f"[OK] Page presente: {p}")

    print("\nTOUS LES TESTS DE DIAGNOSTIC ONT REUSSI SANS ERREUR !")

if __name__ == "__main__":
    test_imports()
