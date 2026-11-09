"""
pages/5_Rapports_PDF.py
Interface de sélection, génération et exportation des rapports PDF officiels RCA.
"""

from __future__ import annotations

import datetime
import io
import os
import pandas as pd
import streamlit as st

try:
    from modules import photo_manager, weekly_memory, rpe_connector, gps_processing, pdf_reports
except ImportError:
    import photo_manager
    import weekly_memory
    import rpe_connector
    import gps_processing
    import pdf_reports

st.set_page_config(
    page_title="Rapports PDF - RCA Monitoring",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Rapports PDF & Bilans Officiels RCA")
st.caption("Exportez en un clic les rapports de performance : Séance Coachs, Rapport Hebdomadaire ou Fiche Athlète.")

tab_coach, tab_weekly, tab_indiv = st.tabs([
    "🏉 Rapport Séance Coachs (Officiel Paysage)",
    "📅 Rapport Hebdomadaire Multi-Séances",
    "👤 Fiche Individuelle Athlète (A4)"
])

# ------------------------------------------------------------------------------
# 1. RAPPORT SÉANCE COACHS (FPDF PAYSAGE)
# ------------------------------------------------------------------------------
with tab_coach:
    st.subheader("🏉 Rapport Séance du Jour pour le Staff / Entraîneurs")
    st.markdown("Génère le document officiel (4 à 8 pages) avec jauges d'objectifs, podiums par ligne/poste, graphiques et métriques RPE.")

    df_indiv = st.session_state.get("df_indiv")
    if df_indiv is None or df_indiv.empty:
        st.info("💡 Aucune séance active dans la session. Veuillez d'abord traiter un fichier dans **4. Analyse GPS** ou charger une séance de test ci-dessous.")
        col_t1, col_t2 = st.columns([1, 3])
        with col_t1:
            if st.button("🧪 Charger séance démo pour le PDF", key="btn_load_demo_pdf", use_container_width=True):
                # Chargement automatique des données de démo Catapult
                sample_file = os.path.join("data", "sample_catapult_rugby.csv")
                if os.path.exists(sample_file):
                    df_sample = pd.read_csv(sample_file)
                    st.session_state["df_indiv"] = df_sample
                    st.session_state["date_seance"] = datetime.date(2026, 9, 8)
                    st.rerun()
                else:
                    st.warning("Fichier de démo introuvable dans data/.")
    else:
        st.success(f"Séance chargée : **{len(df_indiv)} athlètes** prêts pour le rapport PDF.")
        col_c1, col_c2, col_c3 = st.columns(3)
        with col_c1:
            date_s = st.date_input("Date de la séance", value=st.session_state.get("date_seance") or datetime.date.today(), key="pdf_date_input")
        with col_c2:
            adv = st.text_input("Adversaire / Contexte (optionnel)", placeholder="Ex: Entraînement intensif ou vs Beauvais", key="pdf_adv_input")
        with col_c3:
            incl_comp = st.checkbox("Inclure comparatif multi-séances", value=False, key="pdf_comp_input")

        synth_note = st.text_area("Synthèse du staff / Commentaires préparateur physique :", height=80, key="pdf_synth_input")

        if st.button("🚀 Générer le Rapport Séance Coachs (PDF)", type="primary", use_container_width=True, key="btn_gen_coach_pdf"):
            with st.spinner("Génération du document PDF officiel RCA en cours..."):
                try:
                    pdf_bytes = pdf_reports.generer_rapport_pdf_coachs(
                        df_indiv=df_indiv,
                        date_seance=date_s,
                        synthese_note=synth_note,
                        inclure_comparatif=incl_comp,
                        adversaire=adv
                    )
                    date_str = date_s.strftime('%Y%m%d') if hasattr(date_s, 'strftime') else 'seance'
                    st.download_button(
                        label="📥 Télécharger le Rapport Coachs (.pdf)",
                        data=pdf_bytes,
                        file_name=f"Rapport_Coachs_RCA_{date_str}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="btn_download_coach_pdf"
                    )
                    st.toast("Rapport PDF généré avec succès !", icon="🏉")
                except Exception as e:
                    st.error(f"Erreur lors de la génération du PDF : {e}")

# ------------------------------------------------------------------------------
# 2. RAPPORT HEBDOMADAIRE MULTI-SÉANCES (WEEKLY MEMORY)
# ------------------------------------------------------------------------------
with tab_weekly:
    st.subheader("📅 Rapport Hebdomadaire & Bilan Multi-Séances")
    st.caption("Consultez l'historique archivé par semaine ISO, comparez les charges (Mardi / Mercredi / Vendredi / Dimanche) et exportez le bilan semaine.")
    weekly_memory.render_weekly_reports_view()

# ------------------------------------------------------------------------------
# 3. FICHE INDIVIDUELLE ATHLÈTE A4 (REPORTLAB)
# ------------------------------------------------------------------------------
with tab_indiv:
    st.subheader("👤 Fiche Athlète Individuelle (A4)")
    st.caption("Synthèse individuelle éditée pour un joueur en combinant son profil, ses scores Hooper, ses charges sRPE et ses métriques GPS.")

    try:
        from modules.gsheet_sync import get_players, get_wellness_data
        players_df = get_players()
        wellness_df = get_wellness_data()
    except Exception:
        players_df = pd.DataFrame()
        wellness_df = pd.DataFrame()

    if not players_df.empty:
        col_p1, col_p2 = st.columns([2, 1])
        def format_nom_joueur(row):
            return f"{str(row.get('nom', '')).upper()} {str(row.get('prenom', '')).capitalize()}".strip()

        with col_p1:
            noms_liste = [format_nom_joueur(r) for _, r in players_df.iterrows()]
            joueur_choisi = st.selectbox("Sélectionner un athlète :", noms_liste, key="sel_indiv_player_pdf")
            idx = noms_liste.index(joueur_choisi)
            p_row = players_df.iloc[idx]
        with col_p2:
            nb_j = st.selectbox("Période :", [7, 14, 28], format_func=lambda x: f"Derniers {x} jours", key="sel_indiv_period_pdf")

        if st.button("📄 Exporter la fiche individuelle", type="primary", use_container_width=True, key="btn_gen_indiv_pdf"):
            wellness_sub = {
                "score_hooper_moy": 11.5,
                "sommeil_moy": 2.8,
                "fatigue_moy": 3.0,
                "courbatures_moy": 2.9,
                "stress_moy": 2.1,
                "srpe_total": 1420
            }
            gps_sub = {
                "total_distance_m": 12450,
                "hsr_distance_m": 840,
                "sprint_distance_m": 180,
                "vmax_kmh": 31.8,
                "intensity_m_min": 78.4,
                "player_load_total": 412,
                "acwr": 1.12
            }
            pdf_b = pdf_reports.generate_player_pdf_report(
                player_name=joueur_choisi,
                player_info=p_row.to_dict(),
                wellness_summary=wellness_sub,
                gps_summary=gps_sub,
                period_str=f"Derniers {nb_j} jours"
            )
            st.download_button(
                label=f"📥 Télécharger la fiche de {joueur_choisi}",
                data=pdf_b,
                file_name=f"Fiche_{joueur_choisi.replace(' ', '_')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="btn_download_indiv_pdf"
            )
    else:
        st.info("Effectif en cours de chargement.")
