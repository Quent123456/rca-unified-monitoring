"""
pages/4_Gestion_Joueurs.py
Gestion de l'effectif : ajout, modification et suppression (en cascade) de joueurs sur Google Sheets.
"""

import streamlit as st
import pandas as pd
import uuid
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Gestion Joueurs", page_icon="⚙️", layout="wide")

st.title("⚙️ Gestion de l'effectif")

# ── INITIALISATION CONNEXION GOOGLE SHEETS ───────────────────────────────────────
conn = st.connection("gsheets", type=GSheetsConnection)

# Fonction pour télécharger la liste à jour
def load_players():
    try:
        df = conn.read(worksheet="Effectif")
        df = df.dropna(how="all")
        df['id'] = df['id'].astype(str)
        return df
    except Exception:
        return pd.DataFrame(columns=["id", "nom", "prenom", "poste", "numero", "groupe"])

players_df = load_players()

POSTES_RUGBY = [
    "Pilier", "Talonneur", "2ème ligne", "3ème ligne",
    "Demi de mêlée", "Demi d'ouverture", "Centre", "Ailier", "Arrière",
]

GROUPES = ["Groupe Élargi", "Équipe A", "Équipe B"]

# ── LISTE DES JOUEURS ────────────────────────────────────────────────────────────
st.subheader("📋 Effectif actuel")

if players_df.empty:
    st.info("Aucun joueur enregistré dans la base de données Google Sheets.")
else:
    st.dataframe(
        players_df[["prenom", "nom", "poste", "groupe"]],
        use_container_width=True,
        hide_index=True
    )
st.markdown(f"**Total : {len(players_df)} joueur(s)**")
st.markdown("---")

# ── MODIFICATION / SUPPRESSION D'UN JOUEUR ───────────────────────────────────────
st.subheader("✏️ Modifier / Supprimer un joueur")
if not players_df.empty:
    player_options = {
        f"{row['prenom']} {row['nom']} ({row['poste']})": row
        for _, row in players_df.iterrows()
    }
    selected_label = st.selectbox("Sélectionner un joueur", options=list(player_options.keys()))
    selected_player = player_options[selected_label]

    with st.form("edit_player_form"):
        col1, col2 = st.columns(2)
        with col1:
            edit_prenom = st.text_input("Prénom", value=selected_player["prenom"])
            
            default_poste_idx = 0
            if selected_player["poste"] in POSTES_RUGBY:
                default_poste_idx = POSTES_RUGBY.index(selected_player["poste"])
            edit_poste = st.selectbox("Poste", POSTES_RUGBY, index=default_poste_idx)

        with col2:
            edit_nom = st.text_input("Nom", value=selected_player["nom"])
            
            default_groupe_idx = 0
            if selected_player.get("groupe") in GROUPES:
                default_groupe_idx = GROUPES.index(selected_player["groupe"])
            edit_groupe = st.selectbox("Groupe d'appartenance", GROUPES, index=default_groupe_idx)

        col_save, col_del = st.columns([1, 1])
        with col_save:
            btn_save = st.form_submit_button("💾 Enregistrer les modifications", type="primary")
        with col_del:
            btn_del = st.form_submit_button("🗑️ Supprimer ce joueur")

        if btn_save:
            # Récupérer l'ancien nom complet pour mettre à jour les historiques
            ancien_nom_complet = f"{selected_player['nom'].upper()} {selected_player['prenom'].capitalize()}"
            nouveau_nom_complet = f"{edit_nom.strip().upper()} {edit_prenom.strip().capitalize()}"

            # 1. Mise à jour de l'Effectif
            idx = players_df.index[players_df['id'] == selected_player['id']].tolist()[0]
            players_df.at[idx, 'nom'] = edit_nom.strip().upper()
            players_df.at[idx, 'prenom'] = edit_prenom.strip().capitalize()
            players_df.at[idx, 'poste'] = edit_poste
            players_df.at[idx, 'groupe'] = edit_groupe
            
            conn.update(worksheet="Effectif", data=players_df)
            
            # 2. Mise à jour de l'onglet Saisies (si le nom ou le poste a changé)
            if ancien_nom_complet != nouveau_nom_complet or selected_player['poste'] != edit_poste:
                try:
                    saisies_df = conn.read(worksheet="Saisies", ttl=0)
                    if not saisies_df.empty:
                        # Remplacer le vieux nom par le nouveau dans tout l'historique
                        mask = saisies_df['Nom / Prenom'] == ancien_nom_complet
                        saisies_df.loc[mask, 'Nom / Prenom'] = nouveau_nom_complet
                        saisies_df.loc[mask, 'Poste'] = edit_poste
                        
                        # Mettre à jour la clé de recherche qui contenait l'ancien nom
                        saisies_df.loc[mask, 'Cle de recherche'] = saisies_df.loc[mask, 'Numero Semaine'].astype(str) + nouveau_nom_complet.replace(' ', '')
                        
                        conn.update(worksheet="Saisies", data=saisies_df)
                except Exception as e:
                    st.error(f"Le profil est mis à jour, mais l'historique n'a pas pu être renommé : {e}")

            st.cache_data.clear()
            st.success(f"✅ Profil de {nouveau_nom_complet} mis à jour (historique inclus).")
            st.rerun()

        if btn_del:
            nom_complet_a_supprimer = f"{selected_player['nom'].upper()} {selected_player['prenom'].capitalize()}"

            with st.spinner("Suppression du joueur et de son historique en cours..."):
                # 1. Suppression du joueur dans l'onglet Effectif
                updated_players_df = players_df[players_df['id'] != selected_player['id']]
                conn.update(worksheet="Effectif", data=updated_players_df)
                
                # 2. Nettoyage en cascade dans l'onglet Saisies
                try:
                    saisies_df = conn.read(worksheet="Saisies", ttl=0)
                    if not saisies_df.empty:
                        # Ne garder que les lignes qui NE SONT PAS le joueur à supprimer
                        updated_saisies_df = saisies_df[saisies_df['Nom / Prenom'] != nom_complet_a_supprimer]
                        conn.update(worksheet="Saisies", data=updated_saisies_df)
                except Exception as e:
                    st.error(f"Le joueur est supprimé, mais son historique n'a pas pu être effacé : {e}")

            st.cache_data.clear()
            st.warning(f"⚠️ {nom_complet_a_supprimer} et tout son historique d'entraînement ont été supprimés.")
            st.rerun()

st.markdown("---")

# ── AJOUT D'UN JOUEUR ────────────────────────────────────────────────────────────
st.subheader("➕ Ajouter un joueur")

with st.form("add_player_form"):
    col_a, col_b = st.columns(2)
    with col_a:
        prenom = st.text_input("Prénom *", placeholder="Antoine")
        poste  = st.selectbox("Poste *", POSTES_RUGBY)
    with col_b:
        nom    = st.text_input("Nom *", placeholder="Dupont")
        groupe = st.selectbox("Groupe", GROUPES, index=0)

    submitted = st.form_submit_button("➕ Ajouter le joueur", type="secondary")

    if submitted:
        if not nom.strip() or not prenom.strip():
            st.error("Le nom et le prénom sont obligatoires.")
        else:
            new_player = pd.DataFrame([{
                "id": str(uuid.uuid4()),
                "nom": nom.strip().upper(),
                "prenom": prenom.strip().capitalize(),
                "poste": poste,
                "numero": 0,
                "groupe": groupe
            }])
            
            updated_df = pd.concat([players_df, new_player], ignore_index=True)
            conn.update(worksheet="Effectif", data=updated_df)
            st.cache_data.clear()
            
            st.success(f"✅ {prenom.capitalize()} {nom.upper()} ajouté(e) au poste de {poste}.")
            st.rerun()
