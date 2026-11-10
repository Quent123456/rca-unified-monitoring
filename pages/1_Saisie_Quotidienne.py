"""
pages/1_Saisie_Quotidienne.py
Formulaire de saisie quotidienne (100% Google Sheets) avec recherche, mémorisation et masquage dynamique.
"""

import streamlit as st
import pandas as pd
from datetime import date, datetime
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Saisie Quotidienne", page_icon="📝", layout="centered")

# ── Charte graphique officielle RCA et habillage Sidebar ──────────────────────
try:
    from modules.ui import inject_rca_page_layout
    inject_rca_page_layout("1. Saisie Quotidienne", "Formulaire RPE Foster & Bien-être Hooper")
except Exception:
    pass

st.title("Saisie Quotidienne")
st.markdown("Enregistrez la séance d'entraînement et le ressenti de chaque joueur.")

# ── CSS Personnalisé : Bouton Sticky & RPE ──────────────────────────────────────
st.markdown(
    """
    <style>
    /* Bouton d'enregistrement flottant (Sticky Bottom) */
    div.stButton:has(button[kind="primary"]) {
        position: fixed;
        bottom: 14px;
        left: 50%;
        transform: translateX(-50%);
        width: calc(100% - 32px);
        max-width: 680px;
        z-index: 9999;
        background: rgba(12, 35, 64, 0.94);
        backdrop-filter: blur(10px);
        padding: 10px 12px;
        border-radius: 16px;
        box-shadow: 0 -4px 20px rgba(12, 35, 64, 0.28);
        border: 1.5px solid rgba(242, 169, 0, 0.35);
    }
    /* Marge inférieure pour éviter que le bouton flottant ne cache les derniers champs */
    .main .block-container,
    [data-testid="stMainBlockContainer"] {
        padding-bottom: 110px !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── 1. Initialisation des mémoires (Session State) ──────────────────────────────
if "submitted_players" not in st.session_state:
    st.session_state.submitted_players = []
if "saved_duration" not in st.session_state:
    st.session_state.saved_duration = 75
if "saved_group" not in st.session_state:
    st.session_state.saved_group = "Groupe Élargi"

# ── 2. Connexion à Google Sheets et Chargement mis en cache ─────────────────────
@st.cache_data(ttl=600)
def load_players_data(_conn) -> pd.DataFrame:
    """Charge l'effectif depuis Google Sheets avec mise en cache (10 min)."""
    df = _conn.read(worksheet="Effectif", ttl=600)
    return df.dropna(how="all")

try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    players_df = load_players_data(conn).copy()
except Exception as e:
    st.error(f"Erreur de connexion à Google Sheets : {e}")
    st.stop()

if players_df.empty:
    st.warning("Aucun joueur enregistré. Ajoutez des joueurs dans **Gestion des Joueurs**.")
    st.stop()

# ── 3. Formatage, Tri et Filtrage de l'effectif ────────────────────────────────
# Formatage strict : NOM (majuscules) Prénom (majuscule initiale)
players_df["prenom"] = players_df["prenom"].fillna("").astype(str).str.strip().str.title()
players_df["nom"] = players_df["nom"].fillna("").astype(str).str.strip().str.upper()
players_df["poste"] = players_df["poste"].fillna("Inconnu")

# Tri alphabétique 
players_df = players_df.sort_values(by=["nom", "prenom"])

# Création des options
all_player_options = {}
for _, row in players_df.iterrows():
    if row['nom'] != "":
        label = f"{row['nom']} {row['prenom']} — {row['poste']}"
        all_player_options[label] = {
            "nom_complet": f"{row['nom']} {row['prenom']}",
            "poste": row["poste"]
        }

# Exclusion des joueurs déjà validés
available_options = {k: v for k, v in all_player_options.items() if k not in st.session_state.submitted_players}

# En-tête avec bouton de rafraîchissement
col_title, col_btn = st.columns([2, 1])
with col_btn:
    if st.button("🔄 Réafficher l'effectif", use_container_width=True):
        st.session_state.submitted_players = []
        st.rerun()

if not available_options:
    st.success("Toutes les saisies ont été effectuées ! Cliquez sur le bouton 'Réafficher l'effectif' pour recommencer.")
    st.stop()

# ── 4. Interface Dynamique ─────────────────────────────────────────────────────
# Recherche
search_term = st.text_input(
    "🔍 Rechercher un joueur (Nom ou Prénom) :",
    "",
    placeholder="Tapez un nom ou un prénom pour filtrer...",
).lower()

filtered_options = {
    k: v for k, v in available_options.items() 
    if search_term in k.lower()
}

if not filtered_options:
    st.warning("Aucun joueur trouvé. Vérifiez l'orthographe ou réaffichez l'effectif.")
    st.stop()

st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    selected_label = st.selectbox("Joueur sélectionné", options=list(filtered_options.keys()))
with col2:
    session_date = st.date_input("Date", value=date.today(), max_value=date.today())

# Synchronisation Date -> Jour
JOURS_SEMAINE = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
jour_auto = JOURS_SEMAINE[session_date.weekday()]
OPTIONS_JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Jour de match", "Dimanche"]
index_jour = OPTIONS_JOURS.index(jour_auto) if jour_auto in OPTIONS_JOURS else 1

col_sem, col_jour = st.columns(2)
with col_sem:
    semaine = st.number_input(
        "Numéro de semaine",
        min_value=1, max_value=52, value=session_date.isocalendar()[1],
    )
with col_jour:
    jour = st.selectbox("Jour d'entraînement", options=OPTIONS_JOURS, index=index_jour)

st.markdown("---")
st.subheader("Charge de séance")

RPE_LABELS = {
    0: "0 — Repos complet",
    1: "1 — Très très facile",
    2: "2 — Très facile",
    3: "3 — Facile",
    4: "4 — Effort modéré",
    5: "5 — Effort moyen",
    6: "6 — Effort un peu difficile",
    7: "7 — Difficile",
    8: "8 — Très difficile",
    9: "9 — Très très difficile",
    10: "10 — Maximal",
}

# Nuancier de couleurs dynamique : 0 = bleu clair -> 10 = rouge foncé
RPE_COLORS = {
    0:  "#38bdf8",  # Bleu clair (repos)
    1:  "#4ade80",  # Vert très clair
    2:  "#22c55e",  # Vert clair
    3:  "#16a34a",  # Vert
    4:  "#84cc16",  # Vert lime
    5:  "#eab308",  # Jaune
    6:  "#f59e0b",  # Ambre
    7:  "#ea580c",  # Orange
    8:  "#dc2626",  # Rouge
    9:  "#b91c1c",  # Rouge foncé
    10: "#7f1d1d",  # Rouge très foncé (maximal)
}

# Slider RPE pleine largeur (composant central prioritaire)
rpe_container = st.container(key="rpe_container")
with rpe_container:
    st.markdown('<div class="rpe-anchor" style="display:none;"></div>', unsafe_allow_html=True)
    rpe = st.select_slider(
        "🔥 RPE — Perception de l'effort (0 à 10)",
        options=list(RPE_LABELS.keys()),
        value=5,
        format_func=lambda x: RPE_LABELS[x],
    )

# Couleur dynamique selon la valeur RPE choisie
rpe_color = RPE_COLORS.get(rpe, "#eab308")
pct = int((rpe / 10) * 100)
fill_pct = 4 if rpe == 0 else pct
track_gradient = f"linear-gradient(to right, {rpe_color} 0%, {rpe_color} {fill_pct}%, rgba(140, 150, 160, 0.24) {fill_pct}%, rgba(140, 150, 160, 0.24) 100%)"

st.markdown(
    f"""
    <div style="background: linear-gradient(135deg, #0C2340 0%, #16365D 100%); border: 1.5px solid {rpe_color}; border-left: 6px solid {rpe_color}; border-radius: 12px; padding: 10px 16px; margin-top: 4px; margin-bottom: 16px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 3px 10px rgba(12, 35, 64, 0.16);">
        <span style="color: #E2E8F0; font-weight: 600; font-size: 0.92rem;">Niveau d'effort sélectionné</span>
        <span style="color: {rpe_color}; font-weight: 800; font-size: 1.05rem;">{RPE_LABELS[rpe]}</span>
    </div>
    <style>
    /* ── Redéfinition des variables natives Streamlit pour le conteneur RPE ── */
    .st-key-rpe_container,
    div[class*="st-key-rpe_container"],
    div:has(.rpe-anchor) {{
        --primary-color: {rpe_color} !important;
        --st-color-primary: {rpe_color} !important;
        padding: 8px 4px 4px 4px !important;
    }}

    /* ── Barre de fond (Track) : Pleine largeur, épaisseur 22px et dégradé dynamique ── */
    .st-key-rpe_container div[data-baseweb="slider"] div[style*="linear-gradient"],
    .st-key-rpe_container div[data-baseweb="slider"] > div > div > div:last-child,
    .st-key-rpe_container div[data-baseweb="slider"] > div > div:last-child,
    div[class*="st-key-rpe_container"] div[data-baseweb="slider"] div[style*="linear-gradient"],
    div[class*="st-key-rpe_container"] div[data-baseweb="slider"] > div > div > div:last-child,
    div[class*="st-key-rpe_container"] div[data-baseweb="slider"] > div > div:last-child,
    div:has(.rpe-anchor) div[data-baseweb="slider"] div[style*="linear-gradient"],
    div:has(.rpe-anchor) div[data-baseweb="slider"] > div > div > div:last-child,
    div:has(.rpe-anchor) div[data-baseweb="slider"] > div > div:last-child {{
        height: 22px !important;
        min-height: 22px !important;
        border-radius: 11px !important;
        background: {track_gradient} !important;
    }}

    /* ── Curseur (Thumb) : Large 34px tactile, contour blanc et halo ── */
    .st-key-rpe_container div[role="slider"],
    div[class*="st-key-rpe_container"] div[role="slider"],
    div:has(.rpe-anchor) div[role="slider"] {{
        background-color: {rpe_color} !important;
        background: {rpe_color} !important;
        border: 3.5px solid #ffffff !important;
        box-shadow: 0 0 12px {rpe_color}99, 0 3px 8px rgba(0, 0, 0, 0.30) !important;
        height: 34px !important;
        width: 34px !important;
        border-radius: 50% !important;
    }}

    /* ── Valeur au-dessus du curseur en badge sportif dynamique ── */
    .st-key-rpe_container div[data-testid="stSliderThumbValue"],
    div[class*="st-key-rpe_container"] div[data-testid="stSliderThumbValue"],
    div:has(.rpe-anchor) div[data-testid="stSliderThumbValue"] {{
        color: {rpe_color} !important;
        font-weight: 800 !important;
        font-size: 0.96rem !important;
        background-color: #0C2340 !important;
        padding: 4px 12px !important;
        border-radius: 999px !important;
        border: 2px solid {rpe_color} !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.28) !important;
        white-space: nowrap !important;
        margin-bottom: 10px !important;
    }}

    /* ── Sécurité : Conteneur extérieur transparent ── */
    .st-key-rpe_container div[data-baseweb="slider"] > div,
    div[class*="st-key-rpe_container"] div[data-baseweb="slider"] > div,
    div:has(.rpe-anchor) div[data-baseweb="slider"] > div {{
        background-color: transparent !important;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

col4, col5 = st.columns([1.2, 1])
with col4:
    duration = st.number_input("Durée (minutes)", min_value=0, max_value=300, value=st.session_state.saved_duration, step=5)
with col5:
    load = rpe * duration
    st.metric("Charge Foster (UA)", f"{load:.0f}")

st.markdown("---")
is_rest_day = st.checkbox("Jour de repos (charge = 0)")

group_options = ["Équipe A", "Équipe B", "Groupe Élargi"]
default_group_index = group_options.index(st.session_state.saved_group) if st.session_state.saved_group in group_options else 2

groupe_entrainement = st.radio("Groupe d'entraînement *", group_options, horizontal=True, index=default_group_index)

# ── 5. Questionnaire de bien-être & Commentaires (Facultatif) ─────────────────
# Règle vitale : Valeurs par défaut attribuées automatiquement (3/5 et notes vides)
fatigue = 3
courbatures = 3
sommeil = 3
notes = ""

st.markdown("---")
with st.expander("🧘 Questionnaire de bien-être & Commentaires (facultatif)", expanded=False):
    st.caption("ℹ️ Si non renseigné, la note de 3/5 est automatiquement attribuée sans commentaire.")
    col6, col7, col8 = st.columns(3)
    with col6:
        fatigue = st.select_slider("Fatigue", options=[1, 2, 3, 4, 5], value=3)
    with col7:
        courbatures = st.select_slider("Courbatures", options=[1, 2, 3, 4, 5], value=3)
    with col8:
        sommeil = st.select_slider("Sommeil", options=[1, 2, 3, 4, 5], value=3)

    notes = st.text_area("Notes / Observations (optionnel)", height=80)

# ── 6. Traitement Google Sheets (Bouton Sticky) ────────────────────────────────
if st.button("✅ Enregistrer la séance", type="primary", use_container_width=True):
    player_info = filtered_options[selected_label]
    nom_prenom = player_info["nom_complet"]
    poste = player_info["poste"]

    actual_rpe = 0 if is_rest_day else rpe
    actual_duration = 0 if is_rest_day else duration
    actual_load = 0.0 if is_rest_day else load

    # Règle vitale : Sécurisation absolue des scores de bien-être (3 par défaut) et commentaires ("" par défaut)
    final_fatigue = int(fatigue) if fatigue is not None else 3
    final_courbatures = int(courbatures) if courbatures is not None else 3
    final_sommeil = int(sommeil) if sommeil is not None else 3
    final_notes = str(notes).strip() if notes is not None else ""
    
    # Mise à jour des valeurs mémorisées pour le prochain joueur
    st.session_state.saved_duration = duration
    st.session_state.saved_group = groupe_entrainement

    with st.spinner("Enregistrement dans Google Sheets..."):
        try:
            saisies_df = conn.read(worksheet="Saisies", ttl=0)
            notes_gs = f"[{groupe_entrainement}] {final_notes}" if final_notes else f"[{groupe_entrainement}]"
            horodateur = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            
            new_row = pd.DataFrame([{
                "Cle de recherche": f"{semaine}{nom_prenom.replace(' ', '')}",
                "Horodateur": horodateur,
                "Difficulte seance": RPE_LABELS[actual_rpe],
                "Ressenti / Notes": notes_gs,
                "Nom / Prenom": nom_prenom,
                "Poste": poste,
                "Numero Semaine": semaine,
                "Jour seance": jour,
                "RPE (Chiffre)": actual_rpe,
                "Duree (min)": actual_duration,
                "Charge (UA)": actual_load,
                "Fatigue | Courbatures | Sommeil": f"Fatigue:{final_fatigue} | Courbatures:{final_courbatures} | Sommeil:{final_sommeil}"
            }])
            
            updated_df = pd.concat([saisies_df, new_row], ignore_index=True)
            conn.update(worksheet="Saisies", data=updated_df)
            
            # Ajout du joueur à la liste des "déjà saisis"
            st.session_state.submitted_players.append(selected_label)
            
            # Rechargement immédiat de la page pour le masquer
            st.rerun()
            
        except Exception as e:
            st.error(f"Erreur lors de l'enregistrement : {e}")
