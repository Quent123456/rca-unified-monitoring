"""
pages/1_Saisie_Quotidienne.py
Formulaire de saisie quotidienne (100% Google Sheets) avec recherche, mémorisation et masquage dynamique.
"""

import textwrap
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
    textwrap.dedent("""
<style>
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
.main .block-container,
[data-testid="stMainBlockContainer"] {
    padding-bottom: 110px !important;
}
</style>
""").strip(),
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
if rpe == 0:
    fill_stop = "0%"
elif rpe == 10:
    fill_stop = "100%"
else:
    fill_stop = f"calc(23px + (100% - 46px) * {rpe / 10:.1f})"

track_gradient = (
    f"linear-gradient(to right, transparent 0%, transparent {fill_stop}, "
    f"#1E293B {fill_stop}, #1E293B 100%), "
    "linear-gradient(90deg, #10B981 0%, #F59E0B 40%, #EF4444 80%, #7F1D1D 100%)"
)


def _rpe_sel(sub: str = "") -> str:
    """Retourne les sélecteurs CSS scopés au conteneur RPE (spécificité >= assets/style.css)."""
    suffix = f" {sub}" if sub else ""
    return ",\n".join([
        f':is([data-testid="stMain"], section.main) .st-key-rpe_container{suffix}',
        f':is([data-testid="stMain"], section.main) div[class*="st-key-rpe_container"]{suffix}',
        f':is([data-testid="stMain"], section.main) div:has(.rpe-anchor){suffix}',
    ])


rpe_banner_html = textwrap.dedent(f"""
<div style="background: linear-gradient(135deg, #0C2340 0%, #16365D 100%); border: 1.5px solid {rpe_color}; border-left: 6px solid {rpe_color}; border-radius: 12px; padding: 10px 16px; margin-top: 8px; margin-bottom: 18px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 3px 10px rgba(12, 35, 64, 0.16);">
<span style="color: #E2E8F0; font-weight: 600; font-size: 0.92rem;">Niveau d'effort sélectionné</span>
<span style="color: {rpe_color}; font-weight: 800; font-size: 1.05rem;">{RPE_LABELS[rpe]}</span>
</div>
""").strip()

rpe_css_rules = [
    f"{_rpe_sel()} {{\n    --primary-color: {rpe_color} !important;\n    --st-color-primary: {rpe_color} !important;\n    padding: 10px 6px 6px 6px !important;\n}}",
    f"{_rpe_sel('div[data-testid=\"stSlider\"]')} {{\n    padding: 16px 4px 22px 4px !important;\n}}",
    f"{_rpe_sel('div[data-testid=\"stSlider\"] label[data-testid=\"stWidgetLabel\"]')} {{\n    margin-bottom: 22px !important;\n    font-weight: 700 !important;\n    color: #0C2340 !important;\n}}",
    f"{_rpe_sel('div[data-testid=\"stSlider\"] div[data-baseweb=\"slider\"] > div:first-child')} {{\n    background-color: transparent !important;\n    background: transparent !important;\n    padding: 18px 0 14px 0 !important;\n}}",
    f"{_rpe_sel('div[data-testid=\"stSlider\"] div[data-baseweb=\"slider\"] > div:first-child > div')} {{\n    position: relative !important;\n    height: 46px !important;\n    min-height: 46px !important;\n    width: 100% !important;\n    border-radius: 999px !important;\n    background: transparent !important;\n}}",
    f"{_rpe_sel('div[data-testid=\"stSlider\"] div[data-baseweb=\"slider\"] > div:first-child > div > div:first-child')} {{\n    position: absolute !important;\n    top: 0 !important;\n    bottom: 0 !important;\n    left: 23px !important;\n    right: 23px !important;\n    height: 46px !important;\n    min-height: 46px !important;\n    margin: 0 !important;\n    padding: 0 !important;\n    border: none !important;\n    background: transparent !important;\n    box-sizing: border-box !important;\n    overflow: visible !important;\n}}",
    f"{_rpe_sel('div[data-testid=\"stSlider\"] div[data-baseweb=\"slider\"] > div:first-child > div > div:last-child')} {{\n    height: 46px !important;\n    min-height: 46px !important;\n    width: 100% !important;\n    border-radius: 999px !important;\n    border: 1.5px solid rgba(255, 255, 255, 0.14) !important;\n    background: {track_gradient} !important;\n    box-shadow: inset 0 2px 5px rgba(0, 0, 0, 0.35), 0 4px 12px rgba(12, 35, 64, 0.22) !important;\n    box-sizing: border-box !important;\n}}",
    f"{_rpe_sel('div[data-testid=\"stSlider\"] div[role=\"slider\"]')} {{\n    top: 0 !important;\n    height: 42px !important;\n    width: 42px !important;\n    border-radius: 50% !important;\n    background: #FFFFFF !important;\n    background-color: #FFFFFF !important;\n    border: 2px solid #E2E8F0 !important;\n    box-shadow: 0 4px 10px rgba(0, 0, 0, 0.35) !important;\n    box-sizing: border-box !important;\n    cursor: grab !important;\n    transition: box-shadow 0.12s ease, scale 0.12s ease !important;\n}}",
    f"{_rpe_sel('div[data-testid=\"stSlider\"] div[role=\"slider\"]:active')} {{\n    cursor: grabbing !important;\n    scale: 1.05 !important;\n    box-shadow: 0 6px 14px rgba(0, 0, 0, 0.42) !important;\n}}",
    f"{_rpe_sel('div[data-testid=\"stSlider\"] div[data-testid=\"stSliderThumbValue\"]')} {{\n    font-weight: 800 !important;\n    font-size: 0.95rem !important;\n    color: #F2A900 !important;\n    background-color: #0C2340 !important;\n    border: 1.5px solid #F2A900 !important;\n    border-radius: 999px !important;\n    padding: 3px 10px !important;\n    box-shadow: 0 3px 10px rgba(0, 0, 0, 0.28) !important;\n    top: auto !important;\n    bottom: calc(100% + 10px) !important;\n    white-space: nowrap !important;\n}}",
    f"{_rpe_sel('div[data-testid=\"stSlider\"] div[data-testid=\"stSliderTickBar\"]')} {{\n    opacity: 0.9 !important;\n    color: #334155 !important;\n    font-weight: 600 !important;\n    font-size: 0.78rem !important;\n    top: calc(100% - 6px) !important;\n    margin-top: 0 !important;\n}}",
]
rpe_css = "<style>\n" + "\n".join(rpe_css_rules) + "\n</style>"

st.markdown(rpe_banner_html, unsafe_allow_html=True)
st.markdown(rpe_css, unsafe_allow_html=True)

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
