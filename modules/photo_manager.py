"""
==============================================================================
MODULE DE GESTION DES PHOTOS DES JOUEURS — RUGBY CLUB AMIÉNOIS (RCA)
==============================================================================
Gère l'extraction automatique des photos depuis le trombinoscope officiel
'EFFECTIF SÉNIORS (1).pdf', la normalisation, le matching géométrique et fuzzy,
l'encodage Base64 pour l'interface web Streamlit et les exports PDF, ainsi que
l'administration manuelle des photos d'effectif.
==============================================================================
"""

import os
import io
import re
import base64
import unicodedata
import difflib
from typing import Dict, List, Optional, Tuple, Any, Union
from PIL import Image

try:
    import pymupdf  # PyMuPDF (fitz)
except ImportError:
    try:
        import fitz as pymupdf
    except ImportError:
        pymupdf = None

# ==============================================================================
# CHEMINS DES DOSSIERS ET FICHIERS
# ==============================================================================

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS_DIR = os.path.join(PROJECT_ROOT, "assets")
PLAYERS_PHOTOS_DIR = os.path.join(ASSETS_DIR, "players_photos")

# Création automatique du dossier local s'il n'existe pas
os.makedirs(PLAYERS_PHOTOS_DIR, exist_ok=True)

DEFAULT_PDF_CANDIDATES = [
    os.path.join(PROJECT_ROOT, "EFFECTIF SÉNIORS (1).pdf"),
    os.path.join(PROJECT_ROOT, "EFFECTIF_SENIORS.pdf"),
    os.path.join(PROJECT_ROOT, "EFFECTIF SENIORS (1).pdf"),
    os.path.join(PROJECT_ROOT, "EFFECTIF SÉNIORS.pdf"),
    os.path.join(r"C:\Users\33686\Downloads", "EFFECTIF SÉNIORS (1).pdf"),
    os.path.join(r"C:\Users\33686\Downloads", "EFFECTIF SÉNIORS (1) (1).pdf"),
]

# Aliases et surnoms fréquents au sein du club
KNOWN_ALIASES: Dict[str, str] = {
    "bam_s": "clement_bamiere",
    "bams": "clement_bamiere",
    "clement_bam_s": "clement_bamiere",
    "tij": "jean_etienne_lesueur",
    "jean_etienne_tij": "jean_etienne_lesueur",
    "jean_etienne": "jean_etienne_lesueur",
    "maka": "maka_salua",
    "salua_makatea": "maka_salua",
    "makatea_salua": "maka_salua",
    "delacroix_jules": "jules_blanc",
    "jules_delacroix": "jules_blanc",
    "eliot_boeuf": "elliot_boeuf",
    "boeuf_eliot": "elliot_boeuf",
    "boeuf_elliot": "elliot_boeuf",
    "octave_de_bruyn": "octave_de_bryun",
    "de_bruyn_octave": "octave_de_bryun",
    "prevote_olivier": "oliver_prevote",
    "olivier_prevote": "oliver_prevote",
}

# Liste de référence officielle extraite du PDF par page et par poste
REFERENCE_EFFECTIF_PDF: Dict[str, List[Dict[str, str]]] = {
    "Page 1 - Piliers": [
        {"nom_complet": "Noam FRION", "poste": "Pilier"},
        {"nom_complet": "Landry SALUA", "poste": "Pilier"},
        {"nom_complet": "Quentin COZIC", "poste": "Pilier"},
        {"nom_complet": "Adrien PECQUERY", "poste": "Pilier"},
        {"nom_complet": "Thomas TRAMON", "poste": "Pilier"},
        {"nom_complet": "Giovanni FOURNET", "poste": "Pilier"},
        {"nom_complet": "Nolan MAZE", "poste": "Pilier"},
        {"nom_complet": "Jérémy BICHARI", "poste": "Pilier"},
        {"nom_complet": "Thibault LEFEVER", "poste": "Pilier"},
        {"nom_complet": "Arnaud ROGIER", "poste": "Pilier"},
        {"nom_complet": "Sebastien VASELLI", "poste": "Pilier"},
        {"nom_complet": "Edouard TANTER", "poste": "Pilier"},
    ],
    "Page 1 - Talonneurs": [
        {"nom_complet": "Jean-Étienne LESUEUR", "poste": "Talonneur"},
        {"nom_complet": "Lucas SAINTOMER", "poste": "Talonneur"},
        {"nom_complet": "Roméo GENERMONT", "poste": "Talonneur"},
        {"nom_complet": "Samuel GUIBERT", "poste": "Talonneur"},
        {"nom_complet": "Bastien DAULT", "poste": "Talonneur"},
    ],
    "Page 2 - 2èmes lignes": [
        {"nom_complet": "Thomas WAEYAERT", "poste": "2ème ligne"},
        {"nom_complet": "Paul GRENET", "poste": "2ème ligne"},
        {"nom_complet": "Alejo SOTILLO", "poste": "2ème ligne"},
        {"nom_complet": "Simon VDH", "poste": "2ème ligne"},
        {"nom_complet": "Gabin DENIER", "poste": "2ème ligne"},
        {"nom_complet": "Etienne ROY", "poste": "2ème ligne"},
        {"nom_complet": "Donovan BOIDIN", "poste": "2ème ligne"},
        {"nom_complet": "Victorien FOUQUET", "poste": "2ème ligne"},
    ],
    "Page 3 - 3èmes lignes": [
        {"nom_complet": "Enzo VASELLI", "poste": "3ème ligne"},
        {"nom_complet": "Emilien LAFEUILLE", "poste": "3ème ligne"},
        {"nom_complet": "Benjamin MARTIN", "poste": "3ème ligne"},
        {"nom_complet": "Ciro MONTI", "poste": "3ème ligne"},
        {"nom_complet": "Louis WIOTTE", "poste": "3ème ligne"},
        {"nom_complet": "Tim GENTIEN", "poste": "3ème ligne"},
        {"nom_complet": "Lucas ORION", "poste": "3ème ligne"},
        {"nom_complet": "Simon CARPENTIER", "poste": "3ème ligne"},
        {"nom_complet": "Louis CARLIER", "poste": "3ème ligne"},
        {"nom_complet": "Léandre COSSARD", "poste": "3ème ligne"},
        {"nom_complet": "Louis SAINTOMER", "poste": "3ème ligne"},
        {"nom_complet": "Adrien AUDIERE", "poste": "3ème ligne"},
        {"nom_complet": "Alban BINNINGER", "poste": "3ème ligne"},
        {"nom_complet": "Jules BLANC", "poste": "3ème ligne"},
        {"nom_complet": "Camille KAZUMBA", "poste": "3ème ligne"},
    ],
    "Page 4 - Demis de mêlée": [
        {"nom_complet": "Gabriel TESSON", "poste": "Demi de mêlée"},
        {"nom_complet": "Clément BAMIERE", "poste": "Demi de mêlée"},
        {"nom_complet": "Stephan SANTONI", "poste": "Demi de mêlée"},
        {"nom_complet": "Elliot BOEUF", "poste": "Demi de mêlée"},
    ],
    "Page 4 - Demis d'ouverture": [
        {"nom_complet": "Max HARRY", "poste": "Demi d'ouverture"},
        {"nom_complet": "Maël CARPENTIER", "poste": "Demi d'ouverture"},
        {"nom_complet": "Victor ULMUCK", "poste": "Demi d'ouverture"},
        {"nom_complet": "Louis GOMES", "poste": "Demi d'ouverture"},
        {"nom_complet": "Xavier SUEUR", "poste": "Demi d'ouverture"},
    ],
    "Page 5 - Centres": [
        {"nom_complet": "Oliver PRÉVOTÉ", "poste": "Centre"},
        {"nom_complet": "Ope PELESEUMA", "poste": "Centre"},
        {"nom_complet": "Brieuc TREMBLAY", "poste": "Centre"},
        {"nom_complet": "Octave DE BRYUN", "poste": "Centre"},
        {"nom_complet": "Joseph CARPENTIER", "poste": "Centre"},
        {"nom_complet": "Maka SALUA", "poste": "Centre"},
        {"nom_complet": "Antoine SELTZER", "poste": "Centre"},
        {"nom_complet": "Léonard BATON", "poste": "Centre"},
        {"nom_complet": "Nathan MARZIANO", "poste": "Centre"},
    ],
    "Page 5 - Ailiers": [
        {"nom_complet": "Benoît VENIN", "poste": "Ailier"},
        {"nom_complet": "Oscar HURTEL", "poste": "Ailier"},
        {"nom_complet": "Youness SRIDI", "poste": "Ailier"},
        {"nom_complet": "Mathis LEROUX", "poste": "Ailier"},
        {"nom_complet": "Lucien LESPINASSE", "poste": "Ailier"},
        {"nom_complet": "Thomas PRÉVOST", "poste": "Ailier"},
    ],
    "Page 5 - Arrières": [
        {"nom_complet": "Arthur MOINE", "poste": "Arrière"},
        {"nom_complet": "Beryl VILAIN", "poste": "Arrière"},
    ],
}


# ==============================================================================
# NORMALISATION ET IDENTIFICATION DE JOUEUR
# ==============================================================================

def normalize_name(s: str) -> str:
    """
    Normalise une chaîne de nom/prénom sous la forme 'prenom_nom'
    (minuscules, sans accents, sans caractères spéciaux).
    Ex: 'Landry SALUA' -> 'landry_salua'
        'Jean-Étienne LESUEUR' -> 'jean_etienne_lesueur'
        'BAM\\'S' -> 'bam_s'
    """
    if not s:
        return ""
    # Décomposition unicode pour retirer les accents
    s_norm = unicodedata.normalize("NFD", str(s).strip())
    s_clean = "".join(c for c in s_norm if unicodedata.category(c) != "Mn")
    # Remplacement des apostrophes et tirets par underscores
    s_clean = s_clean.replace("'", "_").replace("-", "_").replace("(", " ").replace(")", " ")
    # Remplacement de toute séquence non alphanumérique par un underscore
    s_clean = re.sub(r"[^a-zA-Z0-9]+", "_", s_clean.lower()).strip("_")
    return s_clean


def get_player_initials(player_name: str) -> str:
    """Extrait les 1 ou 2 initiales majuscules d'un joueur (ex: 'Thomas PECQUERY' -> 'TP')."""
    if not player_name or player_name.strip() in ["-", "", "— Sélectionner un joueur —"]:
        return "RCA"
    parts = re.split(r"[\s\-_]+", player_name.strip())
    valid_parts = [p for p in parts if p and not p.startswith("(")]
    if not valid_parts:
        return "RCA"
    if len(valid_parts) == 1:
        return valid_parts[0][:2].upper()
    return f"{valid_parts[0][0]}{valid_parts[-1][0]}".upper()


# ==============================================================================
# RECHERCHE DE LA PHOTO D'UN JOUEUR
# ==============================================================================

def get_player_photo_path(player_name: str) -> Optional[str]:
    """
    Recherche et retourne le chemin absolu de la photo d'un joueur dans assets/players_photos/.
    Applique une chaîne de matching progressive et robuste :
    1. Correspondance exacte sur le nom normalisé (prenom_nom.png, nom_prenom.png).
    2. Correspondance via le dictionnaire d'aliases (BAM'S, TIJ, etc.).
    3. Correspondance par ensemble de tokens (mots) composant le nom.
    4. Fuzzy matching avec difflib contre les fichiers existants.
    """
    if not player_name or str(player_name).strip() in ["", "-", "— Sélectionner un joueur —"]:
        return None

    raw_name = str(player_name).strip()
    norm = normalize_name(raw_name)
    if not norm:
        return None

    # Vérification des fichiers dans PLAYERS_PHOTOS_DIR
    if not os.path.exists(PLAYERS_PHOTOS_DIR):
        return None

    files = [f for f in os.listdir(PLAYERS_PHOTOS_DIR) if not f.startswith(".")]
    if not files:
        return None

    files_map = {os.path.splitext(f)[0].lower(): f for f in files}

    # 1. Correspondance directe
    if norm in files_map:
        return os.path.join(PLAYERS_PHOTOS_DIR, files_map[norm])

    # 2. Inversion Prénom / Nom (ex: 'pecquery_thomas' vs 'thomas_pecquery')
    tokens = [t for t in norm.split("_") if t]
    if len(tokens) >= 2:
        inverted = f"{tokens[-1]}_{'_'.join(tokens[:-1])}"
        if inverted in files_map:
            return os.path.join(PLAYERS_PHOTOS_DIR, files_map[inverted])
        first_last_inv = f"{'_'.join(tokens[1:])}_{tokens[0]}"
        if first_last_inv in files_map:
            return os.path.join(PLAYERS_PHOTOS_DIR, files_map[first_last_inv])

    # 3. Correspondance via les alias connus
    alias_target = KNOWN_ALIASES.get(norm)
    if alias_target and alias_target in files_map:
        return os.path.join(PLAYERS_PHOTOS_DIR, files_map[alias_target])

    for alias_key, alias_val in KNOWN_ALIASES.items():
        if alias_key in norm and alias_val in files_map:
            return os.path.join(PLAYERS_PHOTOS_DIR, files_map[alias_val])

    # 4. Matching par tokens d'inclusion (si tous les mots principaux sont dans le fichier)
    sig_tokens = set(tokens)
    for base_f, orig_f in files_map.items():
        f_tokens = set(base_f.split("_"))
        if sig_tokens and sig_tokens.issubset(f_tokens):
            return os.path.join(PLAYERS_PHOTOS_DIR, orig_f)
        if f_tokens and f_tokens.issubset(sig_tokens):
            return os.path.join(PLAYERS_PHOTOS_DIR, orig_f)

    # 5. Fuzzy matching avec difflib (seuil de tolérance 0.72)
    matches = difflib.get_close_matches(norm, list(files_map.keys()), n=1, cutoff=0.72)
    if matches:
        return os.path.join(PLAYERS_PHOTOS_DIR, files_map[matches[0]])

    return None


# ==============================================================================
# ENCODAGE BASE64 ET AVATAR PAR DÉFAUT (STYLE RCA)
# ==============================================================================

def generate_rca_avatar_svg_data_uri(player_name: str, size: int = 80) -> str:
    """
    Génère un avatar vectoriel SVG autonome aux couleurs officielles du RCA Amiens :
    - Fond : Bleu Nuit RCA (#0B192C / #1F4E79 dégradé)
    - Cercle liseré : Doré officiel RCA (#D4AF37)
    - Texte : Initiales du joueur en blanc immaculé
    - Totalement autonome, aucun appel réseau extérieur requis.
    """
    initials = get_player_initials(player_name)
    font_size = int(size * 0.40)
    svg = f"""<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 {size} {size}">
      <defs>
        <linearGradient id="rcaGrad" x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stop-color="#003399" />
          <stop offset="100%" stop-color="#1F4E79" />
        </linearGradient>
      </defs>
      <circle cx="{size/2}" cy="{size/2}" r="{size/2 - 2}" fill="url(#rcaGrad)" stroke="#FFD700" stroke-width="2.5" />
      <text x="50%" y="54%" font-family="Helvetica, Arial, sans-serif" font-size="{font_size}" font-weight="bold" fill="#FFFFFF" text-anchor="middle" dominant-baseline="middle" letter-spacing="1">
        {initials}
      </text>
    </svg>"""
    b64_svg = base64.b64encode(svg.encode("utf-8")).decode("utf-8")
    return f"data:image/svg+xml;base64,{b64_svg}"


def get_player_image_b64(player_name: str, size: int = 80) -> str:
    """
    Retourne la source Base64 de la photo du joueur pour l'injection HTML :
    - Si la photo existe dans assets/players_photos/, lit le fichier et retourne 'data:image/png;base64,...'.
    - Si absente, génère un avatar vectoriel SVG élégant aux couleurs du club RCA (Bleu #003399 et Or #FFD700).
    """
    photo_path = get_player_photo_path(player_name)
    if photo_path and os.path.isfile(photo_path):
        try:
            with open(photo_path, "rb") as f:
                raw_bytes = f.read()
            ext = os.path.splitext(photo_path)[1].lower()
            mime = "image/png" if ext == ".png" else ("image/jpeg" if ext in [".jpg", ".jpeg"] else "image/webp")
            b64_img = base64.b64encode(raw_bytes).decode("utf-8")
            return f"data:{mime};base64,{b64_img}"
        except Exception:
            pass

    return generate_rca_avatar_svg_data_uri(player_name, size=size)


# ==============================================================================
# MOTEUR D'EXTRACTION DEPUIS LE TROMBINOSCOPE PDF
# ==============================================================================

def find_pdf_file() -> Optional[str]:
    """Trouve le chemin du trombinoscope PDF en vérifiant les emplacements possibles."""
    for p in DEFAULT_PDF_CANDIDATES:
        if os.path.isfile(p):
            return p
    # Recherche souple dans le répertoire du projet
    for f in os.listdir(PROJECT_ROOT):
        if "EFFECTIF" in f.upper() and f.lower().endswith(".pdf"):
            return os.path.join(PROJECT_ROOT, f)
    return None


def extract_all_photos_from_pdf(pdf_path: Optional[str] = None) -> Dict[str, Any]:
    """
    Extrait l'ensemble des portraits individuels du trombinoscope PDF officiel du RCA
    en associant géométriquement chaque photo au joueur correspondant.
    Sauvegarde les images dans assets/players_photos/ sous la forme 'prenom_nom.png'.
    """
    if pymupdf is None:
        return {
            "success": False,
            "error": "La bibliothèque PyMuPDF n'est pas installée. Veuillez lancer 'pip install PyMuPDF'.",
            "saved_count": 0,
        }

    resolved_path = pdf_path or find_pdf_file()
    if not resolved_path or not os.path.isfile(resolved_path):
        return {
            "success": False,
            "error": "Le fichier trombinoscope 'EFFECTIF SÉNIORS (1).pdf' est introuvable à la racine du projet.",
            "saved_count": 0,
        }

    os.makedirs(PLAYERS_PHOTOS_DIR, exist_ok=True)

    try:
        doc = pymupdf.open(resolved_path)
    except Exception as e:
        return {"success": False, "error": f"Erreur lors de l'ouverture du PDF : {str(e)}", "saved_count": 0}

    saved_photos: List[Dict[str, Any]] = []
    skipped_placeholders: List[str] = []

    section_titles = [
        "PILIERS", "TALONNEURS", "2ÈMES LIGNES", "2EMES LIGNES", "3EMES LIGNES",
        "DEMIS MÊLÉE", "DEMIS MELEE", "DEMIS D’OUVERTURE", "DEMIS D'OUVERTURE",
        "CENTRES", "AILIERS", "ARRIÈRE", "ARRIERE", "EFFECTIF SENIORS", "EFFECTIF SÉNIORS"
    ]

    for pno in range(len(doc)):
        page = doc[pno]

        # 1. Analyse des images de la page
        img_infos = page.get_image_info(xrefs=True)
        images = []
        for info in img_infos:
            xref = info["xref"]
            w, h = info["width"], info["height"]
            bbox = pymupdf.Rect(info["bbox"])

            # Filtrage du logo RCA club en haut à droite
            if bbox.x0 > 1250 and bbox.y0 < 180:
                continue

            # Identification des placeholders Canva (paysage colline/nuages 1438x560)
            is_placeholder = (w == 1438 and h == 560) or (w > 1000 and h < 600)

            cx = (bbox.x0 + bbox.x1) / 2
            cy = (bbox.y0 + bbox.y1) / 2
            images.append({
                "xref": xref,
                "bbox": bbox,
                "w": w,
                "h": h,
                "is_placeholder": is_placeholder,
                "cx": cx,
                "cy": cy,
            })

        # 2. Analyse du texte et regroupement des lignes
        text_words = page.get_text("words")
        lines_dict: Dict[Tuple[int, int], Dict[str, Any]] = {}
        for w in text_words:
            key = (w[5], w[6])
            if key not in lines_dict:
                lines_dict[key] = {"words": [], "rect": pymupdf.Rect(w[:4])}
            lines_dict[key]["words"].append(w[4])
            lines_dict[key]["rect"] |= pymupdf.Rect(w[:4])

        lines = []
        for k, v in lines_dict.items():
            line_str = " ".join(v["words"]).strip()
            if line_str.upper() in section_titles or not line_str:
                continue
            lines.append({
                "text": line_str,
                "rect": v["rect"],
                "cx": (v["rect"].x0 + v["rect"].x1) / 2,
                "cy": (v["rect"].y0 + v["rect"].y1) / 2,
                "y0": v["rect"].y0,
                "y1": v["rect"].y1,
            })

        # Appariement des lignes (Ligne 1 Prénom + Ligne 2 NOM)
        lines.sort(key=lambda l: (l["y0"], l["cx"]))
        used_lines = set()
        players_on_page = []

        for i, l1 in enumerate(lines):
            if i in used_lines:
                continue
            cand_j = None
            min_dist = float("inf")
            for j, l2 in enumerate(lines):
                if j == i or j in used_lines:
                    continue
                dy = l2["y0"] - l1["y0"]
                dx = abs(l2["cx"] - l1["cx"])
                if 15 <= dy <= 38 and dx <= 45:
                    dist = dx + dy
                    if dist < min_dist:
                        min_dist = dist
                        cand_j = j

            if cand_j is not None:
                l2 = lines[cand_j]
                used_lines.add(i)
                used_lines.add(cand_j)
                full_name = f"{l1['text']} {l2['text']}"
                p_cx = (l1['cx'] + l2['cx']) / 2
                p_cy = (l1['cy'] + l2['cy']) / 2
                players_on_page.append({
                    "name": full_name,
                    "cx": p_cx,
                    "cy": p_cy,
                })
            else:
                used_lines.add(i)
                players_on_page.append({
                    "name": l1["text"],
                    "cx": l1["cx"],
                    "cy": l1["cy"],
                })

        # Matching géométrique bipartite : image au-dessus du texte
        unmatched_imgs = list(range(len(images)))

        for p in players_on_page:
            best_idx = None
            best_cost = float("inf")
            for idx in unmatched_imgs:
                img = images[idx]
                dx = abs(img["cx"] - p["cx"])
                dy = p["cy"] - img["cy"]
                if dx < 75 and 20 < dy < 290:
                    cost = dx * 2.5 + abs(dy - 110)
                    if cost < best_cost:
                        best_cost = cost
                        best_idx = idx

            if best_idx is not None:
                unmatched_imgs.remove(best_idx)
                img = images[best_idx]
                norm_name = normalize_name(p["name"])

                if img["is_placeholder"]:
                    skipped_placeholders.append(p["name"])
                else:
                    try:
                        xref = img["xref"]
                        base_img = doc.extract_image(xref)
                        smask = base_img.get("smask")

                        pix = pymupdf.Pixmap(doc, xref)
                        if smask and smask > 0:
                            mask_pix = pymupdf.Pixmap(doc, smask)
                            pix = pymupdf.Pixmap(pix, mask_pix)

                        if pix.colorspace and pix.colorspace.name not in (pymupdf.csRGB.name, pymupdf.csGRAY.name):
                            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)

                        out_path = os.path.join(PLAYERS_PHOTOS_DIR, f"{norm_name}.png")
                        pix.save(out_path)
                        saved_photos.append({
                            "name": p["name"],
                            "filename": f"{norm_name}.png",
                            "path": out_path,
                            "width": pix.width,
                            "height": pix.height,
                            "page": pno + 1,
                        })
                    except Exception as err:
                        pass

    return {
        "success": True,
        "pdf_path": resolved_path,
        "saved_count": len(saved_photos),
        "saved_photos": saved_photos,
        "skipped_placeholders": skipped_placeholders,
    }


# ==============================================================================
# OPÉRATIONS D'ADMINISTRATION DES PHOTOS (UPLOAD, SUPPRESSION)
# ==============================================================================

def save_manual_player_photo(player_name: str, file_data: Union[bytes, io.BytesIO, Any], ext: str = ".png") -> Optional[str]:
    """Sauvegarde manuelle d'une photo pour un joueur donné."""
    norm = normalize_name(player_name)
    if not norm:
        return None

    clean_ext = ext.lower() if ext.startswith(".") else f".{ext.lower()}"
    if clean_ext not in [".png", ".jpg", ".jpeg", ".webp"]:
        clean_ext = ".png"

    dest_filename = f"{norm}{clean_ext}"
    dest_path = os.path.join(PLAYERS_PHOTOS_DIR, dest_filename)

    raw_bytes = file_data.getvalue() if hasattr(file_data, "getvalue") else (file_data.read() if hasattr(file_data, "read") else file_data)
    try:
        # Optimisation avec Pillow
        pil_img = Image.open(io.BytesIO(raw_bytes))
        pil_img.save(dest_path)
        return dest_path
    except Exception:
        try:
            with open(dest_path, "wb") as f:
                f.write(raw_bytes)
            return dest_path
        except Exception:
            return None


def delete_player_photo(player_name: str) -> bool:
    """
    Supprime la photo associée à un joueur dans assets/players_photos/
    pour éviter tout fichier orphelin.
    """
    norm = normalize_name(player_name)
    if not norm or not os.path.exists(PLAYERS_PHOTOS_DIR):
        return False

    deleted = False
    for f in os.listdir(PLAYERS_PHOTOS_DIR):
        base, _ = os.path.splitext(f)
        if base.lower() == norm.lower():
            p = os.path.join(PLAYERS_PHOTOS_DIR, f)
            try:
                os.remove(p)
                deleted = True
            except Exception:
                pass
    return deleted


def list_existing_player_photos() -> Dict[str, str]:
    """Retourne la liste de toutes les photos actuelles { 'prenom_nom': chemin_absolu }."""
    if not os.path.exists(PLAYERS_PHOTOS_DIR):
        return {}
    res = {}
    for f in os.listdir(PLAYERS_PHOTOS_DIR):
        if f.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
            base, _ = os.path.splitext(f)
            res[base] = os.path.join(PLAYERS_PHOTOS_DIR, f)
    return res


def ensure_photos_extracted() -> int:
    """
    Vérifie si les photos des joueurs sont présentes dans assets/players_photos/.
    Si moins de 30 photos sont présentes, lance automatiquement l'extraction depuis le PDF.
    Retourne le nombre de photos actuellement disponibles.
    """
    photos = list_existing_player_photos()
    if len(photos) < 30:
        res = extract_all_photos_from_pdf()
        if res.get("success"):
            photos = list_existing_player_photos()
    return len(photos)


# ==============================================================================
# COMPOSANT HTML / CSS PODIUM ULTRA-PRO ET DYNAMIQUE (STREAMLIT)
# ==============================================================================

def render_podium_html_markup(
    category_title: str,
    category_icon: str,
    players: List[Dict[str, Any]],
    unit_label: str = "",
    subtitle: Optional[str] = None,
    **kwargs: Any,
) -> str:
    """
    Génère le markup HTML/CSS premium du podium 3 marches :
    - 2ème marche à gauche (Argent #C0C0C0)
    - 1ère marche au centre surélevée (Or #FFD700)
    - 3ème marche à droite (Bronze #CD7F32)
    - Médaillon circulaire avec photo Base64 (80px au centre, 65px sur les côtés)
    - Ombre portée et bordure officielle
    - Nom, poste de rugby et valeur métrique en gras
    - Strictement sans commentaires HTML ni retours de ligne intra-style pour éviter le bug code block Markdown.
    """
    players_list = players if isinstance(players, list) else []
    p1 = players_list[0] if len(players_list) > 0 and isinstance(players_list[0], dict) else {"name": "-", "stat": "-", "poste": ""}
    p2 = players_list[1] if len(players_list) > 1 and isinstance(players_list[1], dict) else {"name": "-", "stat": "-", "poste": ""}
    p3 = players_list[2] if len(players_list) > 2 and isinstance(players_list[2], dict) else {"name": "-", "stat": "-", "poste": ""}

    p1_name = str(p1.get("name") or "-")
    p2_name = str(p2.get("name") or "-")
    p3_name = str(p3.get("name") or "-")

    img1 = p1.get("avatar_src") or get_player_image_b64(p1_name if p1_name != "-" else "", size=80)
    img2 = p2.get("avatar_src") or get_player_image_b64(p2_name if p2_name != "-" else "", size=65)
    img3 = p3.get("avatar_src") or get_player_image_b64(p3_name if p3_name != "-" else "", size=65)

    poste1 = str(p1.get("poste") or "")
    poste2 = str(p2.get("poste") or "")
    poste3 = str(p3.get("poste") or "")

    stat1 = str(p1.get("stat") or "-")
    stat2 = str(p2.get("stat") or "-")
    stat3 = str(p3.get("stat") or "-")

    sub_str = str(subtitle).strip() if subtitle is not None else ""
    sub_html = f'<div style="font-size:0.75rem;color:#FFD700;font-weight:700;letter-spacing:0.3px;margin-top:2px;">{sub_str}</div>' if sub_str else ""

    html = f"""<div style="background:#FFFFFF;border:1px solid #E2E8F0;border-radius:14px;padding:16px 12px 20px 12px;box-shadow:0 4px 16px rgba(15,23,42,0.06);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;margin-bottom:18px;">
<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(135deg,#0B192C 0%,#1F4E79 100%);color:#FFFFFF;padding:8px 14px;border-radius:8px;margin-bottom:22px;border-bottom:2px solid #D4AF37;">
<div style="display:flex;align-items:center;justify-content:center;gap:8px;">
<span style="font-size:1.15rem;">{category_icon}</span>
<span style="font-size:0.90rem;font-weight:800;letter-spacing:0.5px;text-transform:uppercase;">{category_title}</span>
</div>
{sub_html}
</div>
<div style="display:flex;align-items:flex-end;justify-content:center;gap:10px;padding:0 4px;">
<div style="flex:1;text-align:center;display:flex;flex-direction:column;align-items:center;">
<div style="position:relative;margin-bottom:6px;">
<img src="{img2}" alt="{p2_name}" style="width:65px;height:65px;border-radius:50%;object-fit:cover;border:3px solid #C0C0C0;box-shadow:0 4px 12px rgba(192,192,192,0.55);background:#F1F5F9;display:block;" />
<span style="position:absolute;bottom:-6px;left:50%;transform:translateX(-50%);background:#64748B;color:#FFFFFF;font-size:0.68rem;font-weight:800;padding:1px 7px;border-radius:10px;border:1px solid #FFFFFF;">2e</span>
</div>
<div style="font-size:0.78rem;font-weight:700;color:#1E293B;line-height:1.15;margin-top:6px;max-width:95px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="{p2_name}">{p2_name}</div>
<div style="font-size:0.68rem;color:#64748B;margin-bottom:4px;">{poste2}</div>
<div style="font-size:0.88rem;font-weight:800;color:#1F4E79;background:#F8FAFC;padding:3px 6px;border-radius:5px;border:1px solid #E2E8F0;margin-bottom:6px;">{stat2}</div>
<div style="width:100%;height:52px;background:linear-gradient(180deg,#94A3B8 0%,#64748B 100%);border-radius:6px 6px 0 0;display:flex;align-items:center;justify-content:center;color:#FFFFFF;font-size:1.15rem;font-weight:900;box-shadow:inset 0 2px 4px rgba(255,255,255,0.25);">2</div>
</div>
<div style="flex:1.15;text-align:center;display:flex;flex-direction:column;align-items:center;">
<div style="font-size:1.05rem;margin-bottom:-4px;">👑</div>
<div style="position:relative;margin-bottom:6px;">
<img src="{img1}" alt="{p1_name}" style="width:80px;height:80px;border-radius:50%;object-fit:cover;border:3.5px solid #FFD700;box-shadow:0 4px 16px rgba(217,119,6,0.45);background:#F1F5F9;display:block;" />
<span style="position:absolute;bottom:-6px;left:50%;transform:translateX(-50%);background:#D97706;color:#FFFFFF;font-size:0.72rem;font-weight:900;padding:1px 8px;border-radius:10px;border:1.5px solid #FFFFFF;">1er</span>
</div>
<div style="font-size:0.86rem;font-weight:800;color:#0F172A;line-height:1.15;margin-top:6px;max-width:110px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="{p1_name}">{p1_name}</div>
<div style="font-size:0.70rem;color:#64748B;margin-bottom:4px;">{poste1}</div>
<div style="font-size:0.96rem;font-weight:900;color:#A30000;background:#FEF2F2;padding:3px 8px;border-radius:5px;border:1px solid #FECACA;margin-bottom:6px;">{stat1}</div>
<div style="width:100%;height:74px;background:linear-gradient(180deg,#D4AF37 0%,#B45309 100%);border-radius:8px 8px 0 0;display:flex;align-items:center;justify-content:center;color:#FFFFFF;font-size:1.45rem;font-weight:900;box-shadow:inset 0 2px 4px rgba(255,255,255,0.35);">1</div>
</div>
<div style="flex:1;text-align:center;display:flex;flex-direction:column;align-items:center;">
<div style="position:relative;margin-bottom:6px;">
<img src="{img3}" alt="{p3_name}" style="width:65px;height:65px;border-radius:50%;object-fit:cover;border:3px solid #CD7F32;box-shadow:0 4px 12px rgba(205,127,50,0.45);background:#F1F5F9;display:block;" />
<span style="position:absolute;bottom:-6px;left:50%;transform:translateX(-50%);background:#B45309;color:#FFFFFF;font-size:0.68rem;font-weight:800;padding:1px 7px;border-radius:10px;border:1px solid #FFFFFF;">3e</span>
</div>
<div style="font-size:0.78rem;font-weight:700;color:#1E293B;line-height:1.15;margin-top:6px;max-width:95px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="{p3_name}">{p3_name}</div>
<div style="font-size:0.68rem;color:#64748B;margin-bottom:4px;">{poste3}</div>
<div style="font-size:0.88rem;font-weight:800;color:#1F4E79;background:#F8FAFC;padding:3px 6px;border-radius:5px;border:1px solid #E2E8F0;margin-bottom:6px;">{stat3}</div>
<div style="width:100%;height:38px;background:linear-gradient(180deg,#B45309 0%,#78350F 100%);border-radius:6px 6px 0 0;display:flex;align-items:center;justify-content:center;color:#FFFFFF;font-size:1.05rem;font-weight:900;box-shadow:inset 0 2px 4px rgba(255,255,255,0.25);">3</div>
</div>
</div>
</div>"""

    clean_html = "\n".join([line.lstrip() for line in html.splitlines() if line.strip()])
    return clean_html
