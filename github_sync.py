"""
Module de Synchronisation GitHub (GitHub REST API)
Club : Rugby Club Amiénois (RCA) - Catapult GPS Processor
Description :
    Assure la persistance permanente des données en archivant automatiquement
    sur le dépôt GitHub tous les fichiers modifiés dans l'application Streamlit :
    - Fichiers de séances hebdomadaires (data/weekly_memory/week_YYYY_WXX.json)
    - Configuration des objectifs athlétiques par poste (data/targets_config.json)
    - Effectif personnalisé et nouveaux joueurs (config_rca_joueurs_custom.json)
    - Photos des joueurs (assets/players_photos/...)
    - Mapping des boîtiers GPS (config_rca_gps.json)

    Évite ainsi toute perte de données lors des mises en veille (sleep mode)
    et redémarrages de Streamlit Community Cloud.
"""

from __future__ import annotations

import base64
import os
import pathlib
import posixpath
import re
import subprocess
from typing import Any, Dict, Optional, Tuple, Union

import requests

PROJECT_ROOT = pathlib.Path(__file__).parent.resolve()
if PROJECT_ROOT.name == "modules":
    PROJECT_ROOT = PROJECT_ROOT.parent
DEFAULT_REPO = "Quent123456/rca-unified-monitoring"
DEFAULT_BRANCH = "main"


def _detect_repo_from_git() -> str:
    """Tente de détecter le nom du dépôt (ex: 'Quent123456/Traitement-GPS-RCA') via git config."""
    try:
        res = subprocess.run(
            ["git", "config", "--get", "remote.origin.url"],
            cwd=str(PROJECT_ROOT),
            capture_output=True,
            text=True,
            timeout=3,
        )
        if res.returncode == 0 and res.stdout:
            url = res.stdout.strip()
            m = re.search(r"[:/]([^/:]+/[^/:]+?)(?:\.git)?$", url)
            if m:
                return m.group(1)
    except Exception:
        pass
    return DEFAULT_REPO


def get_github_credentials() -> Dict[str, str]:
    """
    Récupère les identifiants de synchronisation GitHub :
    1. Depuis les variables d'environnement (GITHUB_TOKEN, GITHUB_REPO, GITHUB_BRANCH)
    2. Depuis st.secrets (format plat ou table [github])
    3. Avec fallbacks intelligents (remote origin git et branche main).
    """
    token = os.environ.get("GITHUB_TOKEN", "").strip()
    repo = os.environ.get("GITHUB_REPO", "").strip()
    branch = os.environ.get("GITHUB_BRANCH", "").strip()

    try:
        import streamlit as st

        # Recherche dans st.secrets direct
        if not token and "GITHUB_TOKEN" in st.secrets:
            token = str(st.secrets["GITHUB_TOKEN"]).strip()
        if not repo and "GITHUB_REPO" in st.secrets:
            repo = str(st.secrets["GITHUB_REPO"]).strip()
        if not branch and "GITHUB_BRANCH" in st.secrets:
            branch = str(st.secrets["GITHUB_BRANCH"]).strip()

        # Recherche dans table st.secrets["github"]
        if "github" in st.secrets and isinstance(st.secrets["github"], dict):
            gh_sec = st.secrets["github"]
            if not token:
                token = str(gh_sec.get("token", "")).strip()
            if not repo:
                repo = str(gh_sec.get("repo", "")).strip()
            if not branch:
                branch = str(gh_sec.get("branch", "")).strip()
    except Exception:
        pass

    if not repo:
        repo = _detect_repo_from_git()

    if not branch:
        branch = DEFAULT_BRANCH

    return {
        "token": token,
        "repo": repo,
        "branch": branch,
    }


def is_github_sync_available() -> bool:
    """Retourne True si un token GitHub et un dépôt valide sont configurés."""
    creds = get_github_credentials()
    return bool(creds.get("token") and creds.get("repo"))


def get_github_sync_status() -> Tuple[bool, str, str]:
    """
    Retourne l'état de la synchronisation sous la forme :
    (is_active, status_message, repo_label)
    """
    creds = get_github_credentials()
    token = creds.get("token", "")
    repo = creds.get("repo", DEFAULT_REPO)
    branch = creds.get("branch", DEFAULT_BRANCH)

    if token:
        masked_tok = token[:4] + "..." + token[-4:] if len(token) > 8 else "***"
        return True, f"Synchronisation GitHub active ({masked_tok})", f"{repo} [{branch}]"
    else:
        return False, "Mode local uniquement (Token GitHub non configuré)", f"{repo} [{branch}]"


def _normalize_repo_path(raw_path: Union[str, pathlib.Path]) -> str:
    """Normalise un chemin de fichier pour qu'il soit propre et compatible POSIX pour l'API GitHub."""
    if isinstance(raw_path, pathlib.Path):
        p_obj = raw_path
    else:
        p_obj = pathlib.Path(str(raw_path).strip())

    # Si c'est un chemin absolu dans PROJECT_ROOT, calculer le relatif
    try:
        p_abs = p_obj.resolve()
        if PROJECT_ROOT in p_abs.parents or p_abs == PROJECT_ROOT:
            rel = p_abs.relative_to(PROJECT_ROOT)
            return rel.as_posix()
    except Exception:
        pass

    clean_str = str(raw_path).replace("\\", "/").lstrip("/")
    if clean_str.startswith("./"):
        clean_str = clean_str[2:]
    return clean_str


def get_file_sha_from_github(repo_file_path: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Interroge l'API GitHub pour récupérer le SHA actuel du fichier sur la branche cible.
    Retourne (sha, error_message). Si le fichier n'existe pas encore (404), retourne (None, None).
    """
    creds = get_github_credentials()
    token = creds.get("token")
    repo = creds.get("repo")
    branch = creds.get("branch", DEFAULT_BRANCH)

    if not token or not repo:
        return None, "Identifiants GitHub manquants"

    norm_path = _normalize_repo_path(repo_file_path)
    url = f"https://api.github.com/repos/{repo}/contents/{norm_path}"
    headers = {
        "Accept": "application/vnd.github.v3+json",
        "Authorization": f"Bearer {token}",
        "User-Agent": "Catapult-GPS-RCA-Streamlit",
    }

    try:
        r = requests.get(url, headers=headers, params={"ref": branch}, timeout=10)
        if r.status_code == 200:
            return r.json().get("sha"), None
        elif r.status_code == 404:
            return None, None
        else:
            return None, f"Erreur GET GitHub ({r.status_code}) : {r.text[:200]}"
    except Exception as e:
        return None, f"Erreur réseau lors de la vérification GitHub : {str(e)}"


def push_file_to_github(
    local_file_path: Union[str, pathlib.Path],
    repo_file_path: Optional[str] = None,
    commit_message: Optional[str] = None,
    show_toast: bool = False,
) -> Tuple[bool, str]:
    """
    Pousse un fichier local vers le dépôt GitHub via l'API REST Contents (PUT).
    - Récupère le sha s'il existe déjà sur la branche pour permettre la mise à jour (UPDATE).
    - Encode le contenu en Base64.
    - Gère les conflits 409 avec une nouvelle tentative automatique.
    - Ne fait jamais planter l'application en cas d'échec réseau ou d'absence de token.
    """
    creds = get_github_credentials()
    token = creds.get("token")
    repo = creds.get("repo")
    branch = creds.get("branch", DEFAULT_BRANCH)

    if not token:
        msg = "GITHUB_TOKEN non configuré. Sauvegarde effectuée en local uniquement."
        return False, msg

    if not repo:
        msg = "GITHUB_REPO introuvable."
        return False, msg

    loc_path = pathlib.Path(local_file_path)
    if not loc_path.is_absolute():
        loc_path = PROJECT_ROOT / loc_path

    if not loc_path.exists():
        return False, f"Le fichier local '{loc_path}' n'existe pas."

    try:
        with open(loc_path, "rb") as f:
            raw_bytes = f.read()
    except Exception as e:
        return False, f"Impossible de lire le fichier local : {str(e)}"

    if not repo_file_path:
        repo_file_path = _normalize_repo_path(loc_path)
    else:
        repo_file_path = _normalize_repo_path(repo_file_path)

    return push_content_to_github(
        content_bytes=raw_bytes,
        repo_file_path=repo_file_path,
        commit_message=commit_message or f"data: mise à jour {repo_file_path}",
        show_toast=show_toast,
    )


def push_content_to_github(
    content_bytes: bytes,
    repo_file_path: str,
    commit_message: Optional[str] = None,
    show_toast: bool = False,
) -> Tuple[bool, str]:
    """Pousse directement des octets en Base64 vers l'API GitHub."""
    creds = get_github_credentials()
    token = creds.get("token")
    repo = creds.get("repo")
    branch = creds.get("branch", DEFAULT_BRANCH)

    if not token or not repo:
        return False, "Identifiants GitHub manquants (Mode local actif)"

    norm_path = _normalize_repo_path(repo_file_path)
    content_b64 = base64.b64encode(content_bytes).decode("ascii")

    sha, _ = get_file_sha_from_github(norm_path)

    url = f"https://api.github.com/repos/{repo}/contents/{norm_path}"
    headers = {
        "Accept": "application/vnd.github.v3+json",
        "Authorization": f"Bearer {token}",
        "User-Agent": "Catapult-GPS-RCA-Streamlit",
    }

    body: Dict[str, Any] = {
        "message": commit_message or f"data: synchronisation {norm_path}",
        "content": content_b64,
        "branch": branch,
    }
    if sha:
        body["sha"] = sha

    try:
        r = requests.put(url, headers=headers, json=body, timeout=15)
        if r.status_code in (200, 201):
            commit_sha = r.json().get("commit", {}).get("sha", "")[:7]
            msg = f"Synchronisation GitHub réussie pour '{norm_path}' ({commit_sha})."
            if show_toast:
                try:
                    import streamlit as st
                    st.toast(f"☁️ GitHub Sync : {norm_path} sauvegardé !")
                except Exception:
                    pass
            return True, msg
        elif r.status_code == 409:
            # Conflit de version : re-tenter avec le nouveau SHA
            new_sha, _ = get_file_sha_from_github(norm_path)
            if new_sha:
                body["sha"] = new_sha
                r_retry = requests.put(url, headers=headers, json=body, timeout=15)
                if r_retry.status_code in (200, 201):
                    msg = f"Synchronisation GitHub réussie après résolution de conflit pour '{norm_path}'."
                    return True, msg
            return False, f"Conflit de version GitHub (409) sur '{norm_path}'."
        else:
            err_detail = r.text[:250]
            return False, f"Erreur API GitHub ({r.status_code}) : {err_detail}"
    except Exception as e:
        return False, f"Erreur de communication avec GitHub : {str(e)}"


def delete_file_from_github(
    repo_file_path: Union[str, pathlib.Path],
    commit_message: Optional[str] = None,
    show_toast: bool = False,
) -> Tuple[bool, str]:
    """
    Supprime un fichier sur le dépôt GitHub via l'API REST Contents (DELETE).
    - Récupère d'abord le sha du fichier.
    - Si le fichier n'existe pas sur GitHub (404), considère l'opération réussie.
    - Envoie la requête DELETE avec le sha et le message de commit.
    """
    creds = get_github_credentials()
    token = creds.get("token")
    repo = creds.get("repo")
    branch = creds.get("branch", DEFAULT_BRANCH)

    if not token or not repo:
        return False, "Identifiants GitHub manquants (Mode local actif)"

    norm_path = _normalize_repo_path(repo_file_path)
    sha, err = get_file_sha_from_github(norm_path)

    if not sha:
        # Fichier déjà absent ou 404
        return True, f"Fichier '{norm_path}' déjà absent sur GitHub."

    url = f"https://api.github.com/repos/{repo}/contents/{norm_path}"
    headers = {
        "Accept": "application/vnd.github.v3+json",
        "Authorization": f"Bearer {token}",
        "User-Agent": "Catapult-GPS-RCA-Streamlit",
    }
    body = {
        "message": commit_message or f"data: suppression {norm_path}",
        "sha": sha,
        "branch": branch,
    }

    try:
        r = requests.delete(url, headers=headers, json=body, timeout=15)
        if r.status_code in (200, 204):
            msg = f"Fichier '{norm_path}' supprimé de GitHub avec succès."
            if show_toast:
                try:
                    import streamlit as st
                    st.toast(f"🗑️ GitHub : {norm_path} supprimé.")
                except Exception:
                    pass
            return True, msg
        else:
            return False, f"Erreur suppression GitHub ({r.status_code}) : {r.text[:250]}"
    except Exception as e:
        return False, f"Erreur lors de la suppression sur GitHub : {str(e)}"


def render_github_sync_badge_ui(in_sidebar: bool = True) -> None:
    """
    Composant d'interface Streamlit affichant le statut de la synchronisation GitHub,
    le dépôt connecté et les consignes de configuration sécurisée du token.
    """
    try:
        import streamlit as st
    except Exception:
        return

    is_active, status_msg, repo_label = get_github_sync_status()

    container = st.sidebar if in_sidebar else st.container()
    with container:
        if is_active:
            st.markdown(
                f"<div style='background-color:#F0FDF4; border:1px solid #BBF7D0; border-radius:8px; padding:10px 12px; margin-bottom:12px;'>"
                f"<div style='color:#15803D; font-weight:700; font-size:0.88rem; display:flex; align-items:center; gap:6px;'>"
                f"<span>🟢</span> Synchronisation GitHub Active</div>"
                f"<div style='color:#475569; font-size:0.75rem; margin-top:3px;'><b>Dépôt :</b> {repo_label}</div>"
                f"<div style='color:#64748B; font-size:0.72rem; margin-top:2px;'>Chaque séance ou objectif validé est immédiatement commité et sécurisé.</div>"
                f"</div>",
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f"<div style='background-color:#FFFBEB; border:1px solid #FDE68A; border-radius:8px; padding:10px 12px; margin-bottom:12px;'>"
                f"<div style='color:#B45309; font-weight:700; font-size:0.88rem; display:flex; align-items:center; gap:6px;'>"
                f"<span>🟠</span> Mode local uniquement</div>"
                f"<div style='color:#475569; font-size:0.75rem; margin-top:3px;'><b>Dépôt cible :</b> {repo_label}</div>"
                f"<div style='color:#92400E; font-size:0.72rem; margin-top:3px;'>Token GitHub manquant dans <code>st.secrets</code>. Vos données restent sauvegardées en local mais risquent d'être réinitialisées lors des veilles Streamlit Cloud.</div>"
                f"</div>",
                unsafe_allow_html=True,
            )
            with st.expander("ℹ️ Activer la synchronisation GitHub (Anti-perte Cloud)", expanded=False):
                st.markdown(
                    """
                    **Pour sauvegarder définitivement chaque séance sur votre GitHub :**
                    1. Sur GitHub, créez un token :  
                       *Settings > Developer settings > Personal access tokens (classic)*  
                       avec la coche **`repo`** (lecture/écriture).
                    2. Sur votre tableau de bord **Streamlit Community Cloud** :  
                       Ouvrez l'application, cliquez sur **Settings > Secrets**.
                    3. Collez la configuration suivante :
                    ```toml
                    GITHUB_TOKEN = "ghp_votre_token_secret_ici"
                    GITHUB_REPO = "Quent123456/Traitement-GPS-RCA"
                    GITHUB_BRANCH = "main"
                    ```
                    Dès l'ajout du secret, la pastille passera au vert 🟢 et chaque action sera protégée !
                    """,
                    unsafe_allow_html=True,
                )
