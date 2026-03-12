"""
Gestion des utilisateurs — inscription et validation admin.
"""

import json
from datetime import datetime
from pathlib import Path

import bcrypt
import yaml
from yaml.loader import SafeLoader

import storage_utils as storage

USERS_FILE = Path('users.yaml')


# ── Helpers YAML ───────────────────────────────────────────────────────────────

def load_users() -> dict:
    with open(USERS_FILE, encoding='utf-8') as f:
        return yaml.load(f, Loader=SafeLoader)


def save_users(config: dict):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        yaml.dump(config, f, allow_unicode=True, default_flow_style=False)


# ── Pending users ──────────────────────────────────────────────────────────────

def load_pending() -> list:
    data = storage.download_file('pending_users.json')
    return json.loads(data.decode()) if data else []


def save_pending(pending: list):
    storage.upload_file(
        json.dumps(pending, indent=2, ensure_ascii=False).encode(),
        'pending_users.json',
    )


# ── Helpers ────────────────────────────────────────────────────────────────────

def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()


def username_exists(username: str) -> bool:
    config = load_users()
    return username.lower() in config['credentials']['usernames']


def username_pending(username: str) -> bool:
    return any(p['username'] == username.lower() for p in load_pending())


def is_admin(username: str) -> bool:
    """Vérifie si l'utilisateur a le rôle admin dans users.yaml."""
    try:
        config = load_users()
        user = config['credentials']['usernames'].get(username, {})
        return user.get('role') == 'admin'
    except Exception:
        return False


# ── Inscription ────────────────────────────────────────────────────────────────

def register_user(username: str, name: str, password: str) -> tuple[bool, str]:
    """Ajoute un utilisateur en attente de validation. Retourne (success, message)."""
    username = username.strip().lower()
    name     = name.strip()

    if not username or not name or not password:
        return False, 'Tous les champs sont obligatoires.'
    if len(username) < 3:
        return False, 'Nom d\'utilisateur trop court (3 caractères min).'
    if not username.isalnum():
        return False, 'Nom d\'utilisateur : lettres et chiffres uniquement.'
    if len(password) < 6:
        return False, 'Mot de passe trop court (6 caractères min).'
    if username_exists(username):
        return False, 'Ce nom d\'utilisateur est déjà utilisé.'
    if username_pending(username):
        return False, 'Une demande est déjà en attente pour ce nom d\'utilisateur.'

    pending = load_pending()
    pending.append({
        'username':     username,
        'name':         name,
        'password':     hash_password(password),
        'requested_at': datetime.now().strftime('%d/%m/%Y à %H:%M'),
    })
    save_pending(pending)
    return True, 'Demande envoyée. Un administrateur validera votre compte.'


# ── Validation admin ───────────────────────────────────────────────────────────

def approve_user(username: str) -> bool:
    """Approuve un utilisateur en attente et l'ajoute à users.yaml."""
    pending = load_pending()
    entry   = next((p for p in pending if p['username'] == username), None)
    if not entry:
        return False

    config = load_users()
    config['credentials']['usernames'][username] = {
        'name':     entry['name'],
        'password': entry['password'],
    }
    save_users(config)
    save_pending([p for p in pending if p['username'] != username])
    return True


def reject_user(username: str) -> bool:
    """Supprime un utilisateur en attente."""
    pending     = load_pending()
    new_pending = [p for p in pending if p['username'] != username]
    if len(new_pending) == len(pending):
        return False
    save_pending(new_pending)
    return True


def reset_password(username: str, new_password: str) -> tuple[bool, str]:
    """Réinitialise le mot de passe d'un utilisateur existant."""
    if len(new_password) < 6:
        return False, 'Mot de passe trop court (6 caractères min).'
    config = load_users()
    if username not in config['credentials']['usernames']:
        return False, 'Utilisateur introuvable.'
    config['credentials']['usernames'][username]['password'] = hash_password(new_password)
    save_users(config)
    return True, f'Mot de passe de {username} réinitialisé.'


def delete_user(username: str) -> bool:
    """Supprime un compte utilisateur de users.yaml."""
    config = load_users()
    if username not in config['credentials']['usernames']:
        return False
    del config['credentials']['usernames'][username]
    save_users(config)
    return True
