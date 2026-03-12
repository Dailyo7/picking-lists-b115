"""
Utilitaires de stockage local — remplace drive_utils.py sur VPS.
Interface identique à drive_utils pour minimiser les changements dans l'app.
"""

from pathlib import Path

DATA_DIR = Path('/opt/picking-lists/data')


def get_folder_id(folder_name=None) -> str:
    """Retourne le chemin du dossier data principal."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    return str(DATA_DIR)


def get_subfolder_id(subfolder_name: str, parent_folder_id: str = None) -> str:
    """Retourne le chemin d'un sous-dossier, le crée si nécessaire."""
    parent = Path(parent_folder_id) if parent_folder_id else DATA_DIR
    subdir = parent / subfolder_name
    subdir.mkdir(parents=True, exist_ok=True)
    return str(subdir)


def download_file(file_name: str, folder_id: str = None) -> bytes | None:
    """Lit un fichier depuis le dossier local. Retourne bytes ou None si absent."""
    folder = Path(folder_id) if folder_id else DATA_DIR
    path = folder / file_name
    return path.read_bytes() if path.exists() else None


def upload_file(file_bytes: bytes, file_name: str, folder_id: str = None) -> str:
    """Écrit un fichier dans le dossier local. Retourne le chemin."""
    folder = Path(folder_id) if folder_id else DATA_DIR
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / file_name
    path.write_bytes(file_bytes)
    return str(path)


def list_files(folder_id: str = None, pattern: str = None) -> list[dict]:
    """Liste les fichiers d'un dossier local."""
    folder = Path(folder_id) if folder_id else DATA_DIR
    if not folder.exists():
        return []
    files = sorted([f for f in folder.iterdir() if f.is_file()])
    if pattern:
        files = [f for f in files if pattern in f.name]
    return [{'id': str(f), 'name': f.name} for f in files]


def delete_file(file_name: str, folder_id: str = None) -> bool:
    """Supprime un fichier du dossier local. Retourne True si supprimé."""
    folder = Path(folder_id) if folder_id else DATA_DIR
    path = folder / file_name
    if path.exists():
        path.unlink()
        return True
    return False


def get_service():
    """Compatibilité drive_utils — non utilisé sur VPS."""
    return None
