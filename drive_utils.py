"""
Utilitaires Google Drive — lecture/écriture des fichiers de l'app.
"""

import io
import streamlit as st
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/drive']
FOLDER_NAME = 'Picking List Generator'


@st.cache_resource
def get_service():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets['gcp_service_account'], scopes=SCOPES
    )
    return build('drive', 'v3', credentials=creds)


@st.cache_data(ttl=60)
def get_folder_id(folder_name=FOLDER_NAME):
    service = get_service()
    results = service.files().list(
        q=f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields='files(id, name)',
    ).execute()
    files = results.get('files', [])
    if not files:
        raise FileNotFoundError(f"Dossier Google Drive '{folder_name}' introuvable.")
    return files[0]['id']


def get_subfolder_id(subfolder_name, parent_folder_id=None):
    """Retourne l'id d'un sous-dossier, le crée si nécessaire."""
    service = get_service()
    if parent_folder_id is None:
        parent_folder_id = get_folder_id()
    results = service.files().list(
        q=(f"name='{subfolder_name}' and mimeType='application/vnd.google-apps.folder' "
           f"and '{parent_folder_id}' in parents and trashed=false"),
        fields='files(id)',
    ).execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']
    # Créer le sous-dossier
    meta = {
        'name': subfolder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_folder_id],
    }
    folder = service.files().create(body=meta, fields='id').execute()
    return folder['id']


def download_file(file_name, folder_id=None) -> bytes | None:
    """Télécharge un fichier depuis Drive. Retourne les bytes ou None si absent."""
    service = get_service()
    if folder_id is None:
        folder_id = get_folder_id()
    results = service.files().list(
        q=f"name='{file_name}' and '{folder_id}' in parents and trashed=false",
        fields='files(id, name)',
    ).execute()
    files = results.get('files', [])
    if not files:
        return None
    file_id = files[0]['id']
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


def upload_file(file_bytes: bytes, file_name: str, folder_id=None) -> str:
    """Upload ou met à jour un fichier sur Drive. Retourne l'id du fichier."""
    service = get_service()
    if folder_id is None:
        folder_id = get_folder_id()

    mime = _guess_mime(file_name)
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime, resumable=True)

    # Chercher si le fichier existe déjà
    results = service.files().list(
        q=f"name='{file_name}' and '{folder_id}' in parents and trashed=false",
        fields='files(id)',
    ).execute()
    existing = results.get('files', [])

    if existing:
        file_id = existing[0]['id']
        service.files().update(fileId=file_id, media_body=media).execute()
        return file_id
    else:
        meta = {'name': file_name, 'parents': [folder_id]}
        f = service.files().create(body=meta, media_body=media, fields='id').execute()
        return f['id']


def list_files(folder_id=None, pattern=None) -> list[dict]:
    """Liste les fichiers d'un dossier Drive. Optionnel: filtrer par début de nom."""
    service = get_service()
    if folder_id is None:
        folder_id = get_folder_id()
    q = f"'{folder_id}' in parents and trashed=false"
    if pattern:
        q += f" and name contains '{pattern}'"
    results = service.files().list(
        q=q,
        fields='files(id, name, modifiedTime, size)',
        orderBy='modifiedTime desc',
    ).execute()
    return results.get('files', [])


def _guess_mime(file_name: str) -> str:
    ext = file_name.rsplit('.', 1)[-1].lower() if '.' in file_name else ''
    return {
        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'xls':  'application/vnd.ms-excel',
        'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'pdf':  'application/pdf',
        'json': 'application/json',
    }.get(ext, 'application/octet-stream')
