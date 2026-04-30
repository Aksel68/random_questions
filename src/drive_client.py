"""Cliente para interactuar con Google Drive API.
Maneja autenticación OAuth2, descarga de archivos .docx y subida de archivos generados.
"""

import csv
import io
import os
import tempfile
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive"]

CREDENTIALS_DIR = Path(__file__).parent.parent / "credentials"
TOKEN_PATH = CREDENTIALS_DIR / "token.json"

MIME_GDOC = "application/vnd.google-apps.document"
MIME_GSHEET = "application/vnd.google-apps.spreadsheet"
MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _find_client_secret() -> Path:
    matches = list(CREDENTIALS_DIR.glob("client_secret*.json"))
    if not matches:
        raise FileNotFoundError(
            f"No se encontró ningún archivo client_secret*.json en {CREDENTIALS_DIR}. "
            f"Seguí las instrucciones en credentials/README.md para obtener las credenciales."
        )
    return matches[0]


def _authenticate() -> Credentials:
    creds = None

    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            client_secret_path = _find_client_secret()
            flow = InstalledAppFlow.from_client_secrets_file(
                str(client_secret_path), SCOPES
            )
            creds = flow.run_local_server(port=0)

        TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(TOKEN_PATH, "w") as token_file:
            token_file.write(creds.to_json())

    return creds


def _get_service():
    creds = _authenticate()
    return build("drive", "v3", credentials=creds)


def _resolve_path(service, path: str) -> str:
    """Resuelve una ruta tipo 'Carpeta/Subcarpeta/archivo.docx' a un file ID de Drive."""
    parts = path.strip("/").split("/")
    parent_id = "root"

    for i, part in enumerate(parts):
        is_last = i == len(parts) - 1
        query = f"name = '{part}' and '{parent_id}' in parents and trashed = false"

        if not is_last:
            query += " and mimeType = 'application/vnd.google-apps.folder'"

        results = (
            service.files()
            .list(q=query, fields="files(id, name, mimeType)", spaces="drive")
            .execute()
        )
        files = results.get("files", [])

        if not files:
            raise FileNotFoundError(
                f"No se encontró '{part}' en la ruta '{path}'. "
                f"Verificá que el archivo exista en Google Drive."
            )

        parent_id = files[0]["id"]

    return parent_id


def get_file_mime_type(drive_path: str) -> str:
    """Consulta y retorna el mimeType de un archivo en Google Drive."""
    service = _get_service()
    file_id = _resolve_path(service, drive_path)

    file_metadata = service.files().get(fileId=file_id, fields="mimeType").execute()

    return file_metadata["mimeType"]


def download_file_as_docx(drive_path: str) -> str:
    """Descarga cualquier archivo de Drive como .docx local.

    Soporta:
    - Google Docs nativos (MIME_GDOC): exporta a .docx via export_media.
    - Archivos Word nativos (MIME_DOCX): descarga directamente via get_media.
    - Cualquier otro tipo: lanza ValueError.

    Retorna la ruta al archivo temporal .docx.
    """
    service = _get_service()
    file_id = _resolve_path(service, drive_path)

    file_metadata = service.files().get(fileId=file_id, fields="mimeType").execute()
    mime_type = file_metadata["mimeType"]

    temp_file = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)

    if mime_type == MIME_GDOC:
        request = service.files().export_media(
            fileId=file_id,
            mimeType=MIME_DOCX,
        )
    elif mime_type == MIME_DOCX:
        request = service.files().get_media(fileId=file_id)
    else:
        temp_file.close()
        os.unlink(temp_file.name)
        raise ValueError(
            f"Tipo de archivo no soportado para descarga como .docx: '{mime_type}'. "
            f"Se esperaba un Google Doc ({MIME_GDOC}) o un archivo Word ({MIME_DOCX})."
        )

    downloader = MediaIoBaseDownload(temp_file, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    temp_file.close()
    print(f"✅ Archivo descargado: {drive_path}")
    return temp_file.name


def download_answer_key_as_csv(drive_path: str) -> str:
    """Descarga la pauta de respuestas como string CSV.

    Soporta:
    - Google Sheets nativos (MIME_GSHEET): exporta a CSV via export_media.
    - Archivos Excel nativos (MIME_XLSX): descarga y convierte a CSV con openpyxl.
    - Cualquier otro tipo: lanza ValueError.

    Retorna el contenido CSV como string.
    """
    import openpyxl

    service = _get_service()
    file_id = _resolve_path(service, drive_path)

    file_metadata = service.files().get(fileId=file_id, fields="mimeType").execute()
    mime_type = file_metadata["mimeType"]

    buffer = io.BytesIO()

    if mime_type == MIME_GSHEET:
        request = service.files().export_media(
            fileId=file_id,
            mimeType="text/csv",
        )
        downloader = MediaIoBaseDownload(buffer, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        csv_content = buffer.getvalue().decode("utf-8")

    elif mime_type == MIME_XLSX:
        request = service.files().get_media(fileId=file_id)
        downloader = MediaIoBaseDownload(buffer, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        wb = openpyxl.load_workbook(io.BytesIO(buffer.getvalue()))
        ws = wb.active
        if ws is None:
            raise ValueError("El archivo Excel no contiene hojas activas.")
        output = io.StringIO()
        writer = csv.writer(output)
        for row in ws.iter_rows(values_only=True):
            writer.writerow([str(c) if c is not None else "" for c in row])
        csv_content = output.getvalue()

    else:
        raise ValueError(
            f"Tipo de archivo no soportado para descarga como CSV: '{mime_type}'. "
            f"Se esperaba un Google Sheet ({MIME_GSHEET}) o un archivo Excel ({MIME_XLSX})."
        )

    print(f"✅ Pauta descargada: {drive_path}")
    return csv_content


def download_docx(drive_path: str) -> str:
    """Descarga un archivo .docx desde Google Drive a un archivo temporal local."""
    service = _get_service()
    file_id = _resolve_path(service, drive_path)

    request = service.files().get_media(fileId=file_id)
    temp_file = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)

    downloader = MediaIoBaseDownload(temp_file, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    temp_file.close()
    print(f"✅ Archivo descargado: {drive_path}")
    return temp_file.name


def download_sheet_as_csv(drive_path: str) -> str:
    """Descarga un Google Sheet nativo y lo retorna como string CSV."""
    service = _get_service()
    file_id = _resolve_path(service, drive_path)

    request = service.files().export_media(
        fileId=file_id,
        mimeType="text/csv",
    )

    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    csv_content = buffer.getvalue().decode("utf-8")
    print(f"✅ Pauta descargada: {drive_path}")
    return csv_content


def upload_docx(local_path: str, drive_folder_path: str, filename: str) -> str:
    """Sube un archivo .docx a una carpeta específica en Google Drive."""
    service = _get_service()

    if drive_folder_path:
        parent_id = _resolve_path(service, drive_folder_path)
    else:
        parent_id = "root"

    file_metadata = {"name": filename, "parents": [parent_id]}
    media = MediaFileUpload(
        local_path,
        mimetype=MIME_DOCX,
    )

    file = (
        service.files()
        .create(body=file_metadata, media_body=media, fields="id")
        .execute()
    )

    print(f"✅ Subido: {filename}")
    return file.get("id")


def upload_xlsx(local_path: str, drive_folder_path: str, filename: str) -> str:
    """Sube un archivo .xlsx a una carpeta en Google Drive."""
    service = _get_service()

    if drive_folder_path:
        parent_id = _resolve_path(service, drive_folder_path)
    else:
        parent_id = "root"

    file_metadata = {"name": filename, "parents": [parent_id]}
    media = MediaFileUpload(
        local_path,
        mimetype=MIME_XLSX,
    )

    file = (
        service.files()
        .create(body=file_metadata, media_body=media, fields="id")
        .execute()
    )

    print(f"✅ Pauta subida: {filename}")
    return file.get("id")
