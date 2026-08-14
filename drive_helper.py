"""
Google Drive optional integration.

Lets the app pull campus photos (and, optionally, .docx templates) from a
shared Google Drive folder instead of the local images/ and templates/
folders. This is entirely opt-in: if the required secrets are not set,
`drive_enabled()` returns False and every other function in this module
is simply never called by app.py.

Required Streamlit secrets (.streamlit/secrets.toml or the Cloud "Secrets" UI):

    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
    client_email = "xxxx@xxxx.iam.gserviceaccount.com"
    client_id = "..."
    token_uri = "https://oauth2.googleapis.com/token"

    drive_images_folder_id = "1AbCdEfGhIjKlMnOpQrStUvWxYz"
    drive_templates_folder_id = "1AbCdEfGhIjKlMnOpQrStUvWxYz"   # optional

See DEPLOY.md for the full step-by-step setup guide.
"""
from __future__ import annotations
import os
import streamlit as st

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


def drive_enabled() -> bool:
    """True only if a service account + at least the images folder are configured."""
    try:
        has_account = "gcp_service_account" in st.secrets
        has_folder = bool(st.secrets.get("drive_images_folder_id"))
        return has_account and has_folder
    except Exception:
        return False


def templates_from_drive_enabled() -> bool:
    try:
        return drive_enabled() and bool(st.secrets.get("drive_templates_folder_id"))
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def get_drive_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    info = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


@st.cache_data(ttl=600, show_spinner=False)
def list_drive_files(folder_id: str) -> list[dict]:
    """List all (non-trashed) files directly inside a Drive folder."""
    service = get_drive_service()
    files: list[dict] = []
    page_token = None
    query = f"'{folder_id}' in parents and trashed = false"
    while True:
        resp = service.files().list(
            q=query,
            spaces="drive",
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=page_token,
            pageSize=1000,
        ).execute()
        files.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return files


def download_drive_file(file_id: str, dest_path: str) -> str:
    from googleapiclient.http import MediaIoBaseDownload

    service = get_drive_service()
    request = service.files().get_media(fileId=file_id)
    os.makedirs(os.path.dirname(os.path.abspath(dest_path)), exist_ok=True)
    with open(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return dest_path


def _base_name(fname: str) -> str:
    """Same normalization rule as generator.get_available_images() so the two
    sources (local folder / Drive folder) match university names consistently."""
    name = fname.split("pixta_")[0].split("_pixta")[0].strip().rstrip("_").strip()
    return name


def find_drive_image(university_name: str, folder_id: str) -> dict | None:
    """Return the Drive file dict ({'id','name',...}) best matching a university name."""
    if not university_name or not folder_id:
        return None
    files = list_drive_files(folder_id)
    images = [f for f in files if f["name"].lower().endswith((".jpg", ".jpeg", ".png"))]

    # Prefer an exact base-name match, then a "_M" (medium) variant, then substrings.
    exact = [f for f in images if _base_name(f["name"]) == university_name]
    if exact:
        for f in exact:
            if "_M." in f["name"] or " M." in f["name"]:
                return f
        return exact[0]

    for f in images:
        base = _base_name(f["name"])
        if university_name in base or base in university_name:
            return f
    return None


def find_all_drive_images(university_name: str, folder_id: str) -> list[dict]:
    """Return every Drive photo matching a university name, for a picker UI."""
    if not university_name or not folder_id:
        return []
    files = list_drive_files(folder_id)
    images = [f for f in files if f["name"].lower().endswith((".jpg", ".jpeg", ".png"))]

    exact = [f for f in images if _base_name(f["name"]) == university_name]
    if exact:
        return exact
    return [
        f for f in images
        if university_name in _base_name(f["name"]) or _base_name(f["name"]) in university_name
    ]


def download_drive_file_bytes(file_id: str) -> bytes:
    """Fetch a Drive file's full bytes without writing to disk (for quick previews)."""
    import io
    from googleapiclient.http import MediaIoBaseDownload

    service = get_drive_service()
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


def find_drive_template(tmpl_key: int, folder_id: str) -> dict | None:
    """Look for template_<N>.docx inside the Drive templates folder."""
    if not folder_id:
        return None
    target = f"template_{tmpl_key}.docx"
    for f in list_drive_files(folder_id):
        if f["name"] == target:
            return f
    return None


def list_drive_image_names(folder_id: str) -> list[str]:
    files = list_drive_files(folder_id)
    names = sorted({
        _base_name(f["name"]) for f in files
        if f["name"].lower().endswith((".jpg", ".jpeg", ".png")) and _base_name(f["name"])
    })
    return names
