"""
OneDrive invoice fetcher via Microsoft Graph API.

Requires these environment variables:
  AZURE_TENANT_ID      — your Azure AD tenant ID
  AZURE_CLIENT_ID      — app registration client ID
  AZURE_CLIENT_SECRET  — app registration client secret
  ONEDRIVE_USER        — email/UPN of the mailbox owner (e.g. invoices@highlandcare.com)
  ONEDRIVE_FOLDER      — folder name inside that user's OneDrive (e.g. "Invoice Inbox")
"""

import os
import io
import tempfile
import requests

GRAPH_BASE = 'https://graph.microsoft.com/v1.0'
TOKEN_URL  = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token'


def _is_configured() -> bool:
    return all(os.environ.get(k) for k in (
        'AZURE_TENANT_ID', 'AZURE_CLIENT_ID',
        'AZURE_CLIENT_SECRET', 'ONEDRIVE_USER', 'ONEDRIVE_FOLDER',
    ))


def _get_token() -> str:
    """Obtain an access token using client credentials flow."""
    tenant = os.environ['AZURE_TENANT_ID']
    resp = requests.post(
        TOKEN_URL.format(tenant=tenant),
        data={
            'grant_type':    'client_credentials',
            'client_id':     os.environ['AZURE_CLIENT_ID'],
            'client_secret': os.environ['AZURE_CLIENT_SECRET'],
            'scope':         'https://graph.microsoft.com/.default',
        },
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json()['access_token']


def _headers(token: str) -> dict:
    return {'Authorization': f'Bearer {token}', 'Accept': 'application/json'}


def list_invoice_pdfs() -> list[dict]:
    """
    Return a list of PDF files in the configured OneDrive folder.
    Each item: {id, name, size, lastModified}
    Raises RuntimeError if not configured or on API error.
    """
    if not _is_configured():
        raise RuntimeError(
            'OneDrive not configured — set AZURE_TENANT_ID, AZURE_CLIENT_ID, '
            'AZURE_CLIENT_SECRET, ONEDRIVE_USER, and ONEDRIVE_FOLDER in Railway.'
        )

    token  = _get_token()
    user   = os.environ['ONEDRIVE_USER']
    folder = os.environ['ONEDRIVE_FOLDER']

    url = f'{GRAPH_BASE}/users/{user}/drive/root:/{folder}:/children'
    params = {'$select': 'id,name,size,lastModifiedDateTime,file', '$top': 100}

    items = []
    while url:
        resp = requests.get(url, headers=_headers(token), params=params, timeout=15)
        resp.raise_for_status()
        body = resp.json()
        for item in body.get('value', []):
            # Only PDFs
            mime = (item.get('file') or {}).get('mimeType', '')
            if item['name'].lower().endswith('.pdf') or mime == 'application/pdf':
                items.append({
                    'id':           item['id'],
                    'name':         item['name'],
                    'size':         item.get('size', 0),
                    'lastModified': item.get('lastModifiedDateTime', ''),
                })
        url = body.get('@odata.nextLink')
        params = None  # nextLink already has params baked in

    return items


def download_pdf(item_id: str) -> str:
    """
    Download a OneDrive file by item ID to a temp file.
    Returns the temp file path (caller must delete it).
    """
    token = _get_token()
    user  = os.environ['ONEDRIVE_USER']

    # Get download URL
    url  = f'{GRAPH_BASE}/users/{user}/drive/items/{item_id}'
    resp = requests.get(url, headers=_headers(token), timeout=15)
    resp.raise_for_status()
    download_url = resp.json().get('@microsoft.graph.downloadUrl')

    if not download_url:
        # Fallback: /content endpoint
        download_url = f'{GRAPH_BASE}/users/{user}/drive/items/{item_id}/content'

    resp = requests.get(download_url, timeout=30, stream=True)
    resp.raise_for_status()

    fd, path = tempfile.mkstemp(suffix='.pdf')
    with os.fdopen(fd, 'wb') as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)

    return path
