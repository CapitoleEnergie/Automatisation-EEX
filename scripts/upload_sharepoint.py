"""
Upload eex_prix.xlsx vers SharePoint via Microsoft Graph API.
Appelé depuis GitHub Actions après la collecte EEX.

Variables d'environnement requises (GitHub Secrets) :
  SP_TENANT_ID      - Azure AD Tenant ID
  SP_CLIENT_ID      - Azure AD Application (client) ID
  SP_CLIENT_SECRET  - Azure AD Client Secret Value
  SP_SITE_ID        - SharePoint Site ID (obtenu via Graph Explorer)
  SP_DRIVE_ID       - SharePoint Drive ID (obtenu via Graph Explorer)
"""

import os
import sys
import requests
from datetime import date

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

EXCEL_PATH    = os.path.join(os.path.dirname(__file__), "..", "data", "eex_prix.xlsx")
REMOTE_FOLDER = "PARTAGE/Team/16. Pricing/9. Point marché journalier/2 - PRIX CLOTURE/Test EEX"
REMOTE_FILE   = "eex_prix.xlsx"

GRAPH_AUTH_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


# ---------------------------------------------------------------------------
# Auth — récupère un token OAuth2 via Client Credentials
# ---------------------------------------------------------------------------

def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    url = GRAPH_AUTH_URL.format(tenant_id=tenant_id)
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     client_id,
        "client_secret": client_secret,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=30)
    resp.raise_for_status()
    token = resp.json().get("access_token")
    if not token:
        raise RuntimeError(f"Token non reçu : {resp.json()}")
    return token


# ---------------------------------------------------------------------------
# Upload — PUT le fichier dans le dossier SharePoint cible
# ---------------------------------------------------------------------------

def upload_to_sharepoint(token: str, site_id: str, drive_id: str) -> dict:
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Fichier introuvable : {EXCEL_PATH}")

    # Chemin distant : dossier/fichier
    remote_path = f"{REMOTE_FOLDER}/{REMOTE_FILE}"

    # Endpoint Graph pour upload (create or replace)
    url = f"{GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{remote_path}:/content"

    with open(EXCEL_PATH, "rb") as f:
        file_bytes = f.read()

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }

    resp = requests.put(url, headers=headers, data=file_bytes, timeout=60)
    resp.raise_for_status()
    return resp.json()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    today_str = date.today().strftime("%Y-%m-%d")
    print(f"\n=== Upload SharePoint — {today_str} ===")

    # Lecture des secrets depuis l'environnement
    tenant_id     = os.environ.get("SP_TENANT_ID")
    client_id     = os.environ.get("SP_CLIENT_ID")
    client_secret = os.environ.get("SP_CLIENT_SECRET")
    site_id       = os.environ.get("SP_SITE_ID")
    drive_id      = os.environ.get("SP_DRIVE_ID")

    missing = [k for k, v in {
        "SP_TENANT_ID":     tenant_id,
        "SP_CLIENT_ID":     client_id,
        "SP_CLIENT_SECRET": client_secret,
        "SP_SITE_ID":       site_id,
        "SP_DRIVE_ID":      drive_id,
    }.items() if not v]

    if missing:
        print(f"  ERREUR — Variables manquantes : {', '.join(missing)}")
        sys.exit(1)

    print("  Authentification Azure AD...")
    token = get_access_token(tenant_id, client_id, client_secret)
    print("  Token obtenu.")

    print(f"  Upload vers SharePoint : {REMOTE_FOLDER}/{REMOTE_FILE}")
    result = upload_to_sharepoint(token, site_id, drive_id)

    # Affichage du résultat
    web_url       = result.get("webUrl", "—")
    size_kb       = result.get("size", 0) // 1024
    last_modified = result.get("lastModifiedDateTime", "—")

    print(f"  Succès.")
    print(f"  URL      : {web_url}")
    print(f"  Taille   : {size_kb} Ko")
    print(f"  Modifié  : {last_modified}")


if __name__ == "__main__":
    main()
