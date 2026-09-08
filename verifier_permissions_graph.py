#!/usr/bin/env python3
"""
verifier_permissions_graph.py -- Diagnostic en lecture seule

Verifie, avec les identifiants app-only deja utilises par ce repo (AZURE_TENANT_ID/
AZURE_CLIENT_ID/AZURE_CLIENT_SECRET), si l'application a les permissions Graph
necessaires pour un futur inventaire Office 365 (Fichiers = sites SharePoint,
Reunion = equipes Teams), en plus de Mail.Read deja confirme par les autres scripts
du repo. Ne modifie rien, n'ecrit nulle part.

Usage :
    python verifier_permissions_graph.py
"""
import os
import sys

import requests

GRAPH = "https://graph.microsoft.com/v1.0"


def get_app_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    resp = requests.post(url, data=data, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def check(label: str, url: str, token: str) -> None:
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers, timeout=30)
    if r.ok:
        n = len(r.json().get("value", []))
        print(f"OK   {label} -> HTTP {r.status_code}, {n} element(s) recu(s) sur cette page.")
    else:
        corps = (r.text or "")[:300]
        print(f"ECHEC {label} -> HTTP {r.status_code} : {corps}")


def main() -> int:
    tenant_id = os.environ.get("AZURE_TENANT_ID")
    client_id = os.environ.get("AZURE_CLIENT_ID")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET")
    if not all([tenant_id, client_id, client_secret]):
        print("Erreur : AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET requis.", file=sys.stderr)
        return 1

    print("Authentification Microsoft Graph (app-only)...")
    token = get_app_token(tenant_id, client_id, client_secret)
    print("OK authentification.\n")

    # Decodage (non verifie, juste lecture) du JWT pour afficher le Client ID / Tenant ID
    # de l'application -- utile pour la retrouver dans Entra ID (App registrations), car
    # ce ne sont pas des secrets (contrairement au client secret).
    import base64
    import json as _json
    payload_b64 = token.split(".")[1]
    payload_b64 += "=" * (-len(payload_b64) % 4)
    claims = _json.loads(base64.urlsafe_b64decode(payload_b64))
    print(f"Application (client) ID : {claims.get('appid')}")
    print(f"Nom de l'application    : {claims.get('app_displayname')}")
    print(f"Tenant ID                : {claims.get('tid')}\n")

    # Deja connu comme fonctionnel par les autres scripts du repo (Mail.Read), garde
    # ici comme point de reference.
    check("Mail.Read      (/users/{mailbox}/messages, 1 boite connue)",
          f"{GRAPH}/users/it@madavance.org/messages?$top=1&$select=id", token)

    check("User.Read.All  (/users, tenant entier)",
          f"{GRAPH}/users?$top=1&$select=id,displayName", token)

    check("Group.Read.All (/groups, tenant entier)",
          f"{GRAPH}/groups?$top=1&$select=id,displayName", token)

    check("Sites.Read.All (/sites/root, site racine du tenant)",
          f"{GRAPH}/sites/root?$select=id,displayName,webUrl", token)

    check("Sites.Read.All (/groups/{id}/sites/root, site d'un groupe connu -- IT MadAvance)",
          f"{GRAPH}/groups/21dbfa5a-2db8-4a91-b676-e4e786e9083f/sites/root?$select=id,webUrl", token)

    check("Files (drive Inventaire Office 365.xlsx, lecture)",
          f"{GRAPH}/drives/b!8D4xOy74F0-I2pDx1b5rX8HkGdhgNxpGpD3JvyEKMY4-rXDAT44VQ40NYtVFZG-V/items/01T5F36LFKRCAX3BS5YBHIZDCGCXHWSLR5?$select=id,name", token)

    return 0


if __name__ == "__main__":
    sys.exit(main())
