#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Met a jour la feuille "Email" de Inventaire Office 365.xlsx (classeur SharePoint
partage, meme fichier/dossier que Composantes Office 365.xlsx) avec la liste des
comptes Microsoft 365 actifs et licencies -- visibilite tenant entiere (via une
application Graph app-only), independante des acces personnels de l'executant.

Methode (validee le 07/09/2026, reprise ici en code reutilisable) :
    GET /v3/users -> filtre userType eq 'Member' and accountEnabled eq true, puis
    ne garder que les comptes avec au moins une licence assignee (assignedLicenses
    non vide). Pas de distinction individuelle/partagee (Graph ne l'expose pas de
    facon fiable -- concept Exchange recipientType, pas Graph).

PREREQUIS (bloquant tant que non fait) : cette application doit avoir la permission
d'application Graph "User.Read.All" avec consentement admin -- verifie le
08/09/2026 via diagnostics/verifier_permissions_graph.py : ABSENTE a cette date
(403 Authorization_RequestDenied). Sans elle, ce script echoue des l'appel a
/users. A ajouter dans Azure AD > App registrations > (cette app) > API
permissions > Add > Microsoft Graph > Application permissions > User.Read.All,
puis "Grant admin consent".

Variables d'environnement attendues (memes secrets que les autres scripts du
depot, deja utilises pour Mail.Read) :
    AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET

Usage :
    python office365_email_list.py
"""

from __future__ import annotations

import io
import os
import re
import sys

import requests
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

GRAPH_API = "https://graph.microsoft.com/v1.0"
TIMEOUT = 120

# Site SharePoint ITMadAvance, meme drive que Inventaire mWater.xlsx / Composantes
# Office 365.xlsx (dossier "Office 365").
SHAREPOINT_DRIVE_ID = "b!8D4xOy74F0-I2pDx1b5rX8HkGdhgNxpGpD3JvyEKMY4-rXDAT44VQ40NYtVFZG-V"
INVENTAIRE_ITEM_ID = "01T5F36LFKRCAX3BS5YBHIZDCGCXHWSLR5"  # Inventaire Office 365.xlsx
SHEET = "Email"


def raise_verbose(resp: requests.Response, contexte: str) -> None:
    if resp.ok:
        return
    corps = (resp.text or "")[:1500]
    raise RuntimeError(f"{contexte} : HTTP {resp.status_code} {resp.reason}\n{corps}")


def authentifier_graph(tenant_id: str, client_id: str, client_secret: str) -> str:
    print("Authentification Microsoft Graph (app-only)…", flush=True)
    r = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=TIMEOUT,
    )
    raise_verbose(r, "Authentification Microsoft Graph")
    token = r.json().get("access_token")
    if not token:
        raise RuntimeError("Pas de access_token dans la réponse Graph.")
    print("  → authentification réussie.", flush=True)
    return token


def graph_get_all(url: str, token: str) -> list[dict]:
    items: list[dict] = []
    headers = {"Authorization": f"Bearer {token}"}
    while url:
        r = requests.get(url, headers=headers, timeout=TIMEOUT)
        raise_verbose(r, f"GET {url}")
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def fetch_email_rows(token: str) -> list[tuple[str, str]]:
    print("Récupération des comptes Email (Graph /users)…", flush=True)
    url = (
        f"{GRAPH_API}/users?$filter=userType eq 'Member' and accountEnabled eq true"
        "&$select=id,displayName,mail,assignedLicenses&$top=999"
    )
    users = graph_get_all(url, token)
    licencies = [u for u in users if u.get("assignedLicenses")]
    licencies.sort(key=lambda u: (u.get("displayName") or "").lower())
    print(
        f"  → {len(licencies)} compte(s) avec licence sur {len(users)} Member actif(s).",
        flush=True,
    )
    return [(u.get("displayName") or "", u.get("mail") or "") for u in licencies]


def telecharger_maitre_existant(token: str) -> bytes:
    print("Téléchargement de Inventaire Office 365.xlsx (classeur partagé)…", flush=True)
    r = requests.get(
        f"{GRAPH_API}/drives/{SHAREPOINT_DRIVE_ID}/items/{INVENTAIRE_ITEM_ID}/content",
        headers={"Authorization": f"Bearer {token}"},
        timeout=TIMEOUT,
    )
    raise_verbose(r, "Téléchargement de Inventaire Office 365.xlsx")
    print(f"  → fichier récupéré ({len(r.content)} octets).", flush=True)
    return r.content


def uploader_maitre(token: str, contenu: bytes) -> None:
    print("Envoi de Inventaire Office 365.xlsx mis à jour…", flush=True)
    r = requests.put(
        f"{GRAPH_API}/drives/{SHAREPOINT_DRIVE_ID}/items/{INVENTAIRE_ITEM_ID}/content",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
        },
        data=contenu,
        timeout=TIMEOUT,
    )
    raise_verbose(r, "Envoi de Inventaire Office 365.xlsx")
    print("  → envoyé.", flush=True)


def safe_table_name(sheet_name: str) -> str:
    s = re.sub(r"[^A-Za-z0-9_]", "_", sheet_name)
    if not s or s[0].isdigit():
        s = "T_" + s
    return "Tbl_" + s


def add_header_table(ws, style_name: str = "TableStyleLight9") -> None:
    max_col, max_row = ws.max_column, ws.max_row
    if max_col < 1 or max_row < 1:
        return
    if ws.tables:
        for name in list(ws.tables.keys()):
            del ws.tables[name]
    ref = f"A1:{get_column_letter(max_col)}{max_row}"
    tbl = Table(displayName=safe_table_name(ws.title), ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name=style_name, showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(tbl)


def ecrire_feuille(wb, sheet: str, headers: tuple[str, ...], rows: list[tuple]) -> None:
    ws = wb[sheet]
    zone_max = max(ws.max_row, len(rows) + 1)
    for row in ws.iter_rows(min_row=1, max_row=zone_max, max_col=max(len(headers), ws.max_column)):
        for cell in row:
            cell.value = None
    for col, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=h)
    for i, values in enumerate(rows, start=1):
        ws.cell(row=i + 1, column=1, value=i)
        for col, v in enumerate(values, start=2):
            ws.cell(row=i + 1, column=col, value=v)
    if ws.max_row > len(rows) + 1:
        ws.delete_rows(len(rows) + 2, ws.max_row - (len(rows) + 1))
    add_header_table(ws, style_name="TableStyleLight9")


def main() -> int:
    tenant_id = os.environ.get("AZURE_TENANT_ID")
    client_id = os.environ.get("AZURE_CLIENT_ID")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET")
    if not all([tenant_id, client_id, client_secret]):
        print("Erreur : AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET requis.", file=sys.stderr)
        return 1

    token = authentifier_graph(tenant_id, client_id, client_secret)
    rows = fetch_email_rows(token)

    contenu = telecharger_maitre_existant(token)
    wb = load_workbook(io.BytesIO(contenu))
    ecrire_feuille(wb, SHEET, ("N°", "Nom", "Adresse mail"), rows)

    sortie = io.BytesIO()
    wb.save(sortie)
    uploader_maitre(token, sortie.getvalue())

    print(f"OK — {len(rows)} compte(s) écrit(s) dans la feuille '{SHEET}'.", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
