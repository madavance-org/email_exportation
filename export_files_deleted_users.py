#!/usr/bin/env python3
"""
export_files_deleted_users.py -- "Extraction utilisateurs partis"

Contexte (13/08/2026) : eddy.rajaonarivony@madavance.org et fara@madavance.org
sont deux personnes dont la propre boite mail n'aide plus (celle d'Eddy est
vide, celle de Fara etait sans licence -- voir extract_baobab.py). Pour
retrouver leurs echanges, on scanne la boite de CHAQUE AUTRE membre du staff
MadAvance (celles qui ont une vraie licence, donc une boite Exchange
provisionnee) a la recherche de messages ou l'une des deux personnes
apparait comme EXPEDITEUR OU DESTINATAIRE (a/cc) -- pas de filtre mot-cle,
comme extract_baobab.py/export_attachments.py : toutes les pieces jointes de
ces messages sont extraites.

Resilience (13/08/2026, suite a l'echec de "Extraction pieces jointes -
Equipe" sur fara@madavance.org avant qu'elle ait une licence) : si une boite
de la liste renvoie une erreur (mailbox non provisionnee, inactive, etc.),
on logue et on passe a la boite suivante plutot que de faire echouer tout le
run.

Reprise (13/08/2026, meme principe que extract_equipe.py) : avec 78 boites a
scanner, un run peut etre coupe avant la fin (limite GitHub Actions, etc.).
Un log JSON par cible ("mailbox::message_id" par cle) est depose dans le
sous-dossier SharePoint de CETTE cible, mis a jour tous les LOG_SAVE_EVERY
messages traites -- le run suivant saute directement les messages deja vus.

Rangement (14/08/2026) : les resultats sont deposes dans le sous-dossier de
la personne concernee (eddy.rajaonarivony_madavance.org, fara_madavance.org
-- meme nom que ceux crees par extract_equipe.py, donc fusion dans les memes
dossiers) plutot que dans une arborescence annee/mois a part. Un message qui
concerne les deux cibles a la fois (rare) est deplace dans les deux dossiers.
Le nom du log ("processed_log_deleted_users.json") est volontairement
different de celui d'extract_equipe.py ("processed_log.json") pour ne pas
ecraser son propre suivi de reprise : les deux scripts partagent le dossier
mais pas leur fichier de log.

Cibles par defaut : la liste des comptes staff avec licence (~78, voir
DEFAULT_SENDERS) -- pas les comptes invites externes ni les comptes sans
licence (constate le 13/08/2026 : renvoient 404 MailboxNotEnabledForRESTAPI).
Modifiable via --senders.

Authentification identique aux autres scripts du repo (BackupOffice365,
app-only) :
    AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET

Destination SharePoint :
    SHAREPOINT_FOLDER_LINK_DELETED_USERS

Usage :
    python export_files_deleted_users.py
"""

import argparse
import base64
import hashlib
import json
import os
import re
import sys
import time
import traceback
import unicodedata
from datetime import date
from pathlib import Path
from urllib.parse import quote

import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

TARGET_ADDRESSES = {"eddy.rajaonarivony@madavance.org", "fara@madavance.org"}

# Log de reprise, depose dans le sous-dossier SharePoint de CHAQUE cible (un
# sous-dossier + un log par personne, meme principe que extract_equipe.py).
# Cle composite "mailbox::id" car plusieurs boites source (parmi les ~78
# scannees) peuvent chacune contenir un message concernant la meme cible.
# Nom distinct de celui d'extract_equipe.py ("processed_log.json") pour ne
# pas ecraser son log en partageant le meme dossier.
LOG_FILENAME = "processed_log_deleted_users.json"
LOG_SAVE_EVERY = 10

# Comptes staff MadAvance avec une licence Microsoft 365 (donc une vraie
# boite Exchange), constate le 13/08/2026 via Graph (userType=Member ET
# assignedLicenses non vide). Exclut les comptes invites externes (#EXT#) et
# les comptes staff sans licence (pas de boite du tout, cf. fara@madavance.org
# avant l'ajout de sa licence).
DEFAULT_SENDERS = [
    "ChauffeurMDVC001@madavance.org",
    "ChauffeurMDVC002@madavance.org",
    "ChauffeurMDVC003@madavance.org",
    "ChauffeurMDVC004@madavance.org",
    "ChauffeurMDVC005@madavance.org",
    "Frantzel.andrianalison@madavance.org",
    "Hardy.onitra@madavance.org",
    "Herman@madavance.org",
    "Marie.bristol@madavance.org",
    "admin@madavance.org",
    "adriaan.mol@madavance.org",
    "alidah@madavance.org",
    "amede.rafidimanantsoa@madavance.org",
    "andriniaina@madavance.org",
    "angelo.nahavitatsara@madavance.org",
    "antonio@madavance.org",
    "carole.totozafy@madavanceNGO.onmicrosoft.com",
    "coddy@madavance.org",
    "coddy@madavanceNGO.onmicrosoft.com",
    "comptable@madavance.org",
    "confidentiel_grievance@madavance.org",
    "contact@madavance.org",
    "deichmanfundation@madavance.org",
    "diamondra@madavance.org",
    "dieudonne.razafimahatratra@madavance.org",
    "domoina@madavance.org",
    "don.madavance@madavance.org",
    "dorothee.velonjara@madavance.org",
    "eddy.rajaonarivony@madavance.org",
    "fanja@madavance.org",
    "fara.fara@madavance.org",
    "fara@madavance.org",
    "fideline.fanomezantsoa@madavance.org",
    "florent.ravelotiana@madavance.org",
    "florent@madavance.org",
    "gaetan@madavance.org",
    "hery@madavance.org",
    "holisoa.raharijaona@madavance.org",
    "info@madavance.org",
    "it@madavance.org",
    "karelle@madavance.org",
    "kevin@madavance.org",
    "lalao.patrick@madavance.org",
    "lea@madavance.org",
    "lucienne.rasamoelina@madavance.org",
    "madavance@madavance.org",
    "mamywilliam@madavance.org",
    "massou-franzza.leandera@madavance.org",
    "merci@madavance.org",
    "miando.razafimanantsoa@madavance.org",
    "mickael.consultant@madavance.org",
    "narindra.mauricilla@madavance.org",
    "nary.ramanarivo@madavance.org",
    "nekena@madavance.org",
    "nicolas.blasquez@madavance.org",
    "noreply@madavanceNGO.onmicrosoft.com",
    "ntsoa@madavanceNGO.onmicrosoft.com",
    "olivia@madavance.org",
    "patrick.solofondalambo@madavance.org",
    "priscilla.fung@madavance.org",
    "rakitrynyavo@madavance.org",
    "rasoamiafara.reine@madavance.org",
    "rindra.andriamahefa@madavance.org",
    "sandra.rasoamampionona@madavance.org",
    "sarobidy.fanomezantsoa@madavance.org",
    "starlinkftu@madavance.org",
    "starlinkmaroantsetra@madavance.org",
    "tanjona@madavance.org",
    "theophile@madavance.org",
    "tiana@madavance.org",
    "tonga.ghislain@madavance.org",
    "toto.arivelona@madavance.org",
    "totondalahy.arivelona@madavance.org",
    "vania.nomenimanjaka@madavance.org",
    "vavitiana@madavance.org",
    "vehiclerequest@madavance.org",
    "volatiana.t@madavance.org",
    "william@madavance.org",
]

CHUNK_SIZE = 320 * 1024 * 30
TOKEN_REFRESH_MARGIN_SECONDS = 120
TRANSIENT_STATUS_CODES = {429, 500, 502, 503, 504}
MAX_TRANSIENT_RETRIES = 5
MAX_RETRY_DELAY_SECONDS = 60


def raise_for_status_verbose(resp: requests.Response) -> None:
    if resp.status_code >= 400:
        raise RuntimeError(f"Erreur HTTP {resp.status_code} sur {resp.request.method} {resp.url}\n{resp.text}")


def _compute_retry_delay(resp: requests.Response, attempt: int) -> float:
    retry_after = resp.headers.get("Retry-After")
    if retry_after:
        try:
            return float(retry_after)
        except ValueError:
            pass
    try:
        retry_seconds = resp.json().get("error", {}).get("retryAfterSeconds")
        if retry_seconds:
            return float(retry_seconds)
    except Exception:
        pass
    return min(2 ** attempt, MAX_RETRY_DELAY_SECONDS)


def request_with_retry(method: str, url: str, **kwargs) -> requests.Response:
    attempt = 0
    while True:
        try:
            resp = requests.request(method, url, **kwargs)
        except requests.exceptions.RequestException as exc:
            if attempt >= MAX_TRANSIENT_RETRIES:
                raise
            delay = min(2 ** attempt, MAX_RETRY_DELAY_SECONDS)
            print(f"    (Erreur reseau '{exc}', nouvelle tentative dans {delay:.0f}s...)")
            time.sleep(delay)
            attempt += 1
            continue

        if resp.status_code not in TRANSIENT_STATUS_CODES or attempt >= MAX_TRANSIENT_RETRIES:
            return resp
        delay = _compute_retry_delay(resp, attempt)
        print(f"    (Graph a renvoye {resp.status_code}, nouvelle tentative dans {delay:.0f}s...)")
        time.sleep(delay)
        attempt += 1


class GraphSession:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self._token = None
        self._expires_at = 0

    def _fetch_token(self) -> None:
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        resp = requests.post(url, data=data, timeout=30)
        raise_for_status_verbose(resp)
        payload = resp.json()
        self._token = payload["access_token"]
        self._expires_at = time.time() + int(payload.get("expires_in", 3599))

    def _token_value(self, force_refresh: bool = False) -> str:
        if force_refresh or self._token is None or time.time() >= self._expires_at - TOKEN_REFRESH_MARGIN_SECONDS:
            self._fetch_token()
        return self._token

    def ensure_ready(self) -> None:
        self._token_value()

    def request(self, method: str, url: str, **kwargs) -> requests.Response:
        headers = dict(kwargs.pop("headers", None) or {})
        headers["Authorization"] = f"Bearer {self._token_value()}"
        resp = request_with_retry(method, url, headers=headers, **kwargs)
        if resp.status_code == 401:
            headers["Authorization"] = f"Bearer {self._token_value(force_refresh=True)}"
            resp = request_with_retry(method, url, headers=headers, **kwargs)
        return resp

    def get(self, url: str, **kwargs) -> requests.Response:
        return self.request("GET", url, **kwargs)

    def post(self, url: str, **kwargs) -> requests.Response:
        return self.request("POST", url, **kwargs)

    def put(self, url: str, **kwargs) -> requests.Response:
        return self.request("PUT", url, **kwargs)


def list_all_messages_with_attachments(session: GraphSession, mailbox: str) -> list[dict]:
    url = f"{GRAPH_BASE}/users/{mailbox}/messages"
    params = {
        "$filter": "hasAttachments eq true",
        "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments",
        "$top": "100",
    }
    messages = []
    while url:
        resp = session.get(url, params=params, timeout=30)
        raise_for_status_verbose(resp)
        payload = resp.json()
        messages.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
        params = None
    return messages


def list_attachments(session: GraphSession, mailbox: str, message_id: str) -> list[dict]:
    resp = session.get(f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}/attachments", timeout=30)
    raise_for_status_verbose(resp)
    return resp.json().get("value", [])


def download_attachment_bytes(session: GraphSession, mailbox: str, message_id: str, attachment_id: str) -> bytes:
    resp = session.get(f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}/attachments/{attachment_id}/$value", timeout=120)
    raise_for_status_verbose(resp)
    return resp.content


def slugify(value: str, max_len: int = 60) -> str:
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    value = re.sub(r"[^\w\-. ]", "_", value).strip()
    value = re.sub(r"\s+", "_", value)
    return value[:max_len] or "sans_nom"


def build_filename(received: str, subject: str, original_name: str, max_stem_len: int = 60) -> str:
    base_name, dot, ext = original_name.rpartition(".")
    stem_source = base_name if dot else original_name
    stem = slugify(stem_source, max_stem_len)
    prefix = f"{received}_{slugify(subject, 40)}_{stem}"
    if dot and ext:
        ext_clean = re.sub(r"[^A-Za-z0-9]", "", ext)[:10]
        if ext_clean:
            return f"{prefix}.{ext_clean.lower()}"
    return prefix


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    i = 2
    while True:
        candidate = path.with_name(f"{stem}_{i}{suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def encode_sharing_url(url: str) -> str:
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return f"u!{b64}"


def resolve_share_link(session: GraphSession, share_url: str) -> tuple[str, str]:
    resp = session.get(f"{GRAPH_BASE}/shares/{encode_sharing_url(share_url)}/driveItem", timeout=30)
    raise_for_status_verbose(resp)
    item = resp.json()
    return item["parentReference"]["driveId"], item["id"]


def get_or_create_child_folder(session: GraphSession, drive_id: str, parent_item_id: str, name: str) -> str:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent_item_id}/children"
    resp = session.get(url, params={"$select": "id,name,folder"}, timeout=30)
    raise_for_status_verbose(resp)
    for item in resp.json().get("value", []):
        if item.get("folder") is not None and item.get("name") == name:
            return item["id"]
    resp = session.post(
        url,
        headers={"Content-Type": "application/json"},
        json={"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"},
        timeout=30,
    )
    raise_for_status_verbose(resp)
    return resp.json()["id"]


def get_year_month_folder(session, drive_id, base_folder_id, year, month, cache):
    key = (year, month)
    if key in cache:
        return cache[key]
    year_key = (year, "")
    year_folder_id = cache.get(year_key)
    if year_folder_id is None:
        year_folder_id = get_or_create_child_folder(session, drive_id, base_folder_id, year)
        cache[year_key] = year_folder_id
    month_folder_id = get_or_create_child_folder(session, drive_id, year_folder_id, month)
    cache[key] = month_folder_id
    return month_folder_id


def get_child_item_id(session: GraphSession, drive_id: str, parent_item_id: str, name: str) -> str | None:
    """Trouve l'id d'un enfant (fichier ou dossier) par nom, ou None s'il n'existe pas."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent_item_id}/children"
    resp = session.get(url, params={"$select": "id,name"}, timeout=30)
    raise_for_status_verbose(resp)
    for item in resp.json().get("value", []):
        if item.get("name") == name:
            return item["id"]
    return None


def download_processed_log(session: GraphSession, drive_id: str, folder_id: str) -> set:
    """Cles 'mailbox::message_id' deja traitees, depuis le log JSON depose a la
    racine du dossier cible. Ensemble vide si le fichier n'existe pas encore
    (premier run)."""
    item_id = get_child_item_id(session, drive_id, folder_id, LOG_FILENAME)
    if item_id is None:
        return set()
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"
    resp = session.get(url, timeout=30)
    raise_for_status_verbose(resp)
    try:
        return set(resp.json().get("processed_keys", []))
    except ValueError:
        return set()


def upload_processed_log(session: GraphSession, drive_id: str, folder_id: str, processed_keys: set) -> None:
    """Ecrit/ecrase le log JSON des messages traites, a la racine du dossier cible."""
    safe_name = quote(LOG_FILENAME)
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{folder_id}:/{safe_name}:/content"
    payload = json.dumps({"processed_keys": sorted(processed_keys)}, indent=2).encode("utf-8")
    resp = session.request("PUT", url, headers={"Content-Type": "application/json"}, data=payload, timeout=30)
    raise_for_status_verbose(resp)


MAX_UPLOAD_LOCK_RETRIES = 5
UPLOAD_LOCK_RETRY_DELAYS = [15, 30, 60, 90, 120]  # secondes


def create_upload_session_with_retry(session: GraphSession, url: str) -> requests.Response:
    """Cree une upload session SharePoint, en retentant si Graph renvoie un 409
    'nameAlreadyExists : A file with the same name is currently being uploaded'.
    Voir extract_banque.py pour le detail de ce cas (verrou temporaire, pas une
    vraie collision)."""
    for attempt in range(MAX_UPLOAD_LOCK_RETRIES + 1):
        resp = session.post(
            url,
            headers={"Content-Type": "application/json"},
            json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
            timeout=30,
        )
        is_upload_lock = resp.status_code == 409 and "currently being uploaded" in resp.text
        if not is_upload_lock or attempt >= MAX_UPLOAD_LOCK_RETRIES:
            raise_for_status_verbose(resp)
            return resp
        delay = UPLOAD_LOCK_RETRY_DELAYS[attempt]
        print(f"    (fichier deja en cours d'upload par une precedente tentative, nouvel essai dans {delay}s...)")
        time.sleep(delay)


def upload_file_to_sharepoint(session, drive_id, parent_item_id, filename, content):
    safe_name = quote(filename)
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent_item_id}:/{safe_name}:/createUploadSession"
    resp = create_upload_session_with_retry(session, url)
    upload_url = resp.json()["uploadUrl"]
    total = len(content)
    start = 0
    while start < total:
        end = min(start + CHUNK_SIZE, total) - 1
        chunk = content[start:end + 1]
        put_headers = {"Content-Length": str(len(chunk)), "Content-Range": f"bytes {start}-{end}/{total}"}
        resp = request_with_retry("PUT", upload_url, headers=put_headers, data=chunk, timeout=120)
        raise_for_status_verbose(resp)
        start = end + 1


def send_email(session, sender, recipients, subject, body_text):
    to_recipients = [{"emailAddress": {"address": a.strip()}} for a in recipients.split(",") if a.strip()]
    if not to_recipients:
        print("Aucun destinataire valide dans EMAIL_RECIPIENTS, email non envoye.")
        return
    payload = {
        "message": {"subject": subject, "body": {"contentType": "Text", "content": body_text}, "toRecipients": to_recipients},
        "saveToSentItems": "true",
    }
    resp = session.post(f"{GRAPH_BASE}/users/{sender}/sendMail", json=payload, timeout=30)
    raise_for_status_verbose(resp)
    print(f"Email envoye a {recipients} depuis {sender}.")


def matched_targets(msg: dict, targets: set) -> set:
    """Renvoie l'ensemble des adresses cibles (parmi targets) presentes en
    expediteur ou destinataire (a/cc) de ce message. Un message peut
    concerner plusieurs cibles a la fois (ex: Eddy ET Fara sur le meme fil)."""
    found = set()
    from_addr = ((msg.get("from") or {}).get("emailAddress") or {}).get("address", "").strip().lower()
    if from_addr in targets:
        found.add(from_addr)
    for r in (msg.get("toRecipients") or []) + (msg.get("ccRecipients") or []):
        addr = (r.get("emailAddress") or {}).get("address", "").strip().lower()
        if addr in targets:
            found.add(addr)
    return found


def run_extraction(session: GraphSession, senders: list[str], output_dir: str, sharepoint_link: str | None) -> dict:
    targets = {a.lower() for a in TARGET_ADDRESSES}
    print(f"Filtre actif : {sorted(targets)} en expediteur OU destinataire (a/cc) -- pas de filtre "
          f"mot-cle, toutes les pieces jointes de ces messages sont extraites.")

    sp_drive_id = sp_base_folder_id = None
    if sharepoint_link:
        print("Resolution du dossier SharePoint cible...")
        sp_drive_id, sp_base_folder_id = resolve_share_link(session, sharepoint_link)
        print(f"  -> driveId={sp_drive_id} folderId={sp_base_folder_id}")

    output_root = Path(output_dir)
    output_root.mkdir(parents=True, exist_ok=True)

    # Un etat par cible (Eddy, Fara) : son propre sous-dossier SharePoint
    # (meme nom que celui cree par extract_equipe.py -- slugify de l'adresse
    # -- donc fusion dans les memes dossiers par personne), son propre log de
    # reprise et son propre dedoublonnage (comme extract_equipe.py qui
    # dedoublonne par boite, pas globalement).
    target_state: dict[str, dict] = {}
    for target in sorted(targets):
        slug = slugify(target, max_len=80)
        sp_target_folder_id = None
        if sp_base_folder_id:
            sp_target_folder_id = get_or_create_child_folder(session, sp_drive_id, sp_base_folder_id, slug)
        processed_keys: set = set()
        if sp_target_folder_id:
            processed_keys = download_processed_log(session, sp_drive_id, sp_target_folder_id)
            if processed_keys:
                print(f"  {slug} : {len(processed_keys)} message(s) deja traites lors d'un run precedent (repris du log).")
        target_state[target] = {
            "slug": slug,
            "sp_folder_id": sp_target_folder_id,
            "sp_folder_cache": {},
            "processed_keys": processed_keys,
            "newly_processed_since_save": 0,
            "seen_hashes": set(),
            "count": 0,
        }

    total_files = 0
    total_duplicates = 0
    total_skipped_messages = 0
    skipped_mailboxes: list[dict] = []

    for mailbox in senders:
        print(f"\n--- Boite : {mailbox} ---")
        try:
            messages = list_all_messages_with_attachments(session, mailbox)
        except Exception as exc:
            print(f"  ! Boite ignoree ({exc})")
            skipped_mailboxes.append({"mailbox": mailbox, "detail": str(exc)})
            continue

        print(f"{len(messages)} email(s) avec piece(s) jointe(s) dans cette boite (envoi + reception).")

        for msg in messages:
            matched = matched_targets(msg, targets)
            if not matched:
                continue

            sender_address = ((msg.get("from") or {}).get("emailAddress") or {}).get("address") or ""
            subject = msg.get("subject") or "(sans objet)"
            received = (msg.get("receivedDateTime") or "")[:10]
            year = received[:4] if len(received) >= 7 else "date_inconnue"
            month = received[5:7] if len(received) >= 7 else "date_inconnue"

            attachments_cache = None
            for target in matched:
                state = target_state[target]
                key = f"{mailbox}::{msg['id']}"
                if key in state["processed_keys"]:
                    total_skipped_messages += 1
                    continue

                if attachments_cache is None:
                    try:
                        attachments_cache = list_attachments(session, mailbox, msg["id"])
                    except Exception as exc:
                        print(f"  ! Pieces jointes ignorees pour {subject!r} ({exc})")
                        attachments_cache = []
                file_attachments = [a for a in attachments_cache if a.get("@odata.type") == "#microsoft.graph.fileAttachment"]

                if not file_attachments:
                    state["processed_keys"].add(key)
                    state["newly_processed_since_save"] += 1
                    continue

                print(f"  [{received}] {subject} ({sender_address}) -> {state['slug']} - {len(file_attachments)} piece(s) jointe(s)")

                for att in file_attachments:
                    name = att.get("name") or f"piece_jointe_{att['id']}"
                    filename = build_filename(received, subject, name)
                    content = download_attachment_bytes(session, mailbox, msg["id"], att["id"])
                    content_hash = hashlib.sha256(content).hexdigest()
                    if content_hash in state["seen_hashes"]:
                        total_duplicates += 1
                        print(f"    -> doublon ignore: {name}")
                        continue
                    state["seen_hashes"].add(content_hash)

                    local_dir = output_root / state["slug"] / year / month
                    local_dir.mkdir(parents=True, exist_ok=True)
                    dest = unique_path(local_dir / filename)
                    dest.write_bytes(content)
                    total_files += 1
                    state["count"] += 1
                    print(f"    -> local: {dest}")

                    if state["sp_folder_id"]:
                        sp_month_folder_id = get_year_month_folder(
                            session, sp_drive_id, state["sp_folder_id"], year, month, state["sp_folder_cache"]
                        )
                        upload_file_to_sharepoint(session, sp_drive_id, sp_month_folder_id, dest.name, content)
                        print(f"    -> SharePoint: {state['slug']}/{year}/{month}/{dest.name}")

                state["processed_keys"].add(key)
                state["newly_processed_since_save"] += 1

                if state["sp_folder_id"] and state["newly_processed_since_save"] >= LOG_SAVE_EVERY:
                    upload_processed_log(session, sp_drive_id, state["sp_folder_id"], state["processed_keys"])
                    state["newly_processed_since_save"] = 0

    for state in target_state.values():
        if state["sp_folder_id"] and state["newly_processed_since_save"] > 0:
            upload_processed_log(session, sp_drive_id, state["sp_folder_id"], state["processed_keys"])

    per_target_counts = {state["slug"]: state["count"] for state in target_state.values()}

    print(f"\nTermine. {total_files} piece(s) jointe(s) enregistree(s) dans {output_root.resolve()}")
    if total_duplicates:
        print(f"{total_duplicates} doublon(s) ignore(s).")
    if total_skipped_messages:
        print(f"{total_skipped_messages} message(s) deja traites lors d'un run precedent, sautes.")
    if skipped_mailboxes:
        print(f"{len(skipped_mailboxes)} boite(s) ignoree(s) (erreur) : " + ", ".join(s["mailbox"] for s in skipped_mailboxes))
    return {
        "total_files": total_files,
        "total_duplicates": total_duplicates,
        "total_skipped_messages": total_skipped_messages,
        "per_target_counts": per_target_counts,
        "sharepoint_used": bool(sp_drive_id),
        "skipped_mailboxes": skipped_mailboxes,
    }


def build_summary_text(stats: dict) -> str:
    lines = [f"Extraction utilisateurs partis terminee le {date.today().isoformat()}."]
    lines.append(f"Total : {stats['total_files']} piece(s) jointe(s) enregistree(s).")
    if stats["total_duplicates"]:
        lines.append(f"Doublons ignores : {stats['total_duplicates']}.")
    if stats["total_skipped_messages"]:
        lines.append(f"Messages deja traites lors d'un run precedent (repris du log) : {stats['total_skipped_messages']}.")
    lines.append("")
    lines.append("Detail par personne :")
    for slug, count in stats["per_target_counts"].items():
        lines.append(f"  - {slug} : {count} piece(s) jointe(s)")
    if stats["skipped_mailboxes"]:
        lines.append("")
        lines.append("Boites ignorees (erreur d'acces) :")
        for s in stats["skipped_mailboxes"]:
            lines.append(f"  - {s['mailbox']} : {s['detail']}")
    if stats["sharepoint_used"]:
        lines.append("")
        lines.append("Fichiers deposes sur SharePoint, dans le sous-dossier de chaque personne (classes par annee/mois).")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="Extraction utilisateurs partis : pieces jointes des echanges avec Eddy/Fara, recherchees dans toutes les boites du staff.")
    parser.add_argument("--senders", nargs="+", default=DEFAULT_SENDERS)
    parser.add_argument("--output-dir", default="./deleted_users")
    parser.add_argument("--sharepoint-link", default=os.environ.get("SHAREPOINT_FOLDER_LINK_DELETED_USERS"))
    parser.add_argument("--email-sender", default=os.environ.get("EMAIL_SENDER"))
    parser.add_argument("--email-recipients", default=os.environ.get("EMAIL_RECIPIENTS"))
    args = parser.parse_args()

    tenant_id = os.environ.get("AZURE_TENANT_ID")
    client_id = os.environ.get("AZURE_CLIENT_ID")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET")
    if not all([tenant_id, client_id, client_secret]):
        print("Erreur : AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET requis.", file=sys.stderr)
        return 1

    print("Authentification Microsoft Graph (app-only)...")
    session = GraphSession(tenant_id, client_id, client_secret)
    session.ensure_ready()

    email_enabled = bool(args.email_sender and args.email_recipients)
    if not email_enabled:
        print("EMAIL_SENDER / EMAIL_RECIPIENTS non definis : pas d'email de confirmation envoye.")

    try:
        stats = run_extraction(session, args.senders, args.output_dir, args.sharepoint_link)
    except Exception as exc:
        error_text = f"{exc}\n\n{traceback.format_exc()}"
        print(f"ERREUR : {exc}", file=sys.stderr)
        if email_enabled:
            try:
                send_email(session, args.email_sender, args.email_recipients,
                           f"[ECHEC] Extraction utilisateurs partis - {date.today().isoformat()}",
                           f"L'extraction a echoue.\n\nErreur :\n{error_text}")
            except Exception as mail_exc:
                print(f"Echec de l'envoi de l'email d'alerte : {mail_exc}", file=sys.stderr)
        return 1

    if email_enabled:
        try:
            send_email(session, args.email_sender, args.email_recipients,
                       f"Extraction utilisateurs partis - {date.today().isoformat()} : {stats['total_files']} fichier(s)",
                       build_summary_text(stats))
        except Exception as mail_exc:
            print(f"Echec de l'envoi de l'email de confirmation : {mail_exc}", file=sys.stderr)

    return 0


if __name__ == "__main__":
    sys.exit(main())
