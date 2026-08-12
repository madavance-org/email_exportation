#!/usr/bin/env python3
"""
extract_ov_baobab.py -- "Extraction OV x BAOBAB"

Flow separe (dossier SharePoint et script dedies) : INTERSECTION des filtres
OV (extract_attachments.py) et BAOBAB (extract_baobab.py). Un message doit
satisfaire LES DEUX conditions :
    - signal BAOBAB : expediteur mgildas@baobab.com, uraharitiana@baobab.com,
      ou toute autre adresse se terminant par "@baobab.com" ; OU mot "BAOBAB"
      dans l'objet/nom de fichier (ajoute le 13/08/2026 -- certains OV lies a
      BAOBAB sont envoyes en interne, ex: par eddy.rajaonarivony@madavance.org,
      sans expediteur @baobab.com) ;
    - signal OV : objet ou nom de piece jointe avec le token "OV" (avec
      eventuels chiffres a la suite, ex "OV12"), ou "virement"
      (memes regles que extract_attachments.py / DEFAULT_KEYWORD="OV").

Autrement dit : les Ordres de Virement lies a BAOBAB.

Cibles par defaut : olivia@madavance.org et rakitrynyavo@madavance.org
(memes boites que les deux scripts source). Modifiable via --senders.

Authentification identique aux autres scripts du repo (BackupOffice365,
app-only) :
    AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET

Destination SharePoint :
    SHAREPOINT_FOLDER_LINK_OV_BAOBAB

Usage :
    python extract_ov_baobab.py
"""

import argparse
import base64
import hashlib
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
DEFAULT_SENDERS = ["olivia@madavance.org", "rakitrynyavo@madavance.org"]
BAOBAB_EXACT_SENDERS = {"mgildas@baobab.com", "uraharitiana@baobab.com"}
BAOBAB_DOMAIN_SUFFIX = "@baobab.com"
OV_KEYWORD = "OV"
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
    # Retente aussi sur les erreurs reseau bas niveau (connexion coupee,
    # timeout en cours d'ecriture...), pas seulement sur les codes HTTP
    # transitoires (429/5xx) d'une reponse recue.
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
        "$select": "id,subject,from,receivedDateTime,hasAttachments",
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


def normalize_text(value: str) -> str:
    return unicodedata.normalize("NFKD", value or "").encode("ascii", "ignore").decode("ascii").lower()


def make_keyword_matcher(keyword: str):
    """Meme regle que extract_attachments.py : token exact (insensible a la
    casse), chiffres autorises a la suite (ex: 'OV12')."""
    if not keyword:
        return lambda text: True
    pattern = re.compile(rf"^{re.escape(keyword)}\d*$", re.IGNORECASE)

    def matcher(text: str) -> bool:
        if not text:
            return False
        tokens = re.split(r"[^A-Za-z0-9]+", text)
        return any(pattern.match(t) for t in tokens if t)

    return matcher


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


MAX_UPLOAD_LOCK_RETRIES = 5
UPLOAD_LOCK_RETRY_DELAYS = [15, 30, 60, 90, 120]  # secondes


def create_upload_session_with_retry(session: GraphSession, url: str) -> requests.Response:
    """Cree une upload session SharePoint, en retentant si Graph renvoie un 409
    'nameAlreadyExists : A file with the same name is currently being uploaded'.

    Ce cas n'est pas une vraie collision de nom (conflictBehavior=replace gere
    deja le remplacement d'un fichier existant) : c'est un verrou temporaire
    laisse par une precedente tentative d'upload sur ce meme chemin, interrompue
    avant d'avoir termine (job annule ou relance en cours d'upload). Ce verrou
    expire de lui-meme cote SharePoint apres un moment -- on retente avec un
    delai croissant plutot que de faire planter tout le run."""
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


def is_baobab_sender(sender_address: str) -> bool:
    addr = (sender_address or "").strip().lower()
    if not addr:
        return False
    return addr in BAOBAB_EXACT_SENDERS or addr.endswith(BAOBAB_DOMAIN_SUFFIX)


def is_baobab_text(text: str) -> bool:
    """Deuxieme signal BAOBAB (13/08/2026) : mot 'baobab' dans objet/nom de
    fichier, pour les OV envoyes en interne (ex: par Eddy) sans expediteur
    @baobab.com. Voir aussi extract_baobab.py."""
    return "baobab" in normalize_text(text)


def run_extraction(session: GraphSession, senders: list[str], output_dir: str, sharepoint_link: str | None) -> dict:
    ov_matches = make_keyword_matcher(OV_KEYWORD)

    def ov_content_matches(text: str) -> bool:
        return ov_matches(text) or "virement" in normalize_text(text)

    print(f"Filtre actif (intersection) : signal BAOBAB (expediteur egal a {sorted(BAOBAB_EXACT_SENDERS)} ou "
          f"terminant par '{BAOBAB_DOMAIN_SUFFIX}', OU 'baobab' dans objet/nom de fichier) ET signal OV "
          f"(objet/nom de fichier avec le token '{OV_KEYWORD}' ou contenant 'virement').")

    sp_drive_id = sp_folder_id = None
    if sharepoint_link:
        print("Resolution du dossier SharePoint cible...")
        sp_drive_id, sp_folder_id = resolve_share_link(session, sharepoint_link)
        print(f"  -> driveId={sp_drive_id} folderId={sp_folder_id}")

    output_root = Path(output_dir)
    output_root.mkdir(parents=True, exist_ok=True)

    seen_hashes: set = set()
    sp_folder_cache: dict = {}
    total_files = 0
    total_duplicates = 0
    per_mailbox_counts: dict[str, int] = {}

    for mailbox in senders:
        print(f"\n--- Boite : {mailbox} ---")
        messages = list_all_messages_with_attachments(session, mailbox)
        print(f"{len(messages)} email(s) avec piece(s) jointe(s) dans cette boite (envoi + reception).")
        per_mailbox_counts[mailbox] = 0

        for msg in messages:
            sender_address = ((msg.get("from") or {}).get("emailAddress") or {}).get("address") or ""
            subject = msg.get("subject") or "(sans objet)"
            received = (msg.get("receivedDateTime") or "")[:10]

            subject_baobab = is_baobab_sender(sender_address) or is_baobab_text(subject)
            subject_ov = ov_content_matches(subject)

            attachments = list_attachments(session, mailbox, msg["id"])
            file_attachments = [a for a in attachments if a.get("@odata.type") == "#microsoft.graph.fileAttachment"]
            if not file_attachments:
                continue

            kept = []
            for a in file_attachments:
                name = a.get("name") or ""
                baobab_ok = subject_baobab or is_baobab_text(name)
                ov_ok = subject_ov or ov_content_matches(name)
                if baobab_ok and ov_ok:
                    kept.append(a)
            if not kept:
                continue

            print(f"  [{received}] {subject} ({sender_address}) - {len(kept)}/{len(file_attachments)} piece(s) jointe(s) retenue(s)")
            year = received[:4] if len(received) >= 7 else "date_inconnue"
            month = received[5:7] if len(received) >= 7 else "date_inconnue"

            for att in kept:
                name = att.get("name") or f"piece_jointe_{att['id']}"
                filename = build_filename(received, subject, name)
                content = download_attachment_bytes(session, mailbox, msg["id"], att["id"])
                content_hash = hashlib.sha256(content).hexdigest()
                if content_hash in seen_hashes:
                    total_duplicates += 1
                    print(f"    -> doublon ignore: {name}")
                    continue
                seen_hashes.add(content_hash)

                local_dir = output_root / year / month
                local_dir.mkdir(parents=True, exist_ok=True)
                dest = unique_path(local_dir / filename)
                dest.write_bytes(content)
                total_files += 1
                per_mailbox_counts[mailbox] += 1
                print(f"    -> local: {dest}")

                if sp_folder_id:
                    sp_month_folder_id = get_year_month_folder(session, sp_drive_id, sp_folder_id, year, month, sp_folder_cache)
                    upload_file_to_sharepoint(session, sp_drive_id, sp_month_folder_id, dest.name, content)
                    print(f"    -> SharePoint: {year}/{month}/{dest.name}")

    print(f"\nTermine. {total_files} piece(s) jointe(s) enregistree(s) dans {output_root.resolve()}")
    if total_duplicates:
        print(f"{total_duplicates} doublon(s) ignore(s).")
    return {
        "total_files": total_files,
        "total_duplicates": total_duplicates,
        "per_mailbox_counts": per_mailbox_counts,
        "sharepoint_used": bool(sp_drive_id),
    }


def build_summary_text(stats: dict) -> str:
    lines = [f"Extraction OV x BAOBAB terminee le {date.today().isoformat()}."]
    lines.append(f"Total : {stats['total_files']} piece(s) jointe(s) enregistree(s).")
    if stats["total_duplicates"]:
        lines.append(f"Doublons ignores : {stats['total_duplicates']}.")
    lines.append("")
    lines.append("Detail par boite :")
    for mailbox, count in stats["per_mailbox_counts"].items():
        lines.append(f"  - {mailbox} : {count} piece(s) jointe(s)")
    if stats["sharepoint_used"]:
        lines.append("")
        lines.append("Fichiers deposes sur SharePoint (classes par annee/mois).")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="Extraction OV x BAOBAB : pieces jointes des Ordres de Virement envoyes par BAOBAB.")
    parser.add_argument("--senders", nargs="+", default=DEFAULT_SENDERS)
    parser.add_argument("--output-dir", default="./ov_baobab")
    parser.add_argument("--sharepoint-link", default=os.environ.get("SHAREPOINT_FOLDER_LINK_OV_BAOBAB"))
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
                           f"[ECHEC] Extraction OV x BAOBAB - {date.today().isoformat()}",
                           f"L'extraction OV x BAOBAB a echoue.\n\nErreur :\n{error_text}")
            except Exception as mail_exc:
                print(f"Echec de l'envoi de l'email d'alerte : {mail_exc}", file=sys.stderr)
        return 1

    if email_enabled:
        try:
            send_email(session, args.email_sender, args.email_recipients,
                       f"Extraction OV x BAOBAB - {date.today().isoformat()} : {stats['total_files']} fichier(s)",
                       build_summary_text(stats))
        except Exception as mail_exc:
            print(f"Echec de l'envoi de l'email de confirmation : {mail_exc}", file=sys.stderr)

    return 0


if __name__ == "__main__":
    sys.exit(main())
