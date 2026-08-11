#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Jira Cloud API changelog analyzer - OAuth navigateur par defaut.

Objectif
--------
Appeler l'API Jira Cloud, recuperer les tickets et leur changelog complet,
puis produire :
  - un Excel d'analyse ;
  - une interface HTML locale avec filtres et timeline ticket par ticket.

Authentification
----------------
Ce script utilise OAuth 2.0 / 3LO Atlassian avec le navigateur par defaut.
Il n'ouvre pas Chromium/Playwright.

1) Creer un fichier .env dans le meme dossier que le script :

ATLASSIAN_CLIENT_ID=xxxxxxxxxxxxxxxx
ATLASSIAN_CLIENT_SECRET=xxxxxxxxxxxxxxxx
ATLASSIAN_BASE_URL=https://data-servier.atlassian.net
ATLASSIAN_PROJECT=PPMG

2) Lancer :

python jira_api_changelog_analyzer_oauth.py --types Story,Bug --output-xlsx jira.xlsx --output-html jira.html

Au premier lancement, le navigateur par defaut s'ouvre pour autoriser l'app.
Le token est ensuite stocke localement dans .jira_oauth_token.json.

Notes importantes
-----------------
- Le script appelle /rest/api/3/search/jql avec pagination nextPageToken.
- Le changelog complet est recupere via /rest/api/3/issue/{key}/changelog.
- Les commentaires sont recuperes via /rest/api/3/issue/{key}/comment.
- Les appels OAuth passent par api.atlassian.com/ex/jira/{cloudId}/...
"""

from __future__ import annotations

import argparse
import html
import json
import os
import re
import secrets
import sys
import time
import webbrowser
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import parse_qs, quote, urlencode, urlparse

try:
    import requests
except Exception as exc:  # pragma: no cover
    raise SystemExit("requests est requis. Lancez: pip install requests") from exc

try:
    import xlsxwriter
except Exception:  # pragma: no cover
    xlsxwriter = None


# =============================================================================
# Constantes
# =============================================================================

AUTH_URL = "https://auth.atlassian.com/authorize"
TOKEN_URL = "https://auth.atlassian.com/oauth/token"
ACCESSIBLE_RESOURCES_URL = "https://api.atlassian.com/oauth/token/accessible-resources"
DEFAULT_REDIRECT_URI = "http://localhost:8765/callback"
DEFAULT_TOKEN_FILE = ".jira_oauth_token.json"
DEFAULT_SCOPE = "read:jira-work read:jira-user offline_access"
MAX_EXCEL_TEXT = 32000
CONTROL_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


# =============================================================================
# Config / env
# =============================================================================


def load_env_file(path: Path = Path(".env")) -> None:
    """Charge un .env simple sans dependance externe."""
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


def env_or_arg(value: Optional[str], env_name: str, default: Optional[str] = None) -> Optional[str]:
    return value or os.environ.get(env_name) or default


# =============================================================================
# Helpers texte / dates
# =============================================================================


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return adf_to_text(value)
    text = str(value)
    text = html.unescape(text)
    text = CONTROL_RE.sub("", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def one_line(value: Any, limit: int = 300) -> str:
    text = clean_text(value).replace("\n", " | ")
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 3)].rstrip() + "..."


def normalize(value: Any) -> str:
    text = clean_text(value).lower()
    repl = str.maketrans({
        "à": "a", "â": "a", "ä": "a", "á": "a", "ã": "a", "å": "a",
        "ç": "c", "é": "e", "è": "e", "ê": "e", "ë": "e",
        "î": "i", "ï": "i", "í": "i", "ì": "i",
        "ô": "o", "ö": "o", "ó": "o", "ò": "o", "õ": "o",
        "ù": "u", "û": "u", "ü": "u", "ú": "u",
        "œ": "oe", "æ": "ae", "ñ": "n",
    })
    text = text.translate(repl)
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_jira_datetime(value: Any) -> Optional[datetime]:
    if not value:
        return None
    text = str(value).strip()
    fmts = [
        "%Y-%m-%dT%H:%M:%S.%f%z",
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%d %H:%M:%S%z",
        "%Y-%m-%d",
    ]
    for fmt in fmts:
        try:
            dt = datetime.strptime(text, fmt)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc)
        except Exception:
            pass
    try:
        fixed = text.replace("Z", "+00:00")
        fixed = re.sub(r"([+-]\d{2})(\d{2})$", r"\1:\2", fixed)
        dt = datetime.fromisoformat(fixed)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        return None


def dt_excel(dt: Optional[datetime]) -> Any:
    if not dt:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc).replace(tzinfo=None)


def dt_iso(dt: Optional[datetime]) -> str:
    if not dt:
        return ""
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc).isoformat()


def excel_safe(value: Any) -> Any:
    if isinstance(value, datetime):
        return dt_excel(value)
    if isinstance(value, (int, float, bool)) or value is None:
        return value
    text = clean_text(value)
    if len(text) > MAX_EXCEL_TEXT:
        text = text[: MAX_EXCEL_TEXT - 20] + "... [TRUNCATED]"
    if text.startswith("="):
        text = "'" + text
    return text


def split_csv(value: Optional[str]) -> List[str]:
    if not value:
        return []
    return [v.strip() for v in value.split(",") if v.strip()]


def jql_quote(value: str) -> str:
    return '"' + value.replace('"', '\\"') + '"'


# =============================================================================
# ADF vers texte
# =============================================================================


def adf_to_text(node: Any) -> str:
    """Convertit l'Atlassian Document Format en texte simple."""
    if node is None:
        return ""
    if isinstance(node, str):
        return node
    if isinstance(node, list):
        return "\n".join([adf_to_text(x) for x in node if adf_to_text(x)])
    if not isinstance(node, dict):
        return str(node)

    ntype = node.get("type")
    content = node.get("content", [])
    attrs = node.get("attrs", {}) or {}

    if ntype == "text":
        return node.get("text", "")
    if ntype == "hardBreak":
        return "\n"
    if ntype == "mention":
        return attrs.get("text", "@mention")
    if ntype == "emoji":
        return attrs.get("shortName", "")
    if ntype == "media":
        return attrs.get("alt", "[media]")
    if ntype in {"paragraph", "heading"}:
        return "".join(adf_to_text(x) for x in content).strip()
    if ntype == "bulletList":
        rows = []
        for x in content:
            txt = adf_to_text(x).strip()
            if txt:
                rows.append("- " + txt)
        return "\n".join(rows)
    if ntype == "orderedList":
        rows = []
        for idx, x in enumerate(content, start=1):
            txt = adf_to_text(x).strip()
            if txt:
                rows.append(f"{idx}. {txt}")
        return "\n".join(rows)
    if ntype == "listItem":
        return " ".join(adf_to_text(x).strip() for x in content if adf_to_text(x).strip())
    if ntype in {"blockquote", "codeBlock", "mediaSingle", "panel"}:
        return "\n".join(adf_to_text(x) for x in content)
    if content:
        return "\n".join(adf_to_text(x) for x in content)
    return ""


# =============================================================================
# Models
# =============================================================================


@dataclass
class IssueRow:
    issue_id: str
    key: str
    url: str
    summary: str
    project_key: str
    issue_type: str
    status: str
    status_category: str
    priority: str
    assignee: str
    reporter: str
    sprint_names: str
    sprint_ids: str
    created: Optional[datetime]
    updated: Optional[datetime]
    parent: str = ""


@dataclass
class TimelineEvent:
    issue_id: str
    key: str
    url: str
    summary: str
    project_key: str
    issue_type: str
    current_status: str
    current_assignee: str
    current_sprints: str
    event_date: Optional[datetime]
    author: str
    event_source: str
    field: str
    field_id: str
    change_type: str
    from_value: str
    to_value: str
    from_excerpt: str
    to_excerpt: str
    history_id: str = ""
    comment_id: str = ""

    def as_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["event_date"] = dt_excel(self.event_date)
        return d


# =============================================================================
# OAuth Atlassian via navigateur par defaut
# =============================================================================


class OAuthError(RuntimeError):
    pass


class OAuthCallbackHandler(BaseHTTPRequestHandler):
    auth_code: Optional[str] = None
    auth_state: Optional[str] = None
    auth_error: Optional[str] = None

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        OAuthCallbackHandler.auth_code = (qs.get("code") or [None])[0]
        OAuthCallbackHandler.auth_state = (qs.get("state") or [None])[0]
        OAuthCallbackHandler.auth_error = (qs.get("error") or [None])[0]
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        if OAuthCallbackHandler.auth_error:
            body = "<h1>Erreur OAuth</h1><p>Vous pouvez fermer cette page.</p>"
        else:
            body = "<h1>Autorisation Jira reçue</h1><p>Vous pouvez fermer cette page et revenir au terminal.</p>"
        self.wfile.write(body.encode("utf-8"))

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A003
        return


def read_token_file(token_file: Path) -> Optional[Dict[str, Any]]:
    if not token_file.exists():
        return None
    try:
        return json.loads(token_file.read_text(encoding="utf-8"))
    except Exception:
        return None


def write_token_file(token_file: Path, token_data: Dict[str, Any]) -> None:
    token_file.write_text(json.dumps(token_data, indent=2, ensure_ascii=False), encoding="utf-8")


def token_is_valid(token_data: Dict[str, Any], safety_seconds: int = 90) -> bool:
    expires_at = token_data.get("expires_at")
    if not expires_at:
        return False
    return time.time() < float(expires_at) - safety_seconds


class OAuthJiraClient:
    def __init__(
        self,
        base_url: str,
        client_id: str,
        client_secret: str,
        redirect_uri: str = DEFAULT_REDIRECT_URI,
        token_file: Path = Path(DEFAULT_TOKEN_FILE),
        scope: str = DEFAULT_SCOPE,
        timeout: int = 90,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.client_id = client_id
        self.client_secret = client_secret
        self.redirect_uri = redirect_uri
        self.token_file = token_file
        self.scope = scope
        self.timeout = timeout
        self.session = requests.Session()
        self.token_data = self.ensure_token()
        self.cloud_id = self.resolve_cloud_id()
        self.api_root = f"https://api.atlassian.com/ex/jira/{self.cloud_id}"

    def ensure_token(self) -> Dict[str, Any]:
        token_data = read_token_file(self.token_file)
        if token_data and token_is_valid(token_data):
            return token_data
        if token_data and token_data.get("refresh_token"):
            try:
                return self.refresh_token(token_data["refresh_token"])
            except Exception as exc:
                print(f"Refresh token invalide ou expiré: {exc}")
        return self.interactive_authorize()

    def interactive_authorize(self) -> Dict[str, Any]:
        state = secrets.token_urlsafe(24)
        OAuthCallbackHandler.auth_code = None
        OAuthCallbackHandler.auth_state = None
        OAuthCallbackHandler.auth_error = None

        auth_params = {
            "audience": "api.atlassian.com",
            "client_id": self.client_id,
            "scope": self.scope,
            "redirect_uri": self.redirect_uri,
            "state": state,
            "response_type": "code",
            "prompt": "consent",
        }
        auth_url = AUTH_URL + "?" + urlencode(auth_params)

        port = urlparse(self.redirect_uri).port or 8765
        server = HTTPServer(("localhost", port), OAuthCallbackHandler)
        print("Ouverture du navigateur par défaut pour autoriser Jira...")
        print("Si le navigateur ne s'ouvre pas, copiez-collez cette URL:")
        print(auth_url)
        webbrowser.open(auth_url)

        server.handle_request()
        server.server_close()

        if OAuthCallbackHandler.auth_error:
            raise OAuthError(f"Erreur OAuth: {OAuthCallbackHandler.auth_error}")
        if not OAuthCallbackHandler.auth_code:
            raise OAuthError("Aucun code OAuth reçu.")
        if OAuthCallbackHandler.auth_state != state:
            raise OAuthError("State OAuth incorrect. Autorisation refusée.")
        return self.exchange_code(OAuthCallbackHandler.auth_code)

    def exchange_code(self, code: str) -> Dict[str, Any]:
        payload = {
            "grant_type": "authorization_code",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "code": code,
            "redirect_uri": self.redirect_uri,
        }
        r = requests.post(TOKEN_URL, json=payload, timeout=self.timeout)
        if r.status_code >= 400:
            raise OAuthError(f"Echange code/token impossible: {r.status_code} {r.text[:1000]}")
        token_data = r.json()
        token_data["expires_at"] = time.time() + float(token_data.get("expires_in", 3600))
        write_token_file(self.token_file, token_data)
        return token_data

    def refresh_token(self, refresh_token: str) -> Dict[str, Any]:
        payload = {
            "grant_type": "refresh_token",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "refresh_token": refresh_token,
        }
        r = requests.post(TOKEN_URL, json=payload, timeout=self.timeout)
        if r.status_code >= 400:
            raise OAuthError(f"Refresh token impossible: {r.status_code} {r.text[:1000]}")
        token_data = r.json()
        token_data["expires_at"] = time.time() + float(token_data.get("expires_in", 3600))
        # Avec les refresh tokens rotatifs, Atlassian peut renvoyer un nouveau refresh_token.
        if not token_data.get("refresh_token"):
            token_data["refresh_token"] = refresh_token
        write_token_file(self.token_file, token_data)
        return token_data

    @property
    def access_token(self) -> str:
        if not token_is_valid(self.token_data):
            if self.token_data.get("refresh_token"):
                self.token_data = self.refresh_token(self.token_data["refresh_token"])
            else:
                self.token_data = self.interactive_authorize()
        return self.token_data["access_token"]

    def auth_headers(self) -> Dict[str, str]:
        return {"Authorization": f"Bearer {self.access_token}", "Accept": "application/json"}

    def resolve_cloud_id(self) -> str:
        r = requests.get(ACCESSIBLE_RESOURCES_URL, headers=self.auth_headers(), timeout=self.timeout)
        if r.status_code >= 400:
            raise RuntimeError(f"Impossible de lire accessible-resources: {r.status_code} {r.text[:1000]}")
        resources = r.json()
        if not isinstance(resources, list) or not resources:
            raise RuntimeError("Aucun site Atlassian accessible pour ce token.")
        target = self.base_url.rstrip("/").lower()
        for res in resources:
            if str(res.get("url", "")).rstrip("/").lower() == target:
                return str(res["id"])
        print("Sites accessibles:")
        for res in resources:
            print(f"- {res.get('name')} | {res.get('url')} | {res.get('id')}")
        raise RuntimeError(f"Le site {self.base_url} n'a pas été trouvé dans accessible-resources.")

    def get_json(self, path: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        url = self.api_root + path
        r = self.session.get(url, params=params or {}, headers=self.auth_headers(), timeout=self.timeout)
        if r.status_code == 401 and self.token_data.get("refresh_token"):
            self.token_data = self.refresh_token(self.token_data["refresh_token"])
            r = self.session.get(url, params=params or {}, headers=self.auth_headers(), timeout=self.timeout)
        if r.status_code >= 400:
            raise RuntimeError(f"Erreur API Jira {r.status_code} sur {path}: {r.text[:1000]}")
        return r.json()


# =============================================================================
# API fetch
# =============================================================================


def build_jql(args: argparse.Namespace) -> str:
    if args.jql:
        return args.jql
    project = args.project or "PPMG"
    clauses = [f"project = {project}"]
    types = split_csv(args.types)
    if types:
        quoted = ", ".join(jql_quote(x) for x in types)
        clauses.append(f"issuetype in ({quoted})")
    if args.api_sprint_id:
        clauses.append(f"Sprint = {args.api_sprint_id}")
    if args.api_sprint_name:
        clauses.append(f"Sprint = {jql_quote(args.api_sprint_name)}")
    if args.jql_extra:
        clauses.append(f"({args.jql_extra})")
    return " AND ".join(clauses) + " ORDER BY updated DESC"


def search_issues(client: OAuthJiraClient, jql: str, sprint_field: str, max_results: int, limit: Optional[int]) -> List[Dict[str, Any]]:
    fields = [
        "summary", "issuetype", "status", "assignee", "reporter", "created", "updated",
        "priority", sprint_field, "parent", "project",
    ]
    issues: List[Dict[str, Any]] = []
    token: Optional[str] = None
    while True:
        params: Dict[str, Any] = {
            "jql": jql,
            "maxResults": max_results,
            "fields": ",".join(fields),
        }
        if token:
            params["nextPageToken"] = token
        data = client.get_json("/rest/api/3/search/jql", params=params)
        batch = data.get("issues") or []
        issues.extend(batch)
        print(f"Tickets récupérés: {len(issues)}")
        if limit and len(issues) >= limit:
            return issues[:limit]
        if data.get("isLast") is True:
            break
        token = data.get("nextPageToken")
        if not token:
            break
    return issues


def fetch_all_changelog(client: OAuthJiraClient, issue_key_or_id: str) -> List[Dict[str, Any]]:
    histories: List[Dict[str, Any]] = []
    start_at = 0
    while True:
        data = client.get_json(
            f"/rest/api/3/issue/{quote(str(issue_key_or_id))}/changelog",
            params={"startAt": start_at, "maxResults": 100},
        )
        values = data.get("values") or data.get("histories") or []
        histories.extend(values)
        total = data.get("total")
        max_results = data.get("maxResults") or len(values) or 100
        if data.get("isLast") is True:
            break
        if total is not None and len(histories) >= int(total):
            break
        if not values:
            break
        start_at += int(max_results)
    return histories


def fetch_all_comments(client: OAuthJiraClient, issue_key_or_id: str) -> List[Dict[str, Any]]:
    comments: List[Dict[str, Any]] = []
    start_at = 0
    while True:
        data = client.get_json(
            f"/rest/api/3/issue/{quote(str(issue_key_or_id))}/comment",
            params={"startAt": start_at, "maxResults": 100, "orderBy": "created"},
        )
        values = data.get("comments") or []
        comments.extend(values)
        total = data.get("total")
        max_results = data.get("maxResults") or len(values) or 100
        if total is not None and len(comments) >= int(total):
            break
        if not values:
            break
        start_at += int(max_results)
    return comments


# =============================================================================
# Parse issue/changelog
# =============================================================================


def user_display(user: Any) -> str:
    if not isinstance(user, dict):
        return ""
    return user.get("displayName") or user.get("emailAddress") or user.get("accountId") or ""


def parse_sprint_value(value: Any) -> Tuple[str, str]:
    if not value:
        return "", ""
    if isinstance(value, list):
        names: List[str] = []
        ids: List[str] = []
        for item in value:
            if isinstance(item, dict):
                if item.get("name"):
                    names.append(str(item.get("name")))
                if item.get("id") is not None:
                    ids.append(str(item.get("id")))
            else:
                txt = str(item)
                m_name = re.search(r"name=([^,\]]+)", txt)
                m_id = re.search(r"id=(\d+)", txt)
                names.append(m_name.group(1) if m_name else txt)
                if m_id:
                    ids.append(m_id.group(1))
        return "; ".join(sorted(set(names))), "; ".join(sorted(set(ids)))
    if isinstance(value, dict):
        return str(value.get("name") or ""), str(value.get("id") or "")
    return str(value), ""


def parse_issue_row(issue: Dict[str, Any], base_url: str, sprint_field: str) -> IssueRow:
    fields = issue.get("fields") or {}
    status = fields.get("status") or {}
    status_cat = status.get("statusCategory") or {}
    issue_type = fields.get("issuetype") or {}
    assignee = fields.get("assignee")
    reporter = fields.get("reporter")
    priority = fields.get("priority") or {}
    project = fields.get("project") or {}
    sprint_names, sprint_ids = parse_sprint_value(fields.get(sprint_field))
    parent = fields.get("parent") or {}
    return IssueRow(
        issue_id=str(issue.get("id") or ""),
        key=str(issue.get("key") or ""),
        url=f"{base_url.rstrip('/')}/browse/{issue.get('key')}",
        summary=clean_text(fields.get("summary")),
        project_key=project.get("key") or "",
        issue_type=issue_type.get("name") or "",
        status=status.get("name") or "",
        status_category=status_cat.get("name") or status_cat.get("key") or "",
        priority=priority.get("name") or "",
        assignee=user_display(assignee),
        reporter=user_display(reporter),
        sprint_names=sprint_names,
        sprint_ids=sprint_ids,
        created=parse_jira_datetime(fields.get("created")),
        updated=parse_jira_datetime(fields.get("updated")),
        parent=parent.get("key") or "",
    )


def classify_field(item: Dict[str, Any]) -> Tuple[str, str]:
    field = clean_text(item.get("field") or "")
    field_id = clean_text(item.get("fieldId") or "")
    norm = normalize(field + " " + field_id)
    if field_id == "description" or norm.strip() == "description" or " description" in norm:
        return "Description", field_id or "description"
    if "acceptance criteria" in norm or "critere d acceptation" in norm or "criteres d acceptation" in norm:
        return "Acceptance Criteria", field_id
    if field_id == "assignee" or "assignee" in norm or "personne assignee" in norm:
        return "Assignation", field_id or "assignee"
    if field_id == "status" or norm == "status" or "etat" in norm or "statut" in norm:
        return "Statut", field_id or "status"
    if field_id == "summary" or "resume" in norm or "summary" in norm:
        return "Résumé", field_id or "summary"
    if "sprint" in norm:
        return "Sprint", field_id
    if field_id == "attachment" or "attachment" in norm or "piece jointe" in norm:
        return "Pièce jointe", field_id or "attachment"
    if "rank" in norm or "classement" in norm:
        return "Classement", field_id
    if "priority" in norm or "priorite" in norm:
        return "Priorité", field_id
    return field or field_id or "Autre", field_id


def changelog_to_events(issue_row: IssueRow, histories: List[Dict[str, Any]]) -> List[TimelineEvent]:
    events: List[TimelineEvent] = []
    for h in histories:
        h_id = str(h.get("id") or "")
        created = parse_jira_datetime(h.get("created"))
        author = user_display(h.get("author"))
        for item in h.get("items") or []:
            field_label, field_id = classify_field(item)
            from_val = clean_text(item.get("fromString") if item.get("fromString") is not None else item.get("from"))
            to_val = clean_text(item.get("toString") if item.get("toString") is not None else item.get("to"))
            change_type = f"{field_label} modifié"
            if field_label == "Assignation":
                change_type = "Assignation modifiée"
            elif field_label == "Statut":
                change_type = "Statut modifié"
            elif field_label == "Description":
                change_type = "Description modifiée"
            elif field_label == "Acceptance Criteria":
                change_type = "Acceptance Criteria modifiés"
            elif field_label == "Sprint":
                change_type = "Sprint modifié"
            events.append(TimelineEvent(
                issue_id=issue_row.issue_id,
                key=issue_row.key,
                url=issue_row.url,
                summary=issue_row.summary,
                project_key=issue_row.project_key,
                issue_type=issue_row.issue_type,
                current_status=issue_row.status,
                current_assignee=issue_row.assignee,
                current_sprints=issue_row.sprint_names,
                event_date=created,
                author=author,
                event_source="changelog",
                field=field_label,
                field_id=field_id,
                change_type=change_type,
                from_value=from_val,
                to_value=to_val,
                from_excerpt=one_line(from_val, 450),
                to_excerpt=one_line(to_val, 450),
                history_id=h_id,
            ))
    return events


def comments_to_events(issue_row: IssueRow, comments: List[Dict[str, Any]]) -> List[TimelineEvent]:
    events: List[TimelineEvent] = []
    for c in comments:
        created = parse_jira_datetime(c.get("created"))
        updated = parse_jira_datetime(c.get("updated"))
        author = user_display(c.get("author"))
        body = adf_to_text(c.get("body")) if isinstance(c.get("body"), dict) else clean_text(c.get("body"))
        comment_id = str(c.get("id") or "")
        events.append(TimelineEvent(
            issue_id=issue_row.issue_id,
            key=issue_row.key,
            url=issue_row.url,
            summary=issue_row.summary,
            project_key=issue_row.project_key,
            issue_type=issue_row.issue_type,
            current_status=issue_row.status,
            current_assignee=issue_row.assignee,
            current_sprints=issue_row.sprint_names,
            event_date=created,
            author=author,
            event_source="comment",
            field="Commentaire",
            field_id="comment",
            change_type="Commentaire créé",
            from_value="",
            to_value=body,
            from_excerpt="",
            to_excerpt=one_line(body, 450),
            comment_id=comment_id,
        ))
        if updated and created and abs((updated - created).total_seconds()) > 60:
            events.append(TimelineEvent(
                issue_id=issue_row.issue_id,
                key=issue_row.key,
                url=issue_row.url,
                summary=issue_row.summary,
                project_key=issue_row.project_key,
                issue_type=issue_row.issue_type,
                current_status=issue_row.status,
                current_assignee=issue_row.assignee,
                current_sprints=issue_row.sprint_names,
                event_date=updated,
                author=user_display(c.get("updateAuthor")) or author,
                event_source="comment",
                field="Commentaire",
                field_id="comment",
                change_type="Commentaire modifié",
                from_value="",
                to_value=body,
                from_excerpt="",
                to_excerpt=one_line(body, 450),
                comment_id=comment_id,
            ))
    return events


# =============================================================================
# Metrics
# =============================================================================


def sorted_events(events: Sequence[TimelineEvent]) -> List[TimelineEvent]:
    min_dt = datetime.min.replace(tzinfo=timezone.utc)
    return sorted(events, key=lambda e: (e.key, e.event_date or min_dt, e.field, e.history_id, e.comment_id))


def count_unique(events: Sequence[TimelineEvent]) -> int:
    return len({e.issue_id for e in events})


def count_return_loops(events: Sequence[TimelineEvent]) -> int:
    seq: List[str] = []
    min_dt = datetime.min.replace(tzinfo=timezone.utc)
    for e in sorted(events, key=lambda x: x.event_date or min_dt):
        val = clean_text(e.to_value)
        if val:
            seq.append(val)
    loops = 0
    for i in range(2, len(seq)):
        if seq[i] == seq[i - 2] and seq[i] != seq[i - 1]:
            loops += 1
    return loops


def row_from_issue(i: IssueRow) -> Dict[str, Any]:
    d = asdict(i)
    d["created"] = dt_excel(i.created)
    d["updated"] = dt_excel(i.updated)
    return d


def ticket_metrics(issues: Sequence[IssueRow], events: Sequence[TimelineEvent]) -> List[Dict[str, Any]]:
    by_issue: Dict[str, List[TimelineEvent]] = defaultdict(list)
    for e in events:
        by_issue[e.issue_id].append(e)
    rows: List[Dict[str, Any]] = []
    for issue in sorted(issues, key=lambda x: x.key):
        evs = by_issue.get(issue.issue_id, [])
        desc = [e for e in evs if e.field == "Description"]
        ac = [e for e in evs if e.field == "Acceptance Criteria"]
        assign = [e for e in evs if e.field == "Assignation"]
        status = [e for e in evs if e.field == "Statut"]
        sprint = [e for e in evs if e.field == "Sprint"]
        comments = [e for e in evs if e.field == "Commentaire"]
        first = min((e.event_date for e in evs if e.event_date), default=None)
        last = max((e.event_date for e in evs if e.event_date), default=None)
        churn_score = len(desc) * 4 + len(ac) * 5 + len(status) * 3 + len(assign) * 2 + len(sprint) * 2 + len(comments)
        rows.append({
            "Issue ID": issue.issue_id,
            "Key": issue.key,
            "URL": issue.url,
            "Résumé": issue.summary,
            "Type": issue.issue_type,
            "Statut actuel": issue.status,
            "Assigné actuel": issue.assignee,
            "Sprint actuel": issue.sprint_names,
            "Créé": dt_excel(issue.created),
            "Mis à jour": dt_excel(issue.updated),
            "Premier événement": dt_excel(first),
            "Dernier événement": dt_excel(last),
            "Événements total": len(evs),
            "Description changes": len(desc),
            "AC changes": len(ac),
            "Status changes": len(status),
            "Assignment changes": len(assign),
            "Sprint changes": len(sprint),
            "Comment events": len(comments),
            "Assignment loops": count_return_loops(assign),
            "Status loops": count_return_loops(status),
            "Churn score": churn_score,
        })
    return rows


def events_by_day(events: Sequence[TimelineEvent]) -> List[Dict[str, Any]]:
    counter: Counter[str] = Counter()
    for e in events:
        if e.event_date:
            counter[e.event_date.astimezone(timezone.utc).date().isoformat()] += 1
    return [{"Date": k, "Événements": v} for k, v in sorted(counter.items())]


def events_by_field(events: Sequence[TimelineEvent]) -> List[Dict[str, Any]]:
    counter: Counter[str] = Counter(e.field for e in events)
    return [
        {"Champ": k, "Événements": v, "Tickets": len({e.issue_id for e in events if e.field == k})}
        for k, v in counter.most_common()
    ]


def assignment_transitions(events: Sequence[TimelineEvent]) -> List[Dict[str, Any]]:
    counter: Counter[Tuple[str, str]] = Counter()
    for e in events:
        if e.field == "Assignation":
            counter[(e.from_value or "(vide)", e.to_value or "(vide)")] += 1
    return [{"De": a, "Vers": b, "Transitions": n} for (a, b), n in counter.most_common()]


def status_transitions(events: Sequence[TimelineEvent]) -> List[Dict[str, Any]]:
    counter: Counter[Tuple[str, str]] = Counter()
    for e in events:
        if e.field == "Statut":
            counter[(e.from_value or "(vide)", e.to_value or "(vide)")] += 1
    return [{"De": a, "Vers": b, "Transitions": n} for (a, b), n in counter.most_common()]


# =============================================================================
# Excel output
# =============================================================================


def style_palette(style: str) -> Dict[str, str]:
    if style == "audit":
        return {"dark": "1F2937", "primary": "2563EB", "accent": "F59E0B", "green": "10B981", "red": "EF4444", "light": "F3F4F6", "teal": "0F766E"}
    if style == "kanban":
        return {"dark": "0F172A", "primary": "7C3AED", "accent": "06B6D4", "green": "22C55E", "red": "F97316", "light": "F8FAFC", "teal": "14B8A6"}
    return {"dark": "17365D", "primary": "1F4E78", "accent": "F4B183", "green": "A9D18E", "red": "F8CBAD", "light": "EAF2F8", "teal": "9DC3E6"}


def build_formats(wb: Any, p: Dict[str, str]) -> Dict[str, Any]:
    return {
        "title": wb.add_format({"bold": True, "font_size": 18, "font_color": "white", "bg_color": p["dark"], "align": "center", "valign": "vcenter"}),
        "subtitle": wb.add_format({"bold": True, "font_size": 12, "font_color": p["dark"]}),
        "header": wb.add_format({"bold": True, "font_color": "white", "bg_color": p["primary"], "border": 1, "align": "center", "valign": "vcenter"}),
        "card_label": wb.add_format({"bold": True, "font_color": "white", "bg_color": p["primary"], "align": "center", "valign": "vcenter", "border": 1}),
        "card_value": wb.add_format({"bold": True, "font_size": 18, "font_color": p["dark"], "bg_color": p["light"], "align": "center", "valign": "vcenter", "border": 1}),
        "date": wb.add_format({"num_format": "yyyy-mm-dd hh:mm", "valign": "top"}),
        "number": wb.add_format({"num_format": "#,##0", "valign": "top"}),
        "bool": wb.add_format({"valign": "top"}),
        "note": wb.add_format({"italic": True, "font_color": "#666666"}),
        "wrap": wb.add_format({"text_wrap": True, "valign": "top"}),
        "desc": wb.add_format({"bg_color": "#E2F0D9"}),
        "ac": wb.add_format({"bg_color": "#FFF2CC"}),
        "status": wb.add_format({"bg_color": "#D9EAF7"}),
        "assign": wb.add_format({"bg_color": "#FCE4D6"}),
    }


def cell_to_rowcol(cell: str) -> Tuple[int, int]:
    m = re.match(r"([A-Z]+)(\d+)", cell)
    if not m:
        return 0, 0
    letters, row_s = m.groups()
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return int(row_s) - 1, col - 1


def write_rows(ws: Any, rows: List[Dict[str, Any]], formats: Dict[str, Any], autofilter: bool = True) -> None:
    if not rows:
        ws.write(0, 0, "Aucune donnée", formats["note"])
        return
    headers = list(rows[0].keys())
    for c, h in enumerate(headers):
        ws.write(0, c, h, formats["header"])
    for r, row in enumerate(rows, start=1):
        for c, h in enumerate(headers):
            val = excel_safe(row.get(h))
            fmt = None
            if isinstance(val, datetime):
                fmt = formats["date"]
            elif isinstance(val, bool):
                fmt = formats["bool"]
            elif isinstance(val, (int, float)):
                fmt = formats["number"]
            ws.write(r, c, val, fmt)
    ws.freeze_panes(1, 0)
    if autofilter and rows:
        ws.autofilter(0, 0, len(rows), len(headers) - 1)
    for c, h in enumerate(headers):
        width = len(h) + 2
        for row in rows[:300]:
            width = max(width, min(80, len(str(row.get(h, "") or "")) + 2))
        if any(x in h.lower() for x in ["résumé", "url", "excerpt", "value", "valeur", "lien", "summary"]):
            width = min(max(width, 24), 72)
        else:
            width = min(max(width, 10), 32)
        ws.set_column(c, c, width)


def add_timeline_conditional_formats(ws: Any, nrows: int, fmt: Dict[str, Any]) -> None:
    field_col = 12
    ws.conditional_format(1, field_col, nrows, field_col, {"type": "text", "criteria": "containing", "value": "Description", "format": fmt["desc"]})
    ws.conditional_format(1, field_col, nrows, field_col, {"type": "text", "criteria": "containing", "value": "Acceptance", "format": fmt["ac"]})
    ws.conditional_format(1, field_col, nrows, field_col, {"type": "text", "criteria": "containing", "value": "Statut", "format": fmt["status"]})
    ws.conditional_format(1, field_col, nrows, field_col, {"type": "text", "criteria": "containing", "value": "Assignation", "format": fmt["assign"]})


def write_dashboard(wb: Any, fmt: Dict[str, Any], p: Dict[str, str], issues: List[IssueRow], events: List[TimelineEvent], by_day: List[Dict[str, Any]], by_field: List[Dict[str, Any]], ticket_rows: List[Dict[str, Any]]) -> None:
    ws = wb.add_worksheet("00_Dashboard")
    ws.hide_gridlines(2)
    ws.merge_range("A1:N2", "Jira changelog analysis - API OAuth", fmt["title"])
    cards = [
        ("Tickets", len(issues)),
        ("Événements", len(events)),
        ("Tickets Description", len({e.issue_id for e in events if e.field == "Description"})),
        ("Tickets AC", len({e.issue_id for e in events if e.field == "Acceptance Criteria"})),
        ("Changements statut", sum(e.field == "Statut" for e in events)),
        ("Réassignations", sum(e.field == "Assignation" for e in events)),
        ("Commentaires", sum(e.field == "Commentaire" for e in events)),
        ("Tickets instables", sum(1 for r in ticket_rows if int(r.get("Churn score", 0)) >= 10)),
    ]
    positions = ["A4", "D4", "G4", "J4", "A8", "D8", "G8", "J8"]
    for (label, value), pos in zip(cards, positions):
        row, col = cell_to_rowcol(pos)
        ws.merge_range(row, col, row, col + 1, label, fmt["card_label"])
        ws.merge_range(row + 1, col, row + 2, col + 1, value, fmt["card_value"])
    ws.write("A13", "Lecture", fmt["subtitle"])
    ws.write("A14", "Utilisez 01_Timeline et l'HTML pour filtrer par Type, Sprint, Statut, Champ et Ticket.", fmt["note"])

    start_row = 18
    ws.write(start_row, 0, "Champ", fmt["header"])
    ws.write(start_row, 1, "Événements", fmt["header"])
    for idx, row in enumerate(by_field[:12], start=start_row + 1):
        ws.write(idx, 0, row["Champ"])
        ws.write_number(idx, 1, row["Événements"])
    if by_field:
        chart = wb.add_chart({"type": "column"})
        chart.add_series({
            "name": "Événements",
            "categories": ["00_Dashboard", start_row + 1, 0, start_row + min(12, len(by_field)), 0],
            "values": ["00_Dashboard", start_row + 1, 1, start_row + min(12, len(by_field)), 1],
        })
        chart.set_title({"name": "Événements par champ"})
        chart.set_legend({"none": True})
        chart.set_size({"width": 620, "height": 320})
        ws.insert_chart("D18", chart)

    day_row = 36
    ws.write(day_row, 0, "Date", fmt["header"])
    ws.write(day_row, 1, "Événements", fmt["header"])
    for idx, row in enumerate(by_day[:60], start=day_row + 1):
        ws.write(idx, 0, row["Date"])
        ws.write_number(idx, 1, row["Événements"])
    if by_day:
        chart2 = wb.add_chart({"type": "line"})
        chart2.add_series({
            "name": "Événements",
            "categories": ["00_Dashboard", day_row + 1, 0, day_row + min(60, len(by_day)), 0],
            "values": ["00_Dashboard", day_row + 1, 1, day_row + min(60, len(by_day)), 1],
        })
        chart2.set_title({"name": "Évolution dans le temps"})
        chart2.set_size({"width": 620, "height": 300})
        ws.insert_chart("D36", chart2)
    ws.set_column("A:A", 24)
    ws.set_column("B:B", 14)
    ws.set_column("D:N", 16)


def validate_xlsx(path: Path) -> None:
    import zipfile
    with zipfile.ZipFile(path, "r") as z:
        tables = [n for n in z.namelist() if n.startswith("xl/tables/")]
        if tables:
            raise RuntimeError(f"Le fichier contient des tables Excel structurees interdites: {tables[:5]}")


def write_workbook(path: Path, issues: List[IssueRow], events: List[TimelineEvent], style: str = "executive") -> None:
    if xlsxwriter is None:
        raise SystemExit("xlsxwriter est requis. Lancez: pip install xlsxwriter")
    wb = xlsxwriter.Workbook(str(path))
    palette = style_palette(style)
    formats = build_formats(wb, palette)

    issue_rows = [row_from_issue(i) for i in issues]
    timeline_rows = [e.as_dict() for e in sorted_events(events)]
    ticket_rows = ticket_metrics(issues, events)
    desc_ac_rows = [e.as_dict() for e in sorted_events(events) if e.field in {"Description", "Acceptance Criteria"}]
    comment_rows = [e.as_dict() for e in sorted_events(events) if e.field == "Commentaire"]
    ass_rows = assignment_transitions(events)
    status_rows = status_transitions(events)
    by_day = events_by_day(events)
    by_field = events_by_field(events)

    write_dashboard(wb, formats, palette, issues, events, by_day, by_field, ticket_rows)
    sheets = [
        ("01_Timeline", timeline_rows),
        ("02_Tickets", ticket_rows),
        ("03_Current_Issues", issue_rows),
        ("04_Desc_AC", desc_ac_rows),
        ("05_Comments", comment_rows),
        ("06_Assignations", ass_rows),
        ("07_Status", status_rows),
        ("08_By_Day", by_day),
        ("09_By_Field", by_field),
    ]
    for name, rows in sheets:
        ws = wb.add_worksheet(name)
        write_rows(ws, rows, formats, autofilter=True)
        if name == "01_Timeline" and timeline_rows:
            add_timeline_conditional_formats(ws, len(timeline_rows), formats)
    wb.set_properties({"title": "Jira API changelog analysis", "comments": "Generated from Jira REST API changelog"})
    wb.close()
    validate_xlsx(path)


# =============================================================================
# HTML dashboard
# =============================================================================


def serialize_issue(i: IssueRow) -> Dict[str, Any]:
    d = asdict(i)
    d["created"] = dt_iso(i.created)
    d["updated"] = dt_iso(i.updated)
    return d


def serialize_event(e: TimelineEvent) -> Dict[str, Any]:
    d = asdict(e)
    d["event_date"] = dt_iso(e.event_date)
    d["from_value"] = one_line(e.from_value, 1200)
    d["to_value"] = one_line(e.to_value, 1200)
    return d


def write_html(path: Path, issues: List[IssueRow], events: List[TimelineEvent], style: str = "executive") -> None:
    palette = style_palette(style)
    payload = {
        "issues": [serialize_issue(i) for i in issues],
        "events": [serialize_event(e) for e in sorted_events(events)],
    }
    data_json = json.dumps(payload, ensure_ascii=False)
    css_vars = "\n".join([f"--{k}: #{v};" for k, v in palette.items()])
    html_text = f"""<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Jira changelog timeline</title>
<style>
:root {{ {css_vars} }}
* {{ box-sizing: border-box; }}
body {{ margin:0; font-family: Segoe UI, Arial, sans-serif; background:#f7f9fc; color:#111827; }}
header {{ background:var(--dark); color:white; padding:22px 32px; }}
header h1 {{ margin:0; font-size:24px; }}
header p {{ margin:6px 0 0 0; opacity:.85; }}
.container {{ padding:22px 32px; }}
.filters {{ display:grid; grid-template-columns: repeat(6, minmax(160px,1fr)); gap:12px; background:white; padding:16px; border-radius:14px; box-shadow:0 1px 5px rgba(15,23,42,.10); position:sticky; top:0; z-index:5; }}
label {{ font-size:12px; font-weight:700; color:#374151; display:block; margin-bottom:4px; }}
select,input {{ width:100%; padding:9px; border:1px solid #cbd5e1; border-radius:8px; background:white; }}
.cards {{ display:grid; grid-template-columns: repeat(6, 1fr); gap:14px; margin:18px 0; }}
.card {{ background:white; border-left:6px solid var(--primary); border-radius:14px; padding:14px; box-shadow:0 1px 5px rgba(15,23,42,.10); }}
.card .label {{ font-size:12px; color:#64748b; font-weight:700; }}
.card .value {{ font-size:26px; font-weight:800; color:var(--dark); margin-top:4px; }}
.grid {{ display:grid; grid-template-columns: 390px 1fr; gap:18px; align-items:start; }}
.panel {{ background:white; border-radius:14px; padding:16px; box-shadow:0 1px 5px rgba(15,23,42,.10); }}
.panel h2 {{ margin:0 0 12px 0; color:var(--dark); font-size:18px; }}
.ticket {{ border:1px solid #e5e7eb; border-radius:10px; padding:10px; margin-bottom:8px; cursor:pointer; }}
.ticket:hover {{ border-color:var(--primary); background:#f8fbff; }}
.ticket.active {{ border-color:var(--accent); background:#fffbeb; }}
.ticket .key {{ font-weight:800; color:var(--primary); }}
.ticket .meta {{ font-size:12px; color:#64748b; margin-top:4px; }}
.timeline {{ position:relative; margin-left:12px; border-left:3px solid #dbeafe; padding-left:18px; }}
.event {{ position:relative; margin:0 0 14px 0; padding:12px; border-radius:12px; background:#f8fafc; border:1px solid #e5e7eb; }}
.event:before {{ content:""; position:absolute; left:-27px; top:16px; width:14px; height:14px; border-radius:50%; background:var(--primary); border:3px solid white; box-shadow:0 0 0 2px #bfdbfe; }}
.event.desc:before {{ background:var(--green); }}
.event.ac:before {{ background:var(--accent); }}
.event.status:before {{ background:var(--teal); }}
.event.assign:before {{ background:var(--red); }}
.event .top {{ display:flex; justify-content:space-between; gap:12px; }}
.event .field {{ font-weight:800; color:var(--dark); }}
.event .date {{ color:#64748b; font-size:12px; white-space:nowrap; }}
.diff {{ display:grid; grid-template-columns: 1fr 1fr; gap:10px; margin-top:10px; }}
.diff div {{ background:white; border-radius:8px; padding:8px; border:1px solid #e5e7eb; font-size:12px; white-space:pre-wrap; max-height:180px; overflow:auto; }}
.badge {{ display:inline-block; padding:3px 8px; border-radius:999px; background:#e0f2fe; color:#075985; font-size:11px; font-weight:700; margin-right:4px; }}
.bars {{ margin-top:12px; }}
.bar {{ display:grid; grid-template-columns: 150px 1fr 50px; gap:8px; align-items:center; margin:6px 0; font-size:12px; }}
.bar span:nth-child(2) {{ height:12px; background:linear-gradient(90deg,var(--primary),var(--teal)); border-radius:999px; }}
@media(max-width:1100px) {{ .filters,.cards {{ grid-template-columns: repeat(2, 1fr); }} .grid {{ grid-template-columns:1fr; }} }}
</style>
</head>
<body>
<header><h1>Jira changelog timeline</h1><p>Filtres par type, sprint, ticket, champ et auteur</p></header>
<div class="container">
  <div class="filters">
    <div><label>Type</label><select id="typeFilter"></select></div>
    <div><label>Sprint</label><select id="sprintFilter"></select></div>
    <div><label>Statut actuel</label><select id="statusFilter"></select></div>
    <div><label>Champ modifié</label><select id="fieldFilter"></select></div>
    <div><label>Ticket</label><select id="ticketFilter"></select></div>
    <div><label>Recherche</label><input id="searchBox" placeholder="PPMG, texte, auteur..."/></div>
  </div>
  <div class="cards" id="cards"></div>
  <div class="grid">
    <div class="panel"><h2>Tickets</h2><div id="ticketList"></div></div>
    <div class="panel"><h2>Timeline</h2><div id="timeline"></div></div>
  </div>
</div>
<script>
const DATA = {data_json};
const issues = DATA.issues;
const events = DATA.events;
const state = {{ type:'Tous', sprint:'Tous', status:'Tous', field:'Tous', ticket:'Tous', q:'' }};
function uniq(arr) {{ return ['Tous', ...Array.from(new Set(arr.filter(x => x && String(x).trim()))).sort((a,b)=>String(a).localeCompare(String(b)))] }}
function fillSelect(id, values) {{ const el=document.getElementById(id); el.innerHTML=''; values.forEach(v=>{{ const o=document.createElement('option'); o.value=v; o.textContent=v; el.appendChild(o); }}); el.onchange=()=>{{ const key=id.replace('Filter',''); state[key]=el.value; render(); }} }}
function init() {{
  fillSelect('typeFilter', uniq(issues.map(i=>i.issue_type)));
  fillSelect('sprintFilter', uniq(issues.flatMap(i=>String(i.sprint_names||'').split('; ').filter(Boolean))));
  fillSelect('statusFilter', uniq(issues.map(i=>i.status)));
  fillSelect('fieldFilter', uniq(events.map(e=>e.field)));
  fillSelect('ticketFilter', uniq(issues.map(i=>i.key)));
  document.getElementById('searchBox').oninput = e => {{ state.q=e.target.value.toLowerCase(); render(); }};
  render();
}}
function issuePass(i) {{
  if(state.type !== 'Tous' && i.issue_type !== state.type) return false;
  if(state.sprint !== 'Tous' && !String(i.sprint_names||'').split('; ').includes(state.sprint)) return false;
  if(state.status !== 'Tous' && i.status !== state.status) return false;
  if(state.ticket !== 'Tous' && i.key !== state.ticket) return false;
  const q = state.q;
  if(q && !(String(i.key+' '+i.summary+' '+i.assignee+' '+i.sprint_names).toLowerCase().includes(q))) return false;
  return true;
}}
function eventPass(e, issueSet) {{
  if(!issueSet.has(e.issue_id)) return false;
  if(state.field !== 'Tous' && e.field !== state.field) return false;
  const q = state.q;
  if(q && !(String(e.key+' '+e.summary+' '+e.author+' '+e.field+' '+e.from_excerpt+' '+e.to_excerpt).toLowerCase().includes(q))) return false;
  return true;
}}
function render() {{
  const filteredIssues = issues.filter(issuePass);
  const issueSet = new Set(filteredIssues.map(i=>i.issue_id));
  const filteredEvents = events.filter(e=>eventPass(e, issueSet));
  renderCards(filteredIssues, filteredEvents);
  renderTickets(filteredIssues, filteredEvents);
  renderTimeline(filteredEvents);
}}
function renderCards(is, es) {{
  const descTickets = new Set(es.filter(e=>e.field==='Description').map(e=>e.issue_id)).size;
  const acTickets = new Set(es.filter(e=>e.field==='Acceptance Criteria').map(e=>e.issue_id)).size;
  const cards = [
    ['Tickets', is.length], ['Événements', es.length], ['Tickets Description', descTickets], ['Tickets AC', acTickets],
    ['Réassignations', es.filter(e=>e.field==='Assignation').length], ['Statuts', es.filter(e=>e.field==='Statut').length]
  ];
  document.getElementById('cards').innerHTML = cards.map(c=>`<div class="card"><div class="label">${{c[0]}}</div><div class="value">${{c[1]}}</div></div>`).join('');
}}
function renderTickets(is, es) {{
  const counts = new Map(); es.forEach(e=>counts.set(e.issue_id,(counts.get(e.issue_id)||0)+1));
  const sorted = [...is].sort((a,b)=>(counts.get(b.issue_id)||0)-(counts.get(a.issue_id)||0));
  document.getElementById('ticketList').innerHTML = sorted.slice(0,200).map(i=>`<div class="ticket ${{state.ticket===i.key?'active':''}}" onclick="state.ticket='${{i.key}}'; document.getElementById('ticketFilter').value='${{i.key}}'; render();"><span class="key">${{i.key}}</span> <span class="badge">${{i.issue_type}}</span><div>${{escapeHtml(i.summary||'')}}</div><div class="meta">${{i.status||''}} · ${{i.assignee||''}} · ${{i.sprint_names||''}} · ${{counts.get(i.issue_id)||0}} event(s)</div></div>`).join('');
}}
function eventClass(e) {{ if(e.field==='Description') return 'desc'; if(e.field==='Acceptance Criteria') return 'ac'; if(e.field==='Statut') return 'status'; if(e.field==='Assignation') return 'assign'; return ''; }}
function renderTimeline(es) {{
  const sorted = [...es].sort((a,b)=>String(b.event_date).localeCompare(String(a.event_date))).slice(0,600);
  if(!sorted.length) {{ document.getElementById('timeline').innerHTML='<p>Aucun événement pour les filtres sélectionnés.</p>'; return; }}
  document.getElementById('timeline').innerHTML = `<div class="bars">${{barsHtml(sorted)}}</div><div class="timeline">` + sorted.map(e=>`<div class="event ${{eventClass(e)}}"><div class="top"><div><span class="field">${{escapeHtml(e.field)}}</span> <span class="badge">${{escapeHtml(e.key)}}</span> <span class="badge">${{escapeHtml(e.author||'')}}</span></div><div class="date">${{formatDate(e.event_date)}}</div></div><div>${{escapeHtml(e.change_type)}} · <a href="${{e.url}}" target="_blank">ouvrir Jira</a></div><div class="diff"><div><b>Avant</b><br>${{escapeHtml(e.from_excerpt||e.from_value||'')}}</div><div><b>Après</b><br>${{escapeHtml(e.to_excerpt||e.to_value||'')}}</div></div></div>`).join('') + `</div>`;
}}
function barsHtml(es) {{ const c={{}}; es.forEach(e=>c[e.field]=(c[e.field]||0)+1); const max=Math.max(...Object.values(c),1); return Object.entries(c).sort((a,b)=>b[1]-a[1]).map(([k,v])=>`<div class="bar"><span>${{escapeHtml(k)}}</span><span style="width:${{Math.max(4,Math.round(v/max*100))}}%"></span><b>${{v}}</b></div>`).join(''); }}
function escapeHtml(s) {{ return String(s||'').replace(/[&<>"']/g, m=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[m])); }}
function formatDate(s) {{ if(!s) return ''; const d=new Date(s); return isNaN(d) ? s : d.toLocaleString('fr-FR'); }}
init();
</script>
</body>
</html>"""
    path.write_text(html_text, encoding="utf-8")


# =============================================================================
# Demo + orchestration
# =============================================================================


def demo_data(base_url: str) -> Tuple[List[IssueRow], List[TimelineEvent]]:
    issue = IssueRow(
        issue_id="457866",
        key="PPMG-2185",
        url=f"{base_url.rstrip('/')}/browse/PPMG-2185",
        summary="[INC0911736] - Le Business Project Status reste à Feasibility Study après passage en phase Réalisation",
        project_key="PPMG",
        issue_type="Bug",
        status="Ready for Testing",
        status_category="A faire",
        priority="Medium",
        assignee="Yosra MANAI",
        reporter="Romain CIESLIK",
        sprint_names="Sprint 16 2026",
        sprint_ids="16029",
        created=parse_jira_datetime("2026-08-06T15:32:18.633+0200"),
        updated=parse_jira_datetime("2026-08-10T10:08:46.228+0200"),
        parent="",
    )
    events = [
        TimelineEvent(issue.issue_id, issue.key, issue.url, issue.summary, issue.project_key, issue.issue_type, issue.status, issue.assignee, issue.sprint_names, parse_jira_datetime("2026-08-06T15:32:18.633+0200"), "Romain CIESLIK", "changelog", "Description", "description", "Description modifiée", "Ancienne description", "Nouvelle description avec reprise de données", "Ancienne description", "Nouvelle description avec reprise de données", "1", ""),
        TimelineEvent(issue.issue_id, issue.key, issue.url, issue.summary, issue.project_key, issue.issue_type, issue.status, issue.assignee, issue.sprint_names, parse_jira_datetime("2026-08-06T16:07:53.565+0200"), "Romain CIESLIK", "changelog", "Statut", "status", "Statut modifié", "Open", "Validated", "Open", "Validated", "2", ""),
        TimelineEvent(issue.issue_id, issue.key, issue.url, issue.summary, issue.project_key, issue.issue_type, issue.status, issue.assignee, issue.sprint_names, parse_jira_datetime("2026-08-10T09:45:20.886+0200"), "Romain CIESLIK", "changelog", "Assignation", "assignee", "Assignation modifiée", "Romain CIESLIK", "RACHEL CLOIX", "Romain CIESLIK", "RACHEL CLOIX", "3", ""),
        TimelineEvent(issue.issue_id, issue.key, issue.url, issue.summary, issue.project_key, issue.issue_type, issue.status, issue.assignee, issue.sprint_names, parse_jira_datetime("2026-08-10T10:08:46.228+0200"), "RACHEL CLOIX", "changelog", "Assignation", "assignee", "Assignation modifiée", "RACHEL CLOIX", "Yosra MANAI", "RACHEL CLOIX", "Yosra MANAI", "4", ""),
        TimelineEvent(issue.issue_id, issue.key, issue.url, issue.summary, issue.project_key, issue.issue_type, issue.status, issue.assignee, issue.sprint_names, parse_jira_datetime("2026-08-06T15:39:24.796+0200"), "Romain CIESLIK", "comment", "Commentaire", "comment", "Commentaire créé", "", "Commentaire de test sur la reproduction du bug", "", "Commentaire de test sur la reproduction du bug", "", "356323"),
    ]
    return [issue], events


def fetch_real_data(args: argparse.Namespace) -> Tuple[List[IssueRow], List[TimelineEvent]]:
    client = OAuthJiraClient(
        base_url=args.base_url,
        client_id=args.client_id,
        client_secret=args.client_secret,
        redirect_uri=args.redirect_uri,
        token_file=Path(args.token_file),
        scope=args.scope,
    )
    jql = build_jql(args)
    print("JQL:", jql)
    issues_raw = search_issues(client, jql, args.sprint_field, args.max_results, args.limit)
    issues = [parse_issue_row(i, args.base_url, args.sprint_field) for i in issues_raw]

    sprint_contains = normalize(args.sprint_contains or "")
    if sprint_contains:
        keep_ids = {i.issue_id for i in issues if sprint_contains in normalize(i.sprint_names)}
        issues = [i for i in issues if i.issue_id in keep_ids]
        print(f"Filtre local sprint_contains={args.sprint_contains}: {len(issues)} ticket(s)")

    all_events: List[TimelineEvent] = []
    for idx, issue in enumerate(issues, start=1):
        print(f"[{idx}/{len(issues)}] Historique {issue.key}")
        histories = fetch_all_changelog(client, issue.key)
        all_events.extend(changelog_to_events(issue, histories))
        if not args.no_comments:
            try:
                comments = fetch_all_comments(client, issue.key)
                all_events.extend(comments_to_events(issue, comments))
            except Exception as exc:
                print(f"  Warning commentaires {issue.key}: {exc}")
        time.sleep(args.pause)
    return issues, all_events


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    load_env_file()
    ap = argparse.ArgumentParser(description="Analyse Jira Cloud changelog via OAuth et genere Excel + HTML timeline.")
    ap.add_argument("--base-url", default=os.environ.get("ATLASSIAN_BASE_URL", "https://data-servier.atlassian.net"), help="URL Jira Cloud")
    ap.add_argument("--project", default=os.environ.get("ATLASSIAN_PROJECT", "PPMG"), help="Cle projet Jira, defaut PPMG")
    ap.add_argument("--types", default="Story,Bug", help="Types a recuperer, ex: Story,Bug ou Story")
    ap.add_argument("--api-sprint-id", default=None, help="Filtre JQL par ID de sprint, ex: 17833")
    ap.add_argument("--api-sprint-name", default=None, help="Filtre JQL par nom exact de sprint, ex: Sprint 17 2026")
    ap.add_argument("--sprint-contains", default=None, help="Filtre local si vous donnez juste 15/16/17, ex: 17")
    ap.add_argument("--jql", default=None, help="JQL complet. Remplace --project/--types.")
    ap.add_argument("--jql-extra", default=None, help="Clause JQL additionnelle")
    ap.add_argument("--sprint-field", default="customfield_10020", help="Champ sprint Jira, defaut customfield_10020")
    ap.add_argument("--client-id", default=os.environ.get("ATLASSIAN_CLIENT_ID"), help="OAuth Client ID")
    ap.add_argument("--client-secret", default=os.environ.get("ATLASSIAN_CLIENT_SECRET"), help="OAuth Client Secret")
    ap.add_argument("--redirect-uri", default=os.environ.get("ATLASSIAN_REDIRECT_URI", DEFAULT_REDIRECT_URI))
    ap.add_argument("--scope", default=os.environ.get("ATLASSIAN_SCOPE", DEFAULT_SCOPE))
    ap.add_argument("--token-file", default=os.environ.get("ATLASSIAN_TOKEN_FILE", DEFAULT_TOKEN_FILE))
    ap.add_argument("--max-results", type=int, default=100, help="Taille page search/jql")
    ap.add_argument("--limit", type=int, default=None, help="Limite de tickets pour test")
    ap.add_argument("--pause", type=float, default=0.05, help="Pause entre tickets")
    ap.add_argument("--no-comments", action="store_true", help="Ne pas recuperer les commentaires")
    ap.add_argument("--output-xlsx", default="jira_changelog_analysis.xlsx")
    ap.add_argument("--output-html", default="jira_changelog_dashboard.html")
    ap.add_argument("--style", choices=["executive", "audit", "kanban"], default="executive")
    ap.add_argument("--demo", action="store_true", help="Genere un exemple sans appeler Jira")
    return ap.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = parse_args(argv)
    if not args.demo:
        if not args.client_id:
            raise SystemExit("ATLASSIAN_CLIENT_ID manquant. Mettez-le dans .env ou passez --client-id.")
        if not args.client_secret:
            raise SystemExit("ATLASSIAN_CLIENT_SECRET manquant. Mettez-le dans .env ou passez --client-secret.")

    if args.demo:
        issues, events = demo_data(args.base_url)
    else:
        issues, events = fetch_real_data(args)

    out_xlsx = Path(args.output_xlsx)
    out_html = Path(args.output_html)
    write_workbook(out_xlsx, issues, events, style=args.style)
    write_html(out_html, issues, events, style=args.style)
    print("\nTerminé.")
    print(f"Excel : {out_xlsx.resolve()}")
    print(f"HTML  : {out_html.resolve()}")
    print(f"Tickets: {len(issues)} | Événements: {len(events)}")
    try:
        webbrowser.open(out_html.resolve().as_uri())
    except Exception:
        pass


if __name__ == "__main__":
    main()
