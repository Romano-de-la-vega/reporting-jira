#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Jira Cloud changelog analyzer - browser cookies mode.

Goal:
    Use the user's existing browser session to call Jira REST API, without
    API token, OAuth admin approval, Playwright or Chromium.

What it does:
    - reads cookies from Edge / Chrome / Firefox local profile with browser-cookie3;
    - calls Jira REST API endpoints that the user can already open in the browser;
    - retrieves issues matching project / type / sprint / explicit issue keys;
    - retrieves full changelog and comments per ticket;
    - exports a rich Excel workbook;
    - exports a local HTML dashboard with ticket-by-ticket timeline;
    - optionally opens a simple Tkinter GUI to configure and launch the extraction.

Important:
    This does not bypass Jira permissions. It reuses the currently logged-in
    browser session of the Windows user. If the browser session cannot access a
    Jira issue, the script cannot access it either.

Install:
    pip install -r requirements_jira_cookie.txt

GUI:
    python jira_api_cookie_analyzer.py --gui

CLI example:
    python jira_api_cookie_analyzer.py --browser edge --project PPMG --types Story,Bug --limit 20
"""

from __future__ import annotations

import argparse
import difflib
import html
import json
import difflib
import os
import re
import sys
import threading
import time
import webbrowser
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import quote

try:
    import requests
except Exception:  # pragma: no cover
    requests = None

try:
    import xlsxwriter
except Exception:  # pragma: no cover
    xlsxwriter = None

CONTROL_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
MAX_EXCEL = 32000
DEFAULT_BASE_URL = "https://data-servier.atlassian.net"
DEFAULT_PROJECT = "PPMG"
DEFAULT_SPRINT_FIELD = "customfield_10020"


# =============================================================================
# Text / date helpers
# =============================================================================


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        value = adf_to_text(value)
    text = html.unescape(str(value))
    text = CONTROL_RE.sub("", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def one_line(value: Any, limit: int = 320) -> str:
    text = clean_text(value).replace("\n", " | ")
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 3)].rstrip() + "..."

def clip_text(value: Any, limit: int = 20000) -> str:
    """Keep a long but bounded text value for local HTML diff display."""
    text = clean_text(value)
    if len(text) <= limit:
        return text
    return text[:limit] + "\n...[texte tronqué pour performance]"


def diff_context(before: Any, after: Any, context: int = 120) -> Tuple[str, str]:
    """Plain-text focus on the changed area, useful in Excel.

    Uses common prefix/suffix detection. For a single comma or word change, the
    exact modified fragment is isolated between << >>.
    """
    a = clean_text(before)
    b = clean_text(after)
    if a == b:
        return "", ""
    i = 0
    max_i = min(len(a), len(b))
    while i < max_i and a[i] == b[i]:
        i += 1
    j = 0
    max_j = min(len(a) - i, len(b) - i)
    while j < max_j and a[len(a) - 1 - j] == b[len(b) - 1 - j]:
        j += 1
    a_changed = a[i: len(a) - j if j else len(a)]
    b_changed = b[i: len(b) - j if j else len(b)]
    prefix = a[max(0, i - context): i]
    suffix = a[len(a) - j: min(len(a), len(a) - j + context)] if j else ""
    prefix_marker = "..." if i > context else ""
    suffix_marker = "..." if j and (len(a) - j + context < len(a)) else ""
    before_focus = f"{prefix_marker}{prefix}<<{a_changed}>>{suffix}{suffix_marker}"
    after_focus = f"{prefix_marker}{prefix}<<{b_changed}>>{suffix}{suffix_marker}"
    return before_focus, after_focus




def clip_for_diff(value: Any, limit: int = 12000) -> str:
    text = clean_text(value)
    if len(text) <= limit:
        return text
    half = max(1000, limit // 2)
    return text[:half] + "\n\n[...] TEXTE TRONQUE POUR AFFICHAGE [...]\n\n" + text[-half:]


def diff_tokens(value: Any) -> List[str]:
    text = clip_for_diff(value)
    # Words, numbers, punctuation and whitespace are separate tokens.
    # A comma-only change is therefore highlighted.
    return re.findall(r"\w+|[^\w\s]|\s+", text, flags=re.UNICODE)


def build_diff_html(before: Any, after: Any) -> str:
    a = diff_tokens(before)
    b = diff_tokens(after)
    matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)
    parts: List[str] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            text = ''.join(a[i1:i2])
            parts.append(html.escape(text))
        elif tag == "delete":
            text = ''.join(a[i1:i2])
            if text:
                parts.append(f'<span class="diff-del">{html.escape(text)}</span>')
        elif tag == "insert":
            text = ''.join(b[j1:j2])
            if text:
                parts.append(f'<span class="diff-add">{html.escape(text)}</span>')
        elif tag == "replace":
            old = ''.join(a[i1:i2])
            new = ''.join(b[j1:j2])
            if old:
                parts.append(f'<span class="diff-del">{html.escape(old)}</span>')
            if new:
                parts.append(f'<span class="diff-add">{html.escape(new)}</span>')
    return ''.join(parts) or '<span class="muted">Aucune difference textuelle exploitable</span>'


def build_diff_text(before: Any, after: Any, limit: int = 6000) -> str:
    a = diff_tokens(before)
    b = diff_tokens(after)
    matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)
    parts: List[str] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            parts.append(''.join(a[i1:i2]))
        elif tag == "delete":
            text = ''.join(a[i1:i2])
            if text:
                parts.append(f"[-{text}-]")
        elif tag == "insert":
            text = ''.join(b[j1:j2])
            if text:
                parts.append(f"[+{text}+]")
        elif tag == "replace":
            old = ''.join(a[i1:i2])
            new = ''.join(b[j1:j2])
            if old:
                parts.append(f"[-{old}-]")
            if new:
                parts.append(f"[+{new}+]")
    text = ''.join(parts)
    return text if len(text) <= limit else text[:limit-20] + "... [TRUNCATED]"


def change_magnitude(before: Any, after: Any) -> float:
    a = clean_text(before)
    b = clean_text(after)
    if not a and not b:
        return 0.0
    ratio = difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()
    return round(1 - ratio, 4)

def normalize(value: Any) -> str:
    text = clean_text(value).lower()
    replacements = str.maketrans({
        "à": "a", "â": "a", "ä": "a", "á": "a", "ã": "a", "å": "a",
        "ç": "c", "é": "e", "è": "e", "ê": "e", "ë": "e",
        "î": "i", "ï": "i", "í": "i", "ì": "i",
        "ô": "o", "ö": "o", "ó": "o", "ò": "o", "õ": "o",
        "ù": "u", "û": "u", "ü": "u", "ú": "u",
        "œ": "oe", "æ": "ae", "ñ": "n",
    })
    text = text.translate(replacements)
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_jira_datetime(value: Any) -> Optional[datetime]:
    if not value:
        return None
    text = str(value).strip()
    formats = [
        "%Y-%m-%dT%H:%M:%S.%f%z",
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%d %H:%M:%S%z",
        "%Y-%m-%d",
    ]
    for fmt in formats:
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
    if dt is None:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc).replace(tzinfo=None)


def dt_iso(dt: Optional[datetime]) -> str:
    if dt is None:
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
    if len(text) > MAX_EXCEL:
        text = text[: MAX_EXCEL - 15] + "... [TRUNCATED]"
    if text.startswith("="):
        text = "'" + text
    return text


def split_csv(value: Optional[str]) -> List[str]:
    if not value:
        return []
    return [v.strip() for v in value.split(",") if v.strip()]


def diff_tokens(value: str) -> List[str]:
    # Tokens that keep spaces and punctuation, so even a comma-only change is visible.
    return re.findall(r"\w+|\s+|[^\w\s]", clean_text(value), flags=re.UNICODE)


def compact_equal_text(text: str, limit: int = 140) -> str:
    if len(text) <= limit:
        return text
    head = text[: max(20, limit // 2)].rstrip()
    tail = text[-max(20, limit // 2):].lstrip()
    return head + " … " + tail


def inline_diff_html(before: str, after: str) -> str:
    """HTML diff that highlights exactly what changed.

    It is intentionally token-based and punctuation-aware. It is compact enough
    for long Jira descriptions while still showing tiny changes such as commas.
    """
    before = clean_text(before)
    after = clean_text(after)
    if not before and not after:
        return ""
    a = diff_tokens(before)
    b = diff_tokens(after)
    matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)
    parts: List[str] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        old = "".join(a[i1:i2])
        new = "".join(b[j1:j2])
        if tag == "equal":
            txt = compact_equal_text(old)
            if txt:
                parts.append(html.escape(txt))
        elif tag == "delete":
            if old:
                parts.append(f"<del>{html.escape(old)}</del>")
        elif tag == "insert":
            if new:
                parts.append(f"<ins>{html.escape(new)}</ins>")
        elif tag == "replace":
            if old:
                parts.append(f"<del>{html.escape(old)}</del>")
            if new:
                parts.append(f"<ins>{html.escape(new)}</ins>")
    return "".join(parts)


def diff_text_summary(before: str, after: str, max_items: int = 10) -> str:
    before = clean_text(before)
    after = clean_text(after)
    if before == after:
        return ""
    a = diff_tokens(before)
    b = diff_tokens(after)
    matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)
    chunks: List[str] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        old = one_line("".join(a[i1:i2]), 220)
        new = one_line("".join(b[j1:j2]), 220)
        if tag == "delete" and old:
            chunks.append("- " + old)
        elif tag == "insert" and new:
            chunks.append("+ " + new)
        elif tag == "replace":
            if old:
                chunks.append("- " + old)
            if new:
                chunks.append("+ " + new)
        if len(chunks) >= max_items:
            break
    if not chunks:
        return "Modification d'espaces ou de ponctuation uniquement"
    return " | ".join(chunks[:max_items])


def change_magnitude(before: str, after: str) -> int:
    before = clean_text(before)
    after = clean_text(after)
    if before == after:
        return 0
    ratio = difflib.SequenceMatcher(None, before, after, autojunk=False).ratio()
    return max(1, round((1 - ratio) * 100))


# =============================================================================
# Atlassian Document Format -> text
# =============================================================================


def adf_to_text(node: Any) -> str:
    if node is None:
        return ""
    if isinstance(node, str):
        return node
    if isinstance(node, list):
        parts = [adf_to_text(x) for x in node]
        return "\n".join([p for p in parts if p])
    if not isinstance(node, dict):
        return str(node)

    ntype = node.get("type")
    if ntype == "text":
        return node.get("text", "")
    if ntype == "hardBreak":
        return "\n"
    if ntype == "mention":
        return node.get("attrs", {}).get("text", "@mention")
    if ntype == "emoji":
        return node.get("attrs", {}).get("shortName", "")
    if ntype == "media":
        return node.get("attrs", {}).get("alt", "[media]")

    content = node.get("content", [])
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
    if ntype in {"blockquote", "codeBlock", "mediaSingle"}:
        return "\n".join(adf_to_text(x) for x in content)
    if content:
        return "\n".join(adf_to_text(x) for x in content)
    return ""


# =============================================================================
# Data models
# =============================================================================


@dataclass
class IssueRow:
    issue_id: str
    key: str
    url: str
    summary: str
    project_key: str
    project_name: str
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
        d["diff_precis_texte"] = build_diff_text(self.from_value, self.to_value)
        d["ampleur_changement"] = change_magnitude(self.from_value, self.to_value)
        return d


# =============================================================================
# Jira client using browser cookies
# =============================================================================


class JiraCookieClient:
    def __init__(
        self,
        base_url: str,
        browser: str = "auto",
        timeout: int = 90,
        log: Optional[Callable[[str], None]] = None,
    ):
        if requests is None:
            raise SystemExit("requests est requis. Lancez: pip install requests")
        self.base_url = base_url.rstrip("/")
        self.browser = browser.lower().strip() or "auto"
        self.timeout = timeout
        self.log = log or print
        self.session = requests.Session()
        self.session.headers.update({
            "Accept": "application/json,text/plain,*/*",
            "User-Agent": "Mozilla/5.0 JiraCookieAnalyzer/1.0",
        })
        self.reload_cookies()

    def reload_cookies(self) -> None:
        try:
            import browser_cookie3
        except Exception as exc:
            raise SystemExit("browser-cookie3 est requis. Lancez: pip install browser-cookie3") from exc

        browser_order = [self.browser] if self.browser != "auto" else ["edge", "chrome", "firefox"]
        errors = []
        for b in browser_order:
            try:
                if b in {"edge", "msedge"}:
                    jar = browser_cookie3.edge()
                elif b in {"chrome", "google-chrome", "chromium"}:
                    jar = browser_cookie3.chrome()
                elif b == "firefox":
                    jar = browser_cookie3.firefox()
                else:
                    raise ValueError(f"Navigateur non supporte: {b}")
                self.session.cookies.clear()
                self.session.cookies.update(jar)
                self.log(f"Cookies charges depuis {b}.")
                return
            except Exception as exc:
                errors.append(f"{b}: {exc}")
        raise RuntimeError("Impossible de lire les cookies navigateur. Details: " + " | ".join(errors))

    def open_login_page(self) -> None:
        url = f"{self.base_url}/rest/api/3/myself"
        self.log("Ouverture de Jira dans le navigateur par defaut...")
        webbrowser.open(url)

    def get_json(self, path: str, params: Optional[Dict[str, Any]] = None, retry_login: bool = True) -> Dict[str, Any]:
        url = self.base_url + path
        response = self.session.get(url, params=params or {}, timeout=self.timeout)
        if response.status_code in {401, 403} and retry_login:
            self.log(f"Acces refuse ({response.status_code}). Tentative apres rechargement des cookies...")
            self.reload_cookies()
            response = self.session.get(url, params=params or {}, timeout=self.timeout)
        if response.status_code in {401, 403}:
            msg = response.text[:800]
            raise RuntimeError(
                f"API Jira refusee ({response.status_code}) sur {path}.\n"
                "Ouvrez Jira dans votre navigateur, verifiez que /rest/api/3/myself affiche du JSON, "
                "puis relancez le script.\n"
                f"Reponse: {msg}"
            )
        if response.status_code >= 400:
            raise RuntimeError(f"API Jira error {response.status_code} on {path}: {response.text[:1200]}")
        try:
            return response.json()
        except Exception as exc:
            raise RuntimeError(f"Reponse non JSON pour {path}: {response.text[:500]}") from exc

    def ensure_logged_in(self) -> Dict[str, Any]:
        try:
            me = self.get_json("/rest/api/3/myself", retry_login=False)
            if me.get("accountId"):
                self.log(f"Session Jira OK: {me.get('displayName') or me.get('emailAddress') or me.get('accountId')}")
                return me
        except Exception:
            pass
        self.open_login_page()
        input("Connectez-vous dans votre navigateur par defaut, puis appuyez sur Entree ici... ")
        self.reload_cookies()
        me = self.get_json("/rest/api/3/myself", retry_login=False)
        if not me.get("accountId"):
            raise RuntimeError("Session Jira non detectee apres connexion.")
        self.log(f"Session Jira OK: {me.get('displayName') or me.get('emailAddress') or me.get('accountId')}")
        return me


# =============================================================================
# API fetch
# =============================================================================


def jql_quote(value: str) -> str:
    escaped = value.replace('"', '\\"')
    return f'"{escaped}"'


def build_jql(args: argparse.Namespace) -> str:
    if args.jql:
        return args.jql
    clauses = [f"project = {args.project or DEFAULT_PROJECT}"]
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


def search_issues(
    client: JiraCookieClient,
    jql: str,
    sprint_field: str,
    max_results: int,
    limit: Optional[int],
    log: Callable[[str], None] = print,
) -> List[Dict[str, Any]]:
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
            "expand": "names",
        }
        if token:
            params["nextPageToken"] = token
        data = client.get_json("/rest/api/3/search/jql", params=params)
        batch = data.get("issues") or []
        issues.extend(batch)
        log(f"Tickets trouves: {len(issues)}")
        if limit and len(issues) >= limit:
            issues = issues[:limit]
            break
        if data.get("isLast") is True:
            break
        token = data.get("nextPageToken")
        if not token:
            break
    return issues


def fetch_issue(client: JiraCookieClient, issue_key_or_id: str, sprint_field: str) -> Dict[str, Any]:
    fields = [
        "summary", "issuetype", "status", "assignee", "reporter", "created", "updated",
        "priority", sprint_field, "parent", "project",
    ]
    return client.get_json(f"/rest/api/3/issue/{quote(str(issue_key_or_id))}", params={"fields": ",".join(fields)})


def fetch_all_changelog(client: JiraCookieClient, issue_key_or_id: str) -> List[Dict[str, Any]]:
    histories: List[Dict[str, Any]] = []
    start_at = 0
    while True:
        try:
            data = client.get_json(
                f"/rest/api/3/issue/{quote(str(issue_key_or_id))}/changelog",
                params={"startAt": start_at, "maxResults": 100},
            )
        except Exception:
            # Fallback: expanded issue endpoint. Usually limited, but useful if changelog endpoint is blocked.
            data2 = client.get_json(f"/rest/api/3/issue/{quote(str(issue_key_or_id))}", params={"expand": "changelog"})
            changelog = data2.get("changelog") or {}
            return changelog.get("values") or changelog.get("histories") or []
        values = data.get("values") or data.get("histories") or []
        histories.extend(values)
        total = data.get("total")
        max_res = data.get("maxResults") or len(values) or 100
        if data.get("isLast") is True:
            break
        if total is not None and len(histories) >= int(total):
            break
        if not values:
            break
        start_at += int(max_res)
    return histories


def fetch_all_comments(client: JiraCookieClient, issue_key_or_id: str) -> List[Dict[str, Any]]:
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
        max_res = data.get("maxResults") or len(values) or 100
        if total is not None and len(comments) >= int(total):
            break
        if not values:
            break
        start_at += int(max_res)
    return comments


# =============================================================================
# Parse Jira entities
# =============================================================================


def user_display(user: Any) -> str:
    if not isinstance(user, dict):
        return ""
    return user.get("displayName") or user.get("emailAddress") or user.get("accountId") or ""


def parse_sprint_value(value: Any) -> Tuple[str, str]:
    if not value:
        return "", ""
    names: List[str] = []
    ids: List[str] = []
    items = value if isinstance(value, list) else [value]
    for item in items:
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


def parse_issue_row(issue: Dict[str, Any], base_url: str, sprint_field: str) -> IssueRow:
    fields = issue.get("fields") or {}
    status = fields.get("status") or {}
    status_cat = status.get("statusCategory") or {}
    issue_type = fields.get("issuetype") or {}
    assignee = fields.get("assignee")
    reporter = fields.get("reporter")
    priority = fields.get("priority") or {}
    project = fields.get("project") or {}
    parent = fields.get("parent") or {}
    sprint_names, sprint_ids = parse_sprint_value(fields.get(sprint_field))
    return IssueRow(
        issue_id=str(issue.get("id") or ""),
        key=str(issue.get("key") or ""),
        url=f"{base_url.rstrip('/')}/browse/{issue.get('key')}",
        summary=clean_text(fields.get("summary")),
        project_key=project.get("key") or "",
        project_name=project.get("name") or "",
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
    if field_id == "description" or norm == "description" or " description" in norm:
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
                from_excerpt=one_line(from_val, 500),
                to_excerpt=one_line(to_val, 500),
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
            to_excerpt=one_line(body, 500),
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
                to_excerpt=one_line(body, 500),
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
        churn = len(desc) * 4 + len(ac) * 5 + len(status) * 3 + len(assign) * 2 + len(sprint) * 2 + len(comments)
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
            "Churn score": churn,
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
    return [{"Champ": k, "Événements": v, "Tickets": len({e.issue_id for e in events if e.field == k})} for k, v in counter.most_common()]


def events_by_day_field(events: Sequence[TimelineEvent]) -> List[Dict[str, Any]]:
    """KPI croisé date x champ pour analyser une période dans Excel."""
    grouped: Dict[Tuple[str, str], List[TimelineEvent]] = defaultdict(list)
    for e in events:
        if e.event_date:
            day = e.event_date.astimezone(timezone.utc).date().isoformat()
        else:
            day = "(sans date)"
        grouped[(day, e.field)].append(e)
    rows: List[Dict[str, Any]] = []
    for (day, field), evs in sorted(grouped.items(), key=lambda kv: (kv[0][0], kv[0][1])):
        rows.append({
            "Date": day,
            "Champ": field,
            "Événements": len(evs),
            "Tickets concernés": count_unique(evs),
            "Auteurs": "; ".join(sorted({e.author for e in evs if e.author})),
        })
    return rows


def period_overview_rows(events: Sequence[TimelineEvent]) -> List[Dict[str, Any]]:
    """Vue globale statique: combien de changements par champ sur toute la période extraite."""
    by_field: Dict[str, List[TimelineEvent]] = defaultdict(list)
    for e in events:
        by_field[e.field].append(e)
    rows: List[Dict[str, Any]] = []
    for field, evs in sorted(by_field.items(), key=lambda kv: (-len(kv[1]), kv[0].lower())):
        dates = [e.event_date for e in evs if e.event_date]
        rows.append({
            "Champ": field,
            "Événements": len(evs),
            "Tickets concernés": count_unique(evs),
            "Premier changement": dt_excel(min(dates)) if dates else None,
            "Dernier changement": dt_excel(max(dates)) if dates else None,
            "Top auteur": Counter([e.author for e in evs if e.author]).most_common(1)[0][0] if any(e.author for e in evs) else "",
            "Changements forts": sum(1 for e in evs if change_magnitude(e.from_value, e.to_value) >= 25),
        })
    return rows


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


def row_from_issue(i: IssueRow) -> Dict[str, Any]:
    d = asdict(i)
    d["created"] = dt_excel(i.created)
    d["updated"] = dt_excel(i.updated)
    return d


# =============================================================================
# Excel output
# =============================================================================


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
        if any(x in h.lower() for x in ["résumé", "url", "excerpt", "value", "valeur", "lien"]):
            width = min(max(width, 24), 72)
        else:
            width = min(max(width, 10), 32)
        ws.set_column(c, c, width)


def style_palette(style: str) -> Dict[str, str]:
    if style == "audit":
        return {"dark": "1F2937", "primary": "2563EB", "accent": "F59E0B", "green": "10B981", "red": "EF4444", "light": "F3F4F6", "teal": "0F766E"}
    if style == "kanban":
        return {"dark": "0F172A", "primary": "7C3AED", "accent": "06B6D4", "green": "22C55E", "red": "F97316", "light": "F8FAFC", "teal": "14B8A6"}
    return {"dark": "17365D", "primary": "1F4E78", "accent": "F4B183", "green": "A9D18E", "red": "F8CBAD", "light": "EAF2F8", "teal": "9DC3E6"}


def build_formats(wb: Any, p: Dict[str, str]) -> Dict[str, Any]:
    return {
        "title": wb.add_format({"bold": True, "font_size": 18, "font_color": "white", "bg_color": "#" + p["dark"], "align": "center", "valign": "vcenter"}),
        "subtitle": wb.add_format({"bold": True, "font_size": 12, "font_color": "#" + p["dark"]}),
        "header": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#" + p["primary"], "border": 1, "align": "center", "valign": "vcenter"}),
        "card_label": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#" + p["primary"], "align": "center", "valign": "vcenter", "border": 1}),
        "card_value": wb.add_format({"bold": True, "font_size": 18, "font_color": "#" + p["dark"], "bg_color": "#" + p["light"], "align": "center", "valign": "vcenter", "border": 1}),
        "date": wb.add_format({"num_format": "yyyy-mm-dd hh:mm", "valign": "top"}),
        "number": wb.add_format({"num_format": "#,##0", "valign": "top"}),
        "bool": wb.add_format({"valign": "top"}),
        "note": wb.add_format({"italic": True, "font_color": "#666666"}),
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


def write_dashboard(wb: Any, fmt: Dict[str, Any], p: Dict[str, str], issues: List[IssueRow], events: List[TimelineEvent], by_day: List[Dict[str, Any]], by_field: List[Dict[str, Any]], ticket_rows: List[Dict[str, Any]]) -> None:
    ws = wb.add_worksheet("00_Dashboard")
    ws.hide_gridlines(2)
    ws.merge_range("A1:N2", "Jira changelog analysis - browser session", fmt["title"])
    ws.set_row(0, 26)
    cards = [
        ("Tickets", len(issues)),
        ("Événements", len(events)),
        ("Tickets avec Description", len({e.issue_id for e in events if e.field == "Description"})),
        ("Tickets avec AC", len({e.issue_id for e in events if e.field == "Acceptance Criteria"})),
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
    ws.write("A14", "Utilisez les onglets 01_Timeline et 02_Tickets avec les filtres Excel pour filtrer par Type, Sprint, Statut, Champ et Ticket.", fmt["note"])
    ws.write("A15", "L'interface HTML fournie permet une lecture plus visuelle ticket par ticket.", fmt["note"])

    start = 18
    ws.write(start, 0, "Champ", fmt["header"])
    ws.write(start, 1, "Événements", fmt["header"])
    for idx, row in enumerate(by_field[:10], start=start + 1):
        ws.write(idx, 0, row["Champ"])
        ws.write_number(idx, 1, row["Événements"])
    if by_field:
        chart = wb.add_chart({"type": "column"})
        chart.add_series({
            "name": "Événements",
            "categories": ["00_Dashboard", start + 1, 0, start + min(10, len(by_field)), 0],
            "values": ["00_Dashboard", start + 1, 1, start + min(10, len(by_field)), 1],
        })
        chart.set_title({"name": "Événements par champ"})
        chart.set_legend({"none": True})
        chart.set_size({"width": 620, "height": 320})
        ws.insert_chart("D18", chart)

    day_row = 34
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


def add_timeline_conditional_formats(ws: Any, nrows: int, fmt: Dict[str, Any]) -> None:
    field_col = 12  # field column in TimelineEvent.as_dict order
    ws.conditional_format(1, field_col, nrows, field_col, {"type": "text", "criteria": "containing", "value": "Description", "format": fmt["desc"]})
    ws.conditional_format(1, field_col, nrows, field_col, {"type": "text", "criteria": "containing", "value": "Acceptance", "format": fmt["ac"]})
    ws.conditional_format(1, field_col, nrows, field_col, {"type": "text", "criteria": "containing", "value": "Statut", "format": fmt["status"]})
    ws.conditional_format(1, field_col, nrows, field_col, {"type": "text", "criteria": "containing", "value": "Assignation", "format": fmt["assign"]})


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
    assign_rows = assignment_transitions(events)
    status_rows = status_transitions(events)
    by_day = events_by_day(events)
    by_field = events_by_field(events)
    by_day_field = events_by_day_field(events)
    period_overview = period_overview_rows(events)

    write_dashboard(wb, formats, palette, issues, events, by_day, by_field, ticket_rows)
    sheets = [
        ("00_Period_Overview", period_overview),
        ("01_Timeline", timeline_rows),
        ("02_Tickets", ticket_rows),
        ("03_Current_Issues", issue_rows),
        ("04_Desc_AC", desc_ac_rows),
        ("05_Comments", comment_rows),
        ("06_Assignations", assign_rows),
        ("07_Status", status_rows),
        ("08_By_Day", by_day),
        ("09_By_Field", by_field),
        ("10_By_Day_Field", by_day_field),
    ]
    for name, rows in sheets:
        ws = wb.add_worksheet(name)
        write_rows(ws, rows, formats, autofilter=True)
        if name == "01_Timeline" and timeline_rows:
            add_timeline_conditional_formats(ws, len(timeline_rows), formats)
    wb.set_properties({"title": "Jira API changelog analysis", "comments": "Generated from Jira REST API through browser session"})
    wb.close()
    validate_xlsx(path)


# =============================================================================
# HTML output
# =============================================================================


def tokenize_for_diff(text: Any) -> List[str]:
    text = clean_text(text)
    # Conserve espaces, mots accentues et ponctuation separement pour detecter meme une virgule.
    return re.findall(r"\s+|[A-Za-zÀ-ÖØ-öø-ÿ0-9_]+|[^\w\s]", text, flags=re.UNICODE)


def html_join_tokens(tokens: Sequence[str]) -> str:
    return "".join(html.escape(t) for t in tokens)


def build_side_snippet(tokens: Sequence[str], change_start: int, change_end: int, ctx_start: int, ctx_end: int, css_class: str) -> str:
    before = html_join_tokens(tokens[ctx_start:change_start])
    changed = html_join_tokens(tokens[change_start:change_end])
    after = html_join_tokens(tokens[change_end:ctx_end])
    if not changed:
        changed = "∅"
    return before + f"<mark class=\"{css_class}\">" + changed + "</mark>" + after


def precise_diff_html(old_value: Any, new_value: Any, context_tokens: int = 28, max_blocks: int = 4) -> str:
    """Retourne un HTML safe montrant exactement les zones modifiees.

    Le diff se fait au niveau token: mots, espaces et ponctuation. Une virgule ajoutee ou supprimee
    est donc visible.
    """
    old_tokens = tokenize_for_diff(old_value)
    new_tokens = tokenize_for_diff(new_value)
    if not old_tokens and not new_tokens:
        return ""
    if old_tokens == new_tokens:
        return '<div class="diff-unchanged">Aucune différence textuelle détectée après normalisation.</div>'

    matcher = difflib.SequenceMatcher(a=old_tokens, b=new_tokens, autojunk=False)
    blocks = [op for op in matcher.get_opcodes() if op[0] != "equal"][:max_blocks]
    old_parts: List[str] = []
    new_parts: List[str] = []
    summary_parts: List[str] = []

    for idx, (tag, i1, i2, j1, j2) in enumerate(blocks, start=1):
        old_ctx_start = max(0, i1 - context_tokens)
        old_ctx_end = min(len(old_tokens), i2 + context_tokens)
        new_ctx_start = max(0, j1 - context_tokens)
        new_ctx_end = min(len(new_tokens), j2 + context_tokens)
        if old_ctx_start > 0:
            old_prefix = "… "
        else:
            old_prefix = ""
        if old_ctx_end < len(old_tokens):
            old_suffix = " …"
        else:
            old_suffix = ""
        if new_ctx_start > 0:
            new_prefix = "… "
        else:
            new_prefix = ""
        if new_ctx_end < len(new_tokens):
            new_suffix = " …"
        else:
            new_suffix = ""

        old_html = old_prefix + build_side_snippet(old_tokens, i1, i2, old_ctx_start, old_ctx_end, "diff-del") + old_suffix
        new_html = new_prefix + build_side_snippet(new_tokens, j1, j2, new_ctx_start, new_ctx_end, "diff-add") + new_suffix
        old_parts.append(f'<div class="diff-block"><div class="diff-block-title">Bloc {idx}</div>{old_html}</div>')
        new_parts.append(f'<div class="diff-block"><div class="diff-block-title">Bloc {idx}</div>{new_html}</div>')
        summary_parts.append(f"{tag}: -{max(i2-i1,0)} token(s) / +{max(j2-j1,0)} token(s)")

    if len([op for op in matcher.get_opcodes() if op[0] != "equal"]) > max_blocks:
        summary_parts.append("Diff tronqué aux premiers blocs significatifs")

    return (
        '<div class="precise-diff">'
        '<div class="diff-summary">' + html.escape(" · ".join(summary_parts)) + '</div>'
        '<div class="diff-columns">'
        '<div class="diff-old"><h4>Avant — zones retirées/modifiées</h4>' + "".join(old_parts) + '</div>'
        '<div class="diff-new"><h4>Après — zones ajoutées/modifiées</h4>' + "".join(new_parts) + '</div>'
        '</div></div>'
    )


def text_diff_summary(old_value: Any, new_value: Any) -> str:
    old_tokens = tokenize_for_diff(old_value)
    new_tokens = tokenize_for_diff(new_value)
    if old_tokens == new_tokens:
        return "Aucune différence textuelle"
    matcher = difflib.SequenceMatcher(a=old_tokens, b=new_tokens, autojunk=False)
    inserted = deleted = replaced = 0
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "insert":
            inserted += j2 - j1
        elif tag == "delete":
            deleted += i2 - i1
        elif tag == "replace":
            replaced += max(i2 - i1, j2 - j1)
    return f"tokens ajoutés={inserted}; tokens supprimés={deleted}; tokens remplacés={replaced}"


def serialize_issue(i: IssueRow) -> Dict[str, Any]:
    d = asdict(i)
    d["created"] = dt_iso(i.created)
    d["updated"] = dt_iso(i.updated)
    return d


def serialize_event(e: TimelineEvent) -> Dict[str, Any]:
    """Serialize event for HTML.

    Important: keep a large value payload so the browser dashboard can highlight
    a one-character change (comma, accent, word, punctuation) instead of only
    showing before/after excerpts. Values are capped to avoid generating an
    unusable HTML file if Jira contains very large descriptions.
    """
    d = asdict(e)
    d["event_date"] = dt_iso(e.event_date)
    d["from_value"] = clean_text(e.from_value)[:20000]
    d["to_value"] = clean_text(e.to_value)[:20000]
    d["from_excerpt"] = one_line(e.from_value, 900)
    d["to_excerpt"] = one_line(e.to_value, 900)
    return d


def write_html(path: Path, issues: List[IssueRow], events: List[TimelineEvent], style: str = "executive") -> None:
    palette = style_palette(style)
    data_json = json.dumps({"issues": [serialize_issue(i) for i in issues], "events": [serialize_event(e) for e in sorted_events(events)]}, ensure_ascii=False)
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
.header-inner {{ display:flex; align-items:center; justify-content:space-between; gap:24px; max-width:1800px; margin:0 auto; }}
header h1 {{ margin:0; font-size:24px; }}
header p {{ margin:6px 0 0 0; opacity:.85; }}
.export-button {{ display:inline-flex; align-items:center; gap:9px; border:1px solid rgba(255,255,255,.24); border-radius:10px; padding:10px 15px; color:white; background:linear-gradient(135deg,var(--primary),var(--teal)); font:inherit; font-size:13px; font-weight:800; white-space:nowrap; cursor:pointer; box-shadow:0 8px 22px rgba(0,0,0,.18); }}
.export-button:hover {{ transform:translateY(-1px); filter:brightness(1.08); }}
.export-button svg {{ width:17px; height:17px; fill:none; stroke:currentColor; stroke-width:2; }}
.container {{ padding:22px 32px; }}
.filters {{ display:grid; grid-template-columns: repeat(8, minmax(135px,1fr)); gap:12px; background:white; padding:16px; border-radius:14px; box-shadow:0 1px 5px rgba(15,23,42,.10); position:sticky; top:0; z-index:5; }}
label {{ font-size:12px; font-weight:700; color:#374151; display:block; margin-bottom:4px; }}
select,input {{ width:100%; padding:9px; border:1px solid #cbd5e1; border-radius:8px; background:white; }}
.multi-filter {{ position:relative; }}
.multi-toggle {{ width:100%; min-height:37px; padding:9px 30px 9px 10px; border:1px solid #cbd5e1; border-radius:8px; background:white; color:#111827; text-align:left; cursor:pointer; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; position:relative; }}
.multi-toggle:after {{ content:"▾"; position:absolute; right:10px; color:#64748b; }}
.multi-filter.open .multi-toggle {{ border-color:var(--primary); box-shadow:0 0 0 2px #dbeafe; }}
.multi-options {{ display:none; position:absolute; top:calc(100% + 4px); left:0; min-width:100%; width:max-content; max-width:320px; max-height:280px; overflow:auto; padding:7px; background:white; border:1px solid #cbd5e1; border-radius:9px; box-shadow:0 10px 24px rgba(15,23,42,.18); z-index:30; }}
.multi-filter.open .multi-options {{ display:block; }}
.multi-reset {{ width:100%; padding:6px 8px; margin-bottom:5px; border:0; border-radius:6px; background:#eef6ff; color:var(--primary); font-weight:700; text-align:left; cursor:pointer; }}
.multi-option {{ display:flex; align-items:flex-start; gap:7px; min-width:180px; padding:6px 7px; margin:0; border-radius:6px; font-size:12px; font-weight:500; color:#1f2937; cursor:pointer; }}
.multi-option:hover {{ background:#f1f5f9; }}
.multi-option input {{ width:auto; margin:2px 0 0; padding:0; flex:0 0 auto; }}
.multi-empty {{ padding:7px; color:#64748b; font-size:12px; }}
.cards {{ display:grid; grid-template-columns: repeat(8, 1fr); gap:14px; margin:18px 0; }}
.card {{ background:white; border-left:6px solid var(--primary); border-radius:14px; padding:14px; box-shadow:0 1px 5px rgba(15,23,42,.10); }}
.card .label {{ font-size:12px; color:#64748b; font-weight:700; }}
.card .value {{ font-size:26px; font-weight:800; color:var(--dark); margin-top:4px; }}
.global {{ display:grid; grid-template-columns: 1.2fr 1fr; gap:18px; margin-bottom:18px; }}
.grid {{ display:grid; grid-template-columns: 380px 1fr; gap:18px; align-items:start; }}
.panel {{ background:white; border-radius:14px; padding:16px; box-shadow:0 1px 5px rgba(15,23,42,.10); }}
.panel h2 {{ margin:0 0 12px 0; color:var(--dark); font-size:18px; }}
.ticket {{ border:1px solid #e5e7eb; border-radius:10px; padding:10px; margin-bottom:8px; cursor:pointer; }}
.ticket:hover {{ border-color:var(--primary); background:#f8fbff; }}
.ticket.active {{ border-color:var(--accent); background:#fffbeb; }}
.ticket .key {{ font-weight:800; color:var(--primary); }}
.ticket .meta {{ font-size:12px; color:#64748b; margin-top:4px; }}
.ticket-section {{ margin-bottom:22px; border:1px solid #e5e7eb; border-radius:14px; overflow:hidden; }}
.ticket-section h3 {{ margin:0; padding:12px 14px; background:#eef6ff; color:var(--dark); font-size:16px; }}
.ticket-section .ticket-meta {{ padding:0 14px 12px 14px; color:#64748b; font-size:12px; background:#eef6ff; }}
.timeline {{ position:relative; margin-left:12px; border-left:3px solid #dbeafe; padding:14px 12px 4px 18px; }}
.event {{ position:relative; margin:0 0 14px 0; padding:12px; border-radius:12px; background:#f8fafc; border:1px solid #e5e7eb; }}
.event:before {{ content:""; position:absolute; left:-27px; top:16px; width:14px; height:14px; border-radius:50%; background:var(--primary); border:3px solid white; box-shadow:0 0 0 2px #bfdbfe; }}
.event.desc:before {{ background:var(--green); }}
.event.ac:before {{ background:var(--accent); }}
.event.status:before {{ background:var(--teal); }}
.event.assign:before {{ background:var(--red); }}
.event.comment:before {{ background:#64748b; }}
.event .top {{ display:flex; justify-content:space-between; gap:12px; }}
.event .field {{ font-weight:800; color:var(--dark); }}
.event .date {{ color:#64748b; font-size:12px; white-space:nowrap; }}
.precise {{ margin-top:10px; background:white; border:1px solid #e5e7eb; border-radius:8px; padding:10px; white-space:pre-wrap; line-height:1.55; max-height:260px; overflow:auto; font-size:13px; }}
.precise-title {{ margin-top:10px; font-size:12px; font-weight:800; color:#334155; }}
.diff {{ display:grid; grid-template-columns: 1fr 1fr; gap:10px; margin-top:10px; }}
.diff div {{ background:white; border-radius:8px; padding:8px; border:1px solid #e5e7eb; font-size:12px; white-space:pre-wrap; max-height:160px; overflow:auto; }}
ins {{ background:#dcfce7; color:#166534; text-decoration:none; border-radius:3px; padding:0 2px; font-weight:700; }}
del {{ background:#fee2e2; color:#991b1b; text-decoration:line-through; border-radius:3px; padding:0 2px; font-weight:700; }}
.context {{ color:#475569; }}
.badge {{ display:inline-block; padding:3px 8px; border-radius:999px; background:#e0f2fe; color:#075985; font-size:11px; font-weight:700; margin-right:4px; }}
.bars {{ margin-top:12px; }}
.bar {{ display:grid; grid-template-columns: 170px 1fr 50px; gap:8px; align-items:center; margin:6px 0; font-size:12px; }}
.bar span:nth-child(2) {{ height:12px; background:linear-gradient(90deg,var(--primary),var(--teal)); border-radius:999px; }}
.small-note {{ color:#64748b; font-size:12px; margin-top:8px; }}
.export-dialog {{ width:min(720px,calc(100vw - 32px)); padding:0; border:0; border-radius:18px; color:#111827; box-shadow:0 28px 70px rgba(15,23,42,.28); }}
.export-dialog::backdrop {{ background:rgba(15,23,42,.62); backdrop-filter:blur(3px); }}
.export-dialog-head {{ padding:22px 24px 16px; background:linear-gradient(135deg,var(--dark),var(--primary)); color:white; }}
.export-dialog-head h2 {{ margin:0; font-size:21px; }}
.export-dialog-head p {{ margin:7px 0 0; color:rgba(255,255,255,.8); font-size:13px; line-height:1.45; }}
.export-dialog-body {{ padding:20px 24px; }}
.export-scope {{ display:flex; gap:10px; align-items:center; margin-bottom:17px; padding:11px 13px; border:1px solid #bfdbfe; border-radius:10px; background:#eff6ff; color:#1e3a8a; font-size:13px; }}
.export-scope strong {{ font-size:16px; }}
.export-label {{ margin:0 0 6px; font-size:12px; font-weight:800; color:#334155; }}
.export-title-input {{ width:100%; margin-bottom:17px; }}
.export-options {{ display:grid; grid-template-columns:1fr 1fr; gap:9px; margin-bottom:17px; }}
.export-option {{ display:flex; gap:10px; align-items:flex-start; padding:11px; border:1px solid #e2e8f0; border-radius:10px; cursor:pointer; font-size:13px; font-weight:700; }}
.export-option:hover {{ background:#f8fafc; border-color:#93c5fd; }}
.export-option input {{ width:auto; margin:2px 0 0; accent-color:var(--primary); }}
.export-option small {{ display:block; margin-top:3px; color:#64748b; font-weight:400; line-height:1.35; }}
.export-row {{ display:grid; grid-template-columns:1fr 1fr; gap:12px; }}
.export-help {{ margin:15px 0 0; color:#64748b; font-size:12px; line-height:1.45; }}
.export-actions {{ display:flex; justify-content:flex-end; gap:10px; padding:15px 24px; border-top:1px solid #e2e8f0; background:#f8fafc; }}
.dialog-button {{ border:0; border-radius:9px; padding:10px 15px; font:inherit; font-size:13px; font-weight:800; cursor:pointer; }}
.dialog-button.secondary {{ background:white; color:#475569; border:1px solid #cbd5e1; }}
.dialog-button.primary {{ background:var(--primary); color:white; }}
.print-report {{ display:none; }}
@media(max-width:1200px) {{ .filters,.cards {{ grid-template-columns: repeat(2, 1fr); }} .grid,.global {{ grid-template-columns:1fr; }} }}
@media(max-width:700px) {{ header {{ padding:18px; }} .header-inner {{ align-items:flex-start; }} .export-button span {{ display:none; }} .container {{ padding:16px; }} .export-options,.export-row {{ grid-template-columns:1fr; }} }}
@page {{ size:A4; margin:12mm 11mm 15mm; }}
@media print {{
  html,body {{ width:210mm; background:white !important; color:#172033; font-family:"Segoe UI",Arial,sans-serif; print-color-adjust:exact; -webkit-print-color-adjust:exact; }}
  body > header,body > .container,body > .export-dialog {{ display:none !important; }}
  .print-report {{ display:block !important; width:100%; }}
  .report-cover {{ min-height:253mm; display:flex; flex-direction:column; padding:12mm 10mm 9mm; color:white; background:linear-gradient(145deg,var(--dark) 0 62%,var(--primary) 62% 82%,var(--teal) 82%); break-after:page; }}
  .report-brand {{ display:flex; justify-content:space-between; align-items:center; font-size:9pt; font-weight:700; letter-spacing:.08em; text-transform:uppercase; opacity:.85; }}
  .report-cover-main {{ margin:auto 0; max-width:155mm; }}
  .report-kicker {{ margin-bottom:6mm; color:#bfdbfe; font-size:11pt; font-weight:800; letter-spacing:.12em; text-transform:uppercase; }}
  .report-cover h1 {{ margin:0 0 7mm; font-size:31pt; line-height:1.08; color:white; }}
  .report-cover-subtitle {{ margin:0; font-size:14pt; line-height:1.5; color:#e2e8f0; }}
  .report-cover-meta {{ display:grid; grid-template-columns:repeat(3,1fr); gap:5mm; margin-top:15mm; }}
  .report-cover-stat {{ padding:5mm; border:1px solid rgba(255,255,255,.25); border-radius:4mm; background:rgba(255,255,255,.10); }}
  .report-cover-stat b {{ display:block; font-size:23pt; color:white; }}
  .report-cover-stat span {{ font-size:9pt; color:#dbeafe; }}
  .report-cover-foot {{ display:flex; justify-content:space-between; gap:8mm; font-size:8.5pt; color:#dbeafe; }}
  .report-page {{ break-before:page; }}
  .report-page:first-of-type {{ break-before:auto; }}
  .report-section {{ margin:0 0 9mm; }}
  .report-section h2 {{ margin:0 0 5mm; padding-bottom:3mm; border-bottom:1.5px solid #cbd5e1; color:var(--dark); font-size:18pt; }}
  .report-section h3 {{ margin:0 0 3mm; color:#334155; font-size:11pt; }}
  .report-section-intro {{ margin:-2mm 0 5mm; color:#64748b; font-size:9pt; line-height:1.5; }}
  .report-filters {{ display:flex; flex-wrap:wrap; gap:2mm; margin:0 0 7mm; }}
  .report-chip {{ padding:1.6mm 3mm; border-radius:999px; background:#eaf2f8; color:var(--primary); font-size:8pt; font-weight:700; }}
  .report-kpis {{ display:grid; grid-template-columns:repeat(3,1fr); gap:4mm; }}
  .report-kpi {{ position:relative; overflow:hidden; min-height:29mm; padding:5mm; border:1px solid #dbe4ef; border-radius:4mm; background:#fff; }}
  .report-kpi:after {{ content:""; position:absolute; right:-8mm; bottom:-9mm; width:24mm; height:24mm; border-radius:50%; background:var(--kpi-color,#dbeafe); opacity:.22; }}
  .report-kpi-label {{ color:#64748b; font-size:8pt; font-weight:800; letter-spacing:.04em; text-transform:uppercase; }}
  .report-kpi-value {{ margin-top:2mm; color:var(--dark); font-size:24pt; font-weight:850; line-height:1; }}
  .report-kpi-note {{ margin-top:2mm; color:#64748b; font-size:7.5pt; }}
  .report-highlight {{ margin:6mm 0; padding:5mm 6mm; border-left:4px solid var(--accent); border-radius:0 3mm 3mm 0; background:#fff7ed; color:#334155; font-size:10pt; line-height:1.55; }}
  .report-two-cols {{ display:grid; grid-template-columns:1.35fr .9fr; gap:7mm; align-items:start; }}
  .report-chart-card {{ padding:5mm; border:1px solid #dbe4ef; border-radius:4mm; background:#fff; break-inside:avoid; }}
  .report-chart-card svg {{ display:block; width:100%; height:auto; overflow:visible; }}
  .report-legend {{ display:grid; gap:2.2mm; margin-top:4mm; }}
  .report-legend-row {{ display:grid; grid-template-columns:3mm 1fr auto; gap:2mm; align-items:center; font-size:8pt; }}
  .report-legend-dot {{ width:3mm; height:3mm; border-radius:1mm; }}
  .report-author {{ display:grid; grid-template-columns:35mm 1fr 9mm; gap:3mm; align-items:center; margin:2.5mm 0; font-size:8pt; }}
  .report-author-name {{ overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-weight:700; }}
  .report-author-track {{ height:3.2mm; overflow:hidden; border-radius:99px; background:#e2e8f0; }}
  .report-author-bar {{ height:100%; border-radius:99px; background:linear-gradient(90deg,var(--primary),var(--teal)); }}
  .report-table {{ width:100%; border-collapse:collapse; font-size:7.5pt; }}
  .report-table thead {{ display:table-header-group; }}
  .report-table tr {{ break-inside:avoid; }}
  .report-table th {{ padding:2.5mm 2mm; background:var(--dark); color:white; text-align:left; font-size:7pt; text-transform:uppercase; letter-spacing:.03em; }}
  .report-table td {{ padding:2.5mm 2mm; border-bottom:1px solid #e2e8f0; vertical-align:top; }}
  .report-table tr:nth-child(even) td {{ background:#f8fafc; }}
  .report-ticket-key {{ color:var(--primary); font-weight:850; }}
  .report-change {{ margin:0 0 4mm; padding:4mm 4.5mm; border:1px solid #dbe4ef; border-left:3px solid var(--change-color,var(--primary)); border-radius:3mm; background:#fff; break-inside:avoid; }}
  .report-change-head {{ display:flex; justify-content:space-between; gap:5mm; margin-bottom:2mm; }}
  .report-change-title {{ font-size:9pt; font-weight:850; color:var(--dark); }}
  .report-change-meta {{ color:#64748b; font-size:7.5pt; white-space:nowrap; }}
  .report-change-type {{ margin-bottom:2.5mm; color:#475569; font-size:8pt; }}
  .report-precise-head {{ display:flex; align-items:center; justify-content:space-between; gap:3mm; margin:0 0 1.5mm; }}
  .report-precise-title {{ color:#334155; font-size:7.2pt; font-weight:850; letter-spacing:.04em; text-transform:uppercase; }}
  .report-diff-legend {{ display:flex; gap:1.5mm; color:#64748b; font-size:6.5pt; }}
  .report-diff-key {{ padding:1mm 1.8mm; border-radius:99px; font-weight:750; }}
  .report-diff-key.deleted {{ background:#fee2e2; color:#991b1b; text-decoration:line-through; }}
  .report-diff-key.added {{ background:#dcfce7; color:#166534; }}
  .report-precise-diff {{ margin-bottom:3mm; padding:3mm; border:1px solid #cbd5e1; border-radius:2mm; background:#f8fafc; color:#475569; font-size:7.5pt; line-height:1.55; white-space:pre-wrap; overflow-wrap:anywhere; }}
  .report-precise-diff del {{ padding:.3mm .6mm; background:#fee2e2; color:#991b1b; text-decoration:line-through; font-weight:800; }}
  .report-precise-diff ins {{ padding:.3mm .6mm; background:#dcfce7; color:#166534; text-decoration:none; font-weight:800; }}
  .report-before-after {{ display:grid; grid-template-columns:1fr 1fr; gap:3mm; }}
  .report-value {{ min-height:13mm; padding:3mm; border-radius:2mm; font-size:7.2pt; line-height:1.4; white-space:pre-wrap; overflow-wrap:anywhere; }}
  .report-value.before {{ background:#fff1f2; border:1px solid #fecdd3; }}
  .report-value.after {{ background:#f0fdf4; border:1px solid #bbf7d0; }}
  .report-value b {{ display:block; margin-bottom:1mm; font-size:6.8pt; letter-spacing:.04em; text-transform:uppercase; }}
  .report-empty {{ padding:8mm; border:1px dashed #cbd5e1; border-radius:3mm; color:#64748b; text-align:center; }}
  .report-footer {{ position:fixed; right:0; bottom:-9mm; left:0; display:flex; justify-content:space-between; padding-top:2mm; border-top:1px solid #cbd5e1; color:#64748b; font-size:7pt; }}
}}
</style>
</head>
<body>
<header>
  <div class="header-inner">
    <div><h1>Jira changelog timeline</h1><p>Mode cookies navigateur · vision globale par période + diff précis des champs texte</p></div>
    <button type="button" class="export-button" id="exportButton" aria-haspopup="dialog">
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12m0 0 4-4m-4 4-4-4M5 15v4a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-4"/></svg>
      <span>Exporter en PDF</span>
    </button>
  </div>
</header>
<dialog class="export-dialog" id="exportDialog">
  <div class="export-dialog-head"><h2>Composer le rapport PDF</h2><p>Le rapport reprend la période, les filtres et les tickets actuellement sélectionnés dans le dashboard.</p></div>
  <div class="export-dialog-body">
    <div class="export-scope" id="exportScope"></div>
    <div class="export-label">Titre du rapport</div>
    <input class="export-title-input" id="reportTitle" value="Rapport d’activité Jira" maxlength="100"/>
    <div class="export-label">Sections à inclure</div>
    <div class="export-options">
      <label class="export-option"><input type="checkbox" data-report-section="summary" checked/><span>Synthèse exécutive<small>Indicateurs clés et lecture rapide.</small></span></label>
      <label class="export-option"><input type="checkbox" data-report-section="activity" checked/><span>Activité dans le temps<small>Graphique d’évolution des modifications.</small></span></label>
      <label class="export-option"><input type="checkbox" data-report-section="fields" checked/><span>Répartition des actions<small>Champs modifiés et contributeurs.</small></span></label>
      <label class="export-option"><input type="checkbox" data-report-section="tickets" checked/><span>Tableau des tickets<small>Volume et nature des actions par ticket.</small></span></label>
      <label class="export-option"><input type="checkbox" data-report-section="details" checked/><span>Détail des modifications<small>Qui a fait quoi, quand, avant et après.</small></span></label>
    </div>
    <div class="export-row">
      <div><div class="export-label">Nombre maximal de détails</div><select id="reportDetailLimit"><option value="25">25 modifications</option><option value="50" selected>50 modifications</option><option value="100">100 modifications</option><option value="all">Toutes les modifications</option></select></div>
      <div><div class="export-label">Ordre du détail</div><select id="reportDetailOrder"><option value="newest">Plus récentes d’abord</option><option value="oldest">Plus anciennes d’abord</option><option value="ticket">Regrouper par ticket</option></select></div>
    </div>
    <p class="export-help">La fenêtre d’impression du navigateur s’ouvrira ensuite : choisissez <b>Enregistrer au format PDF</b> comme destination.</p>
  </div>
  <div class="export-actions"><button type="button" class="dialog-button secondary" id="cancelExport">Annuler</button><button type="button" class="dialog-button primary" id="generatePdf">Créer le PDF</button></div>
</dialog>
<div class="container">
  <div class="filters">
    <div><label>Date début</label><input type="date" id="dateStart"/></div>
    <div><label>Date fin</label><input type="date" id="dateEnd"/></div>
    <div><label>Type</label><div class="multi-filter" id="typeFilter"></div></div>
    <div><label>Sprint</label><div class="multi-filter" id="sprintFilter"></div></div>
    <div><label>Statut actuel</label><div class="multi-filter" id="statusFilter"></div></div>
    <div><label>Champ modifié</label><div class="multi-filter" id="fieldFilter"></div></div>
    <div><label>Ticket</label><div class="multi-filter" id="ticketFilter"></div></div>
    <div><label>Recherche</label><input id="searchBox" placeholder="PPMG, texte, auteur..."/></div>
  </div>
  <div class="cards" id="cards"></div>
  <div class="global">
    <div class="panel"><h2>Vue globale sur la période</h2><div id="globalPanel"></div></div>
    <div class="panel"><h2>Champs les plus modifiés</h2><div id="fieldPanel"></div></div>
  </div>
  <div class="grid">
    <div class="panel"><h2>Tickets</h2><div id="ticketList"></div></div>
    <div class="panel"><h2>Timeline ticket par ticket</h2><div id="timeline"></div></div>
  </div>
</div>
<main class="print-report" id="pdfReport" aria-hidden="true"></main>
<script>
const DATA = {data_json};
const issues = DATA.issues;
const events = DATA.events;
const state = {{ type:new Set(), sprint:new Set(), status:new Set(), field:new Set(), ticket:new Set(), q:'', start:'', end:'' }};
function uniq(arr) {{ return Array.from(new Set(arr.filter(x => x && String(x).trim()))).sort((a,b)=>String(a).localeCompare(String(b))) }}
function multiLabel(key) {{
  const selected = state[key];
  if(!selected.size) return 'Tous';
  if(selected.size === 1) return selected.values().next().value;
  return `${{selected.size}} sélectionnés`;
}}
function syncMultiSelect(id) {{
  const el=document.getElementById(id), key=id.replace('Filter','');
  el.querySelector('.multi-toggle').textContent=multiLabel(key);
  el.querySelectorAll('input[type="checkbox"]').forEach(cb=>cb.checked=state[key].has(cb.value));
}}
function fillMultiSelect(id, values) {{
  const el=document.getElementById(id), key=id.replace('Filter','');
  const toggle=document.createElement('button');
  toggle.type='button'; toggle.className='multi-toggle';
  toggle.setAttribute('aria-haspopup','true'); toggle.setAttribute('aria-expanded','false');
  const options=document.createElement('div'); options.className='multi-options';
  const reset=document.createElement('button');
  reset.type='button'; reset.className='multi-reset'; reset.textContent='Afficher tous';
  reset.onclick=e=>{{ e.stopPropagation(); state[key].clear(); syncMultiSelect(id); render(); }};
  options.appendChild(reset);
  if(!values.length) {{
    const empty=document.createElement('div'); empty.className='multi-empty'; empty.textContent='Aucune valeur'; options.appendChild(empty);
  }} else {{
    values.forEach(v=>{{
      const label=document.createElement('label'); label.className='multi-option';
      const cb=document.createElement('input'); cb.type='checkbox'; cb.value=v;
      const text=document.createElement('span'); text.textContent=v;
      cb.onchange=()=>{{ cb.checked ? state[key].add(v) : state[key].delete(v); syncMultiSelect(id); render(); }};
      label.append(cb,text); options.appendChild(label);
    }});
  }}
  toggle.onclick=e=>{{
    e.stopPropagation();
    document.querySelectorAll('.multi-filter.open').forEach(other=>{{ if(other!==el) {{ other.classList.remove('open'); other.querySelector('.multi-toggle').setAttribute('aria-expanded','false'); }} }});
    el.classList.toggle('open'); toggle.setAttribute('aria-expanded',String(el.classList.contains('open')));
  }};
  options.onclick=e=>e.stopPropagation();
  el.replaceChildren(toggle,options); syncMultiSelect(id);
}}
function toggleMultiValue(id, value) {{
  const key=id.replace('Filter','');
  state[key].has(value) ? state[key].delete(value) : state[key].add(value);
  syncMultiSelect(id); render();
}}
function init() {{
  fillMultiSelect('typeFilter', uniq(issues.map(i=>i.issue_type)));
  fillMultiSelect('sprintFilter', uniq(issues.flatMap(i=>String(i.sprint_names||'').split('; ').filter(Boolean))));
  fillMultiSelect('statusFilter', uniq(issues.map(i=>i.status)));
  fillMultiSelect('fieldFilter', uniq(events.map(e=>e.field)));
  fillMultiSelect('ticketFilter', uniq(issues.map(i=>i.key)));
  document.addEventListener('click',()=>document.querySelectorAll('.multi-filter.open').forEach(el=>{{ el.classList.remove('open'); el.querySelector('.multi-toggle').setAttribute('aria-expanded','false'); }}));
  document.getElementById('searchBox').oninput = e => {{ state.q=e.target.value.toLowerCase(); render(); }};
  document.getElementById('dateStart').onchange = e => {{ state.start=e.target.value; render(); }};
  document.getElementById('dateEnd').onchange = e => {{ state.end=e.target.value; render(); }};
  document.getElementById('exportButton').onclick = openExportDialog;
  document.getElementById('cancelExport').onclick = closeExportDialog;
  document.getElementById('generatePdf').onclick = generatePdf;
  prefillDates();
  render();
}}
function prefillDates() {{
  const dates = events.map(e=>e.event_date).filter(Boolean).sort();
  if(dates.length) {{
    document.getElementById('dateStart').value = dates[0].slice(0,10); state.start = dates[0].slice(0,10);
    document.getElementById('dateEnd').value = dates[dates.length-1].slice(0,10); state.end = dates[dates.length-1].slice(0,10);
  }}
}}
function datePass(s) {{
  if(!s) return true;
  const d = new Date(s);
  if(state.start && d < new Date(state.start+'T00:00:00')) return false;
  if(state.end && d > new Date(state.end+'T23:59:59')) return false;
  return true;
}}
function issuePass(i) {{
  if(state.type.size && !state.type.has(i.issue_type)) return false;
  if(state.sprint.size && !String(i.sprint_names||'').split('; ').some(s=>state.sprint.has(s))) return false;
  if(state.status.size && !state.status.has(i.status)) return false;
  if(state.ticket.size && !state.ticket.has(i.key)) return false;
  const q = state.q;
  if(q && !(String(i.key+' '+i.summary+' '+i.assignee+' '+i.sprint_names).toLowerCase().includes(q))) return false;
  return true;
}}
function eventPass(e, issueSet) {{
  if(!issueSet.has(e.issue_id)) return false;
  if(!datePass(e.event_date)) return false;
  if(state.field.size && !state.field.has(e.field)) return false;
  const q = state.q;
  if(q && !(String(e.key+' '+e.summary+' '+e.author+' '+e.field+' '+e.from_excerpt+' '+e.to_excerpt+' '+e.from_value+' '+e.to_value).toLowerCase().includes(q))) return false;
  return true;
}}
function getFilteredData() {{
  const filteredIssues = issues.filter(issuePass);
  const issueSet = new Set(filteredIssues.map(i=>i.issue_id));
  const filteredEvents = events.filter(e=>eventPass(e, issueSet));
  return {{ filteredIssues, filteredEvents }};
}}
function render() {{
  const {{filteredIssues, filteredEvents}} = getFilteredData();
  renderCards(filteredIssues, filteredEvents);
  renderGlobal(filteredIssues, filteredEvents);
  renderFields(filteredEvents);
  renderTickets(filteredIssues, filteredEvents);
  renderTimeline(filteredIssues, filteredEvents);
}}
function renderCards(is, es) {{
  const descTickets = new Set(es.filter(e=>e.field==='Description').map(e=>e.issue_id)).size;
  const acTickets = new Set(es.filter(e=>e.field==='Acceptance Criteria').map(e=>e.issue_id)).size;
  const cards = [
    ['Tickets', is.length], ['Événements', es.length], ['Tickets Description', descTickets], ['Tickets AC', acTickets],
    ['Modifs AC', es.filter(e=>e.field==='Acceptance Criteria').length], ['Modifs Description', es.filter(e=>e.field==='Description').length],
    ['Réassignations', es.filter(e=>e.field==='Assignation').length], ['Statuts', es.filter(e=>e.field==='Statut').length]
  ];
  document.getElementById('cards').innerHTML = cards.map(c=>`<div class="card"><div class="label">${{c[0]}}</div><div class="value">${{c[1]}}</div></div>`).join('');
}}
function renderGlobal(is, es) {{
  const days={{}}; es.forEach(e=>{{ const k=(e.event_date||'').slice(0,10)||'Sans date'; days[k]=(days[k]||0)+1; }});
  const max=Math.max(...Object.values(days),1);
  const bars=Object.entries(days).sort((a,b)=>a[0].localeCompare(b[0])).map(([k,v])=>`<div class="bar"><span>${{escapeHtml(k)}}</span><span style="width:${{Math.max(4,Math.round(v/max*100))}}%"></span><b>${{v}}</b></div>`).join('');
  document.getElementById('globalPanel').innerHTML = `<div class="small-note">Période sélectionnée : <b>${{state.start||'début'}}</b> → <b>${{state.end||'fin'}}</b></div>${{bars||'<p>Aucun événement.</p>'}}`;
}}
function renderFields(es) {{
  const c={{}}; es.forEach(e=>c[e.field]=(c[e.field]||0)+1); const max=Math.max(...Object.values(c),1);
  document.getElementById('fieldPanel').innerHTML = Object.entries(c).sort((a,b)=>b[1]-a[1]).map(([k,v])=>`<div class="bar"><span>${{escapeHtml(k)}}</span><span style="width:${{Math.max(4,Math.round(v/max*100))}}%"></span><b>${{v}}</b></div>`).join('') || '<p>Aucun événement.</p>';
}}
function renderTickets(is, es) {{
  const counts = new Map(); es.forEach(e=>counts.set(e.issue_id,(counts.get(e.issue_id)||0)+1));
  const sorted = [...is].sort((a,b)=>(counts.get(b.issue_id)||0)-(counts.get(a.issue_id)||0));
  document.getElementById('ticketList').innerHTML = sorted.slice(0,200).map(i=>`<div class="ticket ${{state.ticket.has(i.key)?'active':''}}" data-ticket-key="${{escapeHtml(i.key)}}"><span class="key">${{escapeHtml(i.key)}}</span> <span class="badge">${{escapeHtml(i.issue_type)}}</span><div>${{escapeHtml(i.summary||'')}}</div><div class="meta">${{escapeHtml(i.status||'')}} · ${{escapeHtml(i.assignee||'')}} · ${{escapeHtml(i.sprint_names||'')}} · ${{counts.get(i.issue_id)||0}} event(s)</div></div>`).join('');
  document.querySelectorAll('#ticketList .ticket').forEach(el=>el.onclick=()=>toggleMultiValue('ticketFilter',el.dataset.ticketKey));
}}
function eventClass(e) {{ if(e.field==='Description') return 'desc'; if(e.field==='Acceptance Criteria') return 'ac'; if(e.field==='Statut') return 'status'; if(e.field==='Assignation') return 'assign'; if(e.field==='Commentaire') return 'comment'; return ''; }}
function renderTimeline(is, es) {{
  const byIssue={{}}; es.forEach(e=>{{ if(!byIssue[e.issue_id]) byIssue[e.issue_id]=[]; byIssue[e.issue_id].push(e); }});
  const issueMap = Object.fromEntries(is.map(i=>[i.issue_id,i]));
  const issueIds = Object.keys(byIssue).sort((a,b)=> byIssue[b].length - byIssue[a].length).slice(0,80);
  if(!issueIds.length) {{ document.getElementById('timeline').innerHTML='<p>Aucun événement pour les filtres sélectionnés.</p>'; return; }}
  document.getElementById('timeline').innerHTML = issueIds.map(id=>{{
    const i=issueMap[id]||{{key:id,summary:'',status:'',assignee:'',sprint_names:''}};
    const evs=byIssue[id].sort((a,b)=>String(b.event_date).localeCompare(String(a.event_date))).slice(0,100);
    return `<section class="ticket-section"><h3>${{escapeHtml(i.key)}} · ${{escapeHtml(i.summary||'')}}</h3><div class="ticket-meta">${{escapeHtml(i.status||'')}} · ${{escapeHtml(i.assignee||'')}} · ${{escapeHtml(i.sprint_names||'')}} · ${{evs.length}} événement(s)</div><div class="timeline">` + evs.map(eventHtml).join('') + `</div></section>`;
  }}).join('');
}}
function eventHtml(e) {{
  const precise = (e.field==='Description' || e.field==='Acceptance Criteria' || e.field==='Résumé' || e.field==='Commentaire') ? `<div class="precise-title">Diff précis</div><div class="precise">${{preciseDiff(e.from_value||'', e.to_value||'')}}</div>` : '';
  const beforeAfter = `<div class="diff"><div><b>Avant</b><br>${{escapeHtml(e.from_excerpt||e.from_value||'')}}</div><div><b>Après</b><br>${{escapeHtml(e.to_excerpt||e.to_value||'')}}</div></div>`;
  return `<div class="event ${{eventClass(e)}}"><div class="top"><div><span class="field">${{escapeHtml(e.field)}}</span> <span class="badge">${{escapeHtml(e.key)}}</span> <span class="badge">${{escapeHtml(e.author||'')}}</span></div><div class="date">${{formatDate(e.event_date)}}</div></div><div>${{escapeHtml(e.change_type)}} · <a href="${{e.url}}" target="_blank">ouvrir Jira</a></div>${{precise}}${{beforeAfter}}</div>`;
}}
function closeExportDialog() {{
  const dialog=document.getElementById('exportDialog');
  if(typeof dialog.close==='function') dialog.close(); else dialog.removeAttribute('open');
}}
function openExportDialog() {{
  const {{filteredIssues,filteredEvents}}=getFilteredData();
  const modifications=filteredEvents.filter(e=>e.event_source==='changelog').length;
  document.getElementById('exportScope').innerHTML=`<span>Le PDF portera sur</span><strong>${{filteredIssues.length}} ticket(s)</strong><span>et</span><strong>${{modifications}} modification(s)</strong>`;
  const dialog=document.getElementById('exportDialog');
  if(typeof dialog.showModal==='function') dialog.showModal(); else dialog.setAttribute('open','');
}}
function selectedReportSections() {{
  return new Set(Array.from(document.querySelectorAll('[data-report-section]:checked')).map(el=>el.dataset.reportSection));
}}
function countBy(items,keyFn) {{
  const counts={{}};
  items.forEach(item=>{{ const key=keyFn(item)||'Non renseigné'; counts[key]=(counts[key]||0)+1; }});
  return counts;
}}
function sortedCounts(counts) {{ return Object.entries(counts).sort((a,b)=>b[1]-a[1] || String(a[0]).localeCompare(String(b[0]))); }}
function truncateReportText(value,max=420) {{
  const text=String(value||'').replace(/\s+/g,' ').trim();
  return text.length>max ? text.slice(0,max-1)+'…' : text;
}}
function reportFilterChips() {{
  const chips=[];
  const add=(label,value)=>{{ if(value) chips.push(`<span class="report-chip">${{escapeHtml(label)}} : ${{escapeHtml(value)}}</span>`); }};
  add('Période',`${{state.start||'début'}} → ${{state.end||'fin'}}`);
  [['Type','type'],['Sprint','sprint'],['Statut','status'],['Champ','field'],['Ticket','ticket']].forEach(([label,key])=>{{ if(state[key].size) add(label,Array.from(state[key]).join(', ')); }});
  if(state.q) add('Recherche',document.getElementById('searchBox').value);
  if(chips.length===1) add('Périmètre','Tous les éléments du dashboard');
  return chips.join('');
}}
function reportKpi(label,value,note,color) {{
  return `<div class="report-kpi" style="--kpi-color:${{color}}"><div class="report-kpi-label">${{escapeHtml(label)}}</div><div class="report-kpi-value">${{value}}</div><div class="report-kpi-note">${{escapeHtml(note)}}</div></div>`;
}}
function executiveText(is,es) {{
  const modifications=es.filter(e=>e.event_source==='changelog');
  const fields=sortedCounts(countBy(modifications,e=>e.field));
  const authors=sortedCounts(countBy(es,e=>e.author));
  const ticketCounts=sortedCounts(countBy(es,e=>e.key));
  if(!es.length) return 'Aucune action ne correspond aux filtres sélectionnés sur cette période.';
  const parts=[`${{modifications.length}} modification(s) et ${{es.length-modifications.length}} action(s) sur les commentaires ont été relevées sur ${{is.length}} ticket(s).`];
  if(fields.length) parts.push(`Le champ le plus travaillé est « ${{fields[0][0]}} » avec ${{fields[0][1]}} modification(s).`);
  if(authors.length) parts.push(`${{authors[0][0]}} est le contributeur le plus actif (${{authors[0][1]}} action(s)).`);
  if(ticketCounts.length) parts.push(`Le ticket ${{ticketCounts[0][0]}} concentre le plus d’activité (${{ticketCounts[0][1]}} action(s)).`);
  return parts.join(' ');
}}
function activityChartHtml(es) {{
  let entries=sortedCounts(countBy(es,e=>(e.event_date||'').slice(0,10))).sort((a,b)=>a[0].localeCompare(b[0]));
  if(!entries.length) return '<div class="report-empty">Aucune activité sur la période sélectionnée.</div>';
  if(entries.length>18) {{
    const size=Math.ceil(entries.length/18), grouped=[];
    for(let i=0;i<entries.length;i+=size) {{
      const part=entries.slice(i,i+size), total=part.reduce((sum,row)=>sum+row[1],0);
      grouped.push([part.length===1?part[0][0]:`${{part[0][0]}} → ${{part[part.length-1][0]}}`,total]);
    }}
    entries=grouped;
  }}
  const width=860,height=245,left=44,right=15,top=18,bottom=54,plotW=width-left-right,plotH=height-top-bottom;
  const max=Math.max(...entries.map(row=>row[1]),1), step=plotW/entries.length, barW=Math.max(8,Math.min(34,step*.64));
  let grid='',bars='';
  for(let i=0;i<=4;i++) {{
    const y=top+plotH-(plotH*i/4), value=Math.round(max*i/4);
    grid+=`<line x1="${{left}}" x2="${{width-right}}" y1="${{y}}" y2="${{y}}" stroke="#dbe4ef" stroke-width="1"/><text x="${{left-8}}" y="${{y+4}}" text-anchor="end" font-size="10" fill="#64748b">${{value}}</text>`;
  }}
  entries.forEach(([label,value],idx)=>{{
    const x=left+step*idx+(step-barW)/2,h=Math.max(2,value/max*plotH),y=top+plotH-h;
    const shortLabel=label.length>10 ? label.slice(5,10)+'…' : label.slice(5).replace('-','/');
    bars+=`<rect x="${{x}}" y="${{y}}" width="${{barW}}" height="${{h}}" rx="4" fill="url(#activityGradient)"><title>${{escapeHtml(label)}} : ${{value}} action(s)</title></rect><text x="${{x+barW/2}}" y="${{Math.max(top+10,y-5)}}" text-anchor="middle" font-size="10" font-weight="700" fill="#334155">${{value}}</text><text x="${{x+barW/2}}" y="${{height-29}}" text-anchor="end" transform="rotate(-38 ${{x+barW/2}} ${{height-29}})" font-size="9" fill="#64748b">${{escapeHtml(shortLabel)}}</text>`;
  }});
  return `<svg viewBox="0 0 ${{width}} ${{height}}" role="img" aria-label="Évolution des actions"><defs><linearGradient id="activityGradient" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1F4E78"/><stop offset="1" stop-color="#0F766E"/></linearGradient></defs>${{grid}}${{bars}}</svg>`;
}}
function fieldDistributionHtml(es) {{
  let rows=sortedCounts(countBy(es,e=>e.field));
  if(!rows.length) return '<div class="report-empty">Aucune action à répartir.</div>';
  if(rows.length>6) rows=[...rows.slice(0,6),['Autres',rows.slice(6).reduce((sum,row)=>sum+row[1],0)]];
  const colors=['#1F4E78','#0F766E','#F59E0B','#7C3AED','#EF4444','#06B6D4','#94A3B8'];
  const total=rows.reduce((sum,row)=>sum+row[1],0), circumference=2*Math.PI*58;
  let offset=0,circles='',legend='';
  rows.forEach(([label,value],idx)=>{{
    const length=value/total*circumference,color=colors[idx%colors.length];
    circles+=`<circle cx="85" cy="85" r="58" fill="none" stroke="${{color}}" stroke-width="27" stroke-dasharray="${{length}} ${{circumference-length}}" stroke-dashoffset="${{-offset}}" transform="rotate(-90 85 85)"/>`;
    offset+=length;
    legend+=`<div class="report-legend-row"><span class="report-legend-dot" style="background:${{color}}"></span><span>${{escapeHtml(label)}}</span><b>${{value}} · ${{Math.round(value/total*100)}} %</b></div>`;
  }});
  return `<svg viewBox="0 0 170 170" style="max-height:48mm"><circle cx="85" cy="85" r="58" fill="none" stroke="#e2e8f0" stroke-width="27"/>${{circles}}<text x="85" y="81" text-anchor="middle" font-size="25" font-weight="800" fill="#17365D">${{total}}</text><text x="85" y="100" text-anchor="middle" font-size="10" fill="#64748b">actions</text></svg><div class="report-legend">${{legend}}</div>`;
}}
function authorChartHtml(es) {{
  const rows=sortedCounts(countBy(es,e=>e.author)).slice(0,10);
  if(!rows.length) return '<div class="report-empty">Aucun contributeur identifié.</div>';
  const max=rows[0][1];
  return rows.map(([author,value])=>`<div class="report-author"><div class="report-author-name">${{escapeHtml(author)}}</div><div class="report-author-track"><div class="report-author-bar" style="width:${{Math.max(4,value/max*100)}}%"></div></div><b>${{value}}</b></div>`).join('');
}}
function ticketTableHtml(is,es) {{
  const eventByIssue={{}};
  es.forEach(e=>{{ if(!eventByIssue[e.issue_id]) eventByIssue[e.issue_id]=[]; eventByIssue[e.issue_id].push(e); }});
  const sorted=[...is].sort((a,b)=>(eventByIssue[b.issue_id]||[]).length-(eventByIssue[a.issue_id]||[]).length || a.key.localeCompare(b.key));
  if(!sorted.length) return '<div class="report-empty">Aucun ticket dans la sélection.</div>';
  const rows=sorted.map(i=>{{
    const ticketEvents=eventByIssue[i.issue_id]||[], fieldCount=field=>ticketEvents.filter(e=>e.field===field).length;
    const mods=ticketEvents.filter(e=>e.event_source==='changelog').length, comments=ticketEvents.length-mods;
    return `<tr><td><a class="report-ticket-key" href="${{escapeHtml(i.url)}}">${{escapeHtml(i.key)}}</a><br><span style="color:#64748b">${{escapeHtml(i.issue_type||'')}}</span></td><td>${{escapeHtml(truncateReportText(i.summary,105))}}</td><td>${{escapeHtml(i.status||'—')}}</td><td><b>${{mods}}</b></td><td>${{fieldCount('Description')}}</td><td>${{fieldCount('Acceptance Criteria')}}</td><td>${{fieldCount('Statut')}}</td><td>${{fieldCount('Assignation')}}</td><td>${{comments}}</td></tr>`;
  }}).join('');
  return `<table class="report-table"><thead><tr><th>Ticket</th><th>Résumé</th><th>Statut actuel</th><th>Modif.</th><th>Descr.</th><th>AC</th><th>Statut</th><th>Réassign.</th><th>Comm.</th></tr></thead><tbody>${{rows}}</tbody></table>`;
}}
function reportChangeColor(e) {{
  return {{desc:'#10B981',ac:'#F59E0B',status:'#0F766E',assign:'#EF4444',comment:'#64748B'}}[eventClass(e)]||'#1F4E78';
}}
function changeDetailHtml(e) {{
  const rawBefore=e.from_value||e.from_excerpt||'',rawAfter=e.to_value||e.to_excerpt||'';
  const before=truncateReportText(rawBefore,520),after=truncateReportText(rawAfter,520);
  const exactDiff=`<div class="report-precise-head"><div class="report-precise-title">Diff précis</div><div class="report-diff-legend"><span class="report-diff-key deleted">Supprimé</span><span class="report-diff-key added">Ajouté</span></div></div><div class="report-precise-diff">${{preciseDiff(rawBefore,rawAfter)}}</div>`;
  return `<article class="report-change" style="--change-color:${{reportChangeColor(e)}}"><div class="report-change-head"><div class="report-change-title">${{escapeHtml(e.key)}} · ${{escapeHtml(e.field||'Action')}}</div><div class="report-change-meta">${{escapeHtml(formatDate(e.event_date))}}</div></div><div class="report-change-type">${{escapeHtml(e.change_type||'Modification')}} · par <b>${{escapeHtml(e.author||'Non renseigné')}}</b></div>${{exactDiff}}<div class="report-before-after"><div class="report-value before"><b>Avant</b>${{escapeHtml(before||'—')}}</div><div class="report-value after"><b>Après</b>${{escapeHtml(after||'—')}}</div></div></article>`;
}}
function detailsHtml(es) {{
  const order=document.getElementById('reportDetailOrder').value;
  let sorted=[...es];
  if(order==='oldest') sorted.sort((a,b)=>String(a.event_date).localeCompare(String(b.event_date)));
  else if(order==='ticket') sorted.sort((a,b)=>String(a.key).localeCompare(String(b.key)) || String(b.event_date).localeCompare(String(a.event_date)));
  else sorted.sort((a,b)=>String(b.event_date).localeCompare(String(a.event_date)));
  const rawLimit=document.getElementById('reportDetailLimit').value,limit=rawLimit==='all'?sorted.length:Number(rawLimit);
  const visible=sorted.slice(0,limit),note=visible.length<sorted.length?`<p class="report-section-intro">${{visible.length}} actions affichées sur ${{sorted.length}}. Modifiez la limite dans les options d’export pour en inclure davantage.</p>`:'';
  return visible.length ? note+visible.map(changeDetailHtml).join('') : '<div class="report-empty">Aucune modification à détailler.</div>';
}}
function reportPage(title,intro,content) {{
  return `<div class="report-page"><section class="report-section"><h2>${{escapeHtml(title)}}</h2>${{intro?`<p class="report-section-intro">${{escapeHtml(intro)}}</p>`:''}}${{content}}</section></div>`;
}}
function buildPdfReport(is,es,sections) {{
  const title=document.getElementById('reportTitle').value.trim()||'Rapport d’activité Jira';
  const modifications=es.filter(e=>e.event_source==='changelog'),comments=es.length-modifications.length;
  const authors=new Set(es.map(e=>e.author).filter(Boolean)),fields=new Set(modifications.map(e=>e.field).filter(Boolean));
  const generatedAt=new Date(),generatedLabel=generatedAt.toLocaleString('fr-FR');
  let html=`<section class="report-cover"><div class="report-brand"><span>Jira · Rapport d’activité</span><span>Document de synthèse</span></div><div class="report-cover-main"><div class="report-kicker">Sélection du dashboard</div><h1>${{escapeHtml(title)}}</h1><p class="report-cover-subtitle">Une lecture claire des actions réalisées, de leur volume et des tickets concernés.</p><div class="report-cover-meta"><div class="report-cover-stat"><b>${{is.length}}</b><span>tickets analysés</span></div><div class="report-cover-stat"><b>${{modifications.length}}</b><span>modifications</span></div><div class="report-cover-stat"><b>${{authors.size}}</b><span>contributeurs</span></div></div></div><div class="report-cover-foot"><span>Période : ${{escapeHtml(state.start||'début')}} → ${{escapeHtml(state.end||'fin')}}</span><span>Généré le ${{escapeHtml(generatedLabel)}}</span></div></section>`;
  if(sections.has('summary')) {{
    const kpis=reportKpi('Tickets',is.length,'dans la sélection','#93c5fd')+reportKpi('Modifications',modifications.length,'hors commentaires','#86efac')+reportKpi('Actions totales',es.length,'modifications + commentaires','#67e8f9')+reportKpi('Champs modifiés',fields.size,'types de champs distincts','#fcd34d')+reportKpi('Contributeurs',authors.size,'auteurs identifiés','#c4b5fd')+reportKpi('Commentaires',comments,'créations ou éditions','#fda4af');
    html+=reportPage('Synthèse exécutive','Périmètre exact repris depuis les filtres du dashboard.',`<div class="report-filters">${{reportFilterChips()}}</div><div class="report-kpis">${{kpis}}</div><div class="report-highlight"><b>À retenir.</b> ${{escapeHtml(executiveText(is,es))}}</div>`);
  }}
  if(sections.has('activity')) html+=reportPage('Activité dans le temps','Nombre d’actions Jira relevées pour chaque date de la période sélectionnée.',`<div class="report-chart-card"><h3>Évolution des actions</h3>${{activityChartHtml(es)}}</div>`);
  if(sections.has('fields')) html+=reportPage('Nature des actions et contributeurs','Répartition des actions retenues par champ Jira et par auteur.',`<div class="report-two-cols"><div class="report-chart-card"><h3>Champs concernés</h3>${{fieldDistributionHtml(es)}}</div><div class="report-chart-card"><h3>Contributeurs les plus actifs</h3>${{authorChartHtml(es)}}</div></div>`);
  if(sections.has('tickets')) html+=reportPage('Tickets concernés','Le volume détaille les modifications, descriptions, critères d’acceptation, transitions, réassignations et commentaires.',ticketTableHtml(is,es));
  if(sections.has('details')) html+=reportPage('Ce qui a été fait','Détail chronologique des actions avec auteur et comparaison avant / après.',detailsHtml(es));
  html+=`<div class="report-footer"><span>${{escapeHtml(title)}} · Sélection du dashboard Jira</span><span>Généré le ${{escapeHtml(generatedLabel)}}</span></div>`;
  return {{html,title}};
}}
function generatePdf() {{
  const sections=selectedReportSections();
  if(!sections.size) {{ window.alert('Sélectionnez au moins une section à inclure dans le rapport.'); return; }}
  const {{filteredIssues,filteredEvents}}=getFilteredData(),report=buildPdfReport(filteredIssues,filteredEvents,sections);
  const target=document.getElementById('pdfReport'),previousTitle=document.title;
  target.innerHTML=report.html; target.setAttribute('aria-hidden','false'); closeExportDialog();
  document.title=(report.title+'_'+(state.start||'debut')+'_'+(state.end||'fin')).replace(/[^a-zA-Z0-9À-ÿ_-]+/g,'-');
  const cleanup=()=>{{ target.innerHTML=''; target.setAttribute('aria-hidden','true'); document.title=previousTitle; window.removeEventListener('afterprint',cleanup); }};
  window.addEventListener('afterprint',cleanup);
  requestAnimationFrame(()=>requestAnimationFrame(()=>window.print()));
}}
function preciseDiff(a,b) {{
  a=String(a||''); b=String(b||'');
  if(a===b) return '<span class="context">Aucun changement textuel détecté.</span>';
  const ta = tokenize(a), tb = tokenize(b);
  if(ta.length>850 || tb.length>850) return diffWindow(a,b);
  const n=ta.length,m=tb.length;
  const dp=Array.from({{length:n+1}},()=>new Uint16Array(m+1));
  for(let i=n-1;i>=0;i--) for(let j=m-1;j>=0;j--) dp[i][j]=ta[i]===tb[j] ? dp[i+1][j+1]+1 : Math.max(dp[i+1][j],dp[i][j+1]);
  let i=0,j=0,out='';
  while(i<n && j<m) {{
    if(ta[i]===tb[j]) {{ out += '<span class="context">'+escapeHtml(ta[i])+'</span>'; i++; j++; }}
    else if(dp[i+1][j] >= dp[i][j+1]) {{ out += '<del>'+escapeHtml(ta[i])+'</del>'; i++; }}
    else {{ out += '<ins>'+escapeHtml(tb[j])+'</ins>'; j++; }}
  }}
  while(i<n) out += '<del>'+escapeHtml(ta[i++])+'</del>';
  while(j<m) out += '<ins>'+escapeHtml(tb[j++])+'</ins>';
  return out;
}}
function tokenize(s) {{ return String(s||'').match(/\s+|[A-Za-zÀ-ÿ0-9_]+|[^A-Za-zÀ-ÿ0-9_\s]/g) || []; }}
function diffWindow(a,b) {{
  let start=0; const min=Math.min(a.length,b.length);
  while(start<min && a[start]===b[start]) start++;
  let endA=a.length-1,endB=b.length-1;
  while(endA>=start && endB>=start && a[endA]===b[endB]) {{ endA--; endB--; }}
  const ctx=180;
  const prefix=a.slice(Math.max(0,start-ctx),start);
  const oldFrag=a.slice(start,endA+1);
  const newFrag=b.slice(start,endB+1);
  const suffix=a.slice(endA+1,Math.min(a.length,endA+1+ctx));
  return '<span class="context">'+escapeHtml((start>ctx?'…':'')+prefix)+'</span>' + (oldFrag?'<del>'+escapeHtml(oldFrag)+'</del>':'') + (newFrag?'<ins>'+escapeHtml(newFrag)+'</ins>':'') + '<span class="context">'+escapeHtml(suffix+(endA+1+ctx<a.length?'…':''))+'</span>';
}}
function escapeHtml(s) {{ return String(s||'').replace(/[&<>"']/g, m=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[m])); }}
function formatDate(s) {{ if(!s) return ''; const d=new Date(s); return isNaN(d) ? s : d.toLocaleString('fr-FR'); }}
init();
</script>
</body>
</html>"""
    path.write_text(html_text, encoding="utf-8")

# =============================================================================
# Orchestration
# =============================================================================


def fetch_real_data(args: argparse.Namespace, log: Callable[[str], None] = print) -> Tuple[List[IssueRow], List[TimelineEvent]]:
    client = JiraCookieClient(args.base_url, browser=args.browser, log=log)
    client.ensure_logged_in()

    raw_issues: List[Dict[str, Any]] = []
    if args.issue_keys:
        keys = split_csv(args.issue_keys)
        log(f"Récupération de {len(keys)} ticket(s) explicitement demandés...")
        for key in keys:
            raw_issues.append(fetch_issue(client, key, args.sprint_field))
    else:
        jql = build_jql(args)
        log("JQL: " + jql)
        raw_issues = search_issues(client, jql, args.sprint_field, args.max_results, args.limit, log=log)

    issues = [parse_issue_row(i, args.base_url, args.sprint_field) for i in raw_issues]
    sprint_contains = normalize(args.sprint_contains or "")
    if sprint_contains:
        keep = [i for i in issues if sprint_contains in normalize(i.sprint_names)]
        keep_ids = {i.issue_id for i in keep}
        issues = keep
        raw_issues = [i for i in raw_issues if str(i.get("id")) in keep_ids]
        log(f"Filtre local sprint_contains={args.sprint_contains}: {len(issues)} ticket(s)")

    all_events: List[TimelineEvent] = []
    for idx, issue in enumerate(issues, start=1):
        log(f"[{idx}/{len(issues)}] Historique {issue.key}")
        histories = fetch_all_changelog(client, issue.key)
        all_events.extend(changelog_to_events(issue, histories))
        if not args.no_comments:
            try:
                comments = fetch_all_comments(client, issue.key)
                all_events.extend(comments_to_events(issue, comments))
            except Exception as exc:
                log(f"  Warning commentaires {issue.key}: {exc}")
        if args.pause:
            time.sleep(args.pause)
    return issues, all_events


def demo_data(base_url: str) -> Tuple[List[IssueRow], List[TimelineEvent]]:
    issue = IssueRow(
        issue_id="457866", key="PPMG-2185", url=f"{base_url.rstrip('/')}/browse/PPMG-2185",
        summary="[INC0911736] - Le Business Project Status reste à Feasibility Study après passage en phase Réalisation",
        project_key="PPMG", project_name="[DDSI - PRJ] PPM Group (VYP)", issue_type="Bug", status="Ready for Testing", status_category="A faire",
        priority="Medium", assignee="Yosra MANAI", reporter="Romain CIESLIK", sprint_names="Sprint 16 2026", sprint_ids="16029",
        created=parse_jira_datetime("2026-08-06T15:32:18.633+0200"), updated=parse_jira_datetime("2026-08-10T10:08:46.228+0200"), parent="",
    )
    events = [
        TimelineEvent(issue.issue_id, issue.key, issue.url, issue.summary, issue.project_key, issue.issue_type, issue.status, issue.assignee, issue.sprint_names, parse_jira_datetime("2026-08-06T15:32:18.633+0200"), "Romain CIESLIK", "changelog", "Description", "description", "Description modifiée", "Ancienne description", "Nouvelle description avec reprise de données", "Ancienne description", "Nouvelle description avec reprise de données", "1", ""),
        TimelineEvent(issue.issue_id, issue.key, issue.url, issue.summary, issue.project_key, issue.issue_type, issue.status, issue.assignee, issue.sprint_names, parse_jira_datetime("2026-08-06T15:48:00.000+0200"), "Romain CIESLIK", "changelog", "Acceptance Criteria", "customfield_10056", "Acceptance Criteria modifiés", "AC 1 - fréquence quotidienne AC 2 - validation métier", "AC 1, fréquence quotidienne AC 2 - validation métier", "AC 1 - fréquence quotidienne AC 2 - validation métier", "AC 1, fréquence quotidienne AC 2 - validation métier", "1b", ""),
        TimelineEvent(issue.issue_id, issue.key, issue.url, issue.summary, issue.project_key, issue.issue_type, issue.status, issue.assignee, issue.sprint_names, parse_jira_datetime("2026-08-06T16:03:00.000+0200"), "Romain CIESLIK", "changelog", "Acceptance Criteria", "customfield_10056", "Acceptance Criteria modifiés", "AC1 verifier la frequence quotidienne", "AC1 verifier la frequence, quotidienne", "AC1 verifier la frequence quotidienne", "AC1 verifier la frequence, quotidienne", "2a", ""),
        TimelineEvent(issue.issue_id, issue.key, issue.url, issue.summary, issue.project_key, issue.issue_type, issue.status, issue.assignee, issue.sprint_names, parse_jira_datetime("2026-08-06T16:07:53.565+0200"), "Romain CIESLIK", "changelog", "Statut", "status", "Statut modifié", "Open", "Validated", "Open", "Validated", "2", ""),
        TimelineEvent(issue.issue_id, issue.key, issue.url, issue.summary, issue.project_key, issue.issue_type, issue.status, issue.assignee, issue.sprint_names, parse_jira_datetime("2026-08-10T09:45:20.886+0200"), "Romain CIESLIK", "changelog", "Assignation", "assignee", "Assignation modifiée", "Romain CIESLIK", "RACHEL CLOIX", "Romain CIESLIK", "RACHEL CLOIX", "3", ""),
        TimelineEvent(issue.issue_id, issue.key, issue.url, issue.summary, issue.project_key, issue.issue_type, issue.status, issue.assignee, issue.sprint_names, parse_jira_datetime("2026-08-10T10:08:46.228+0200"), "RACHEL CLOIX", "changelog", "Assignation", "assignee", "Assignation modifiée", "RACHEL CLOIX", "Yosra MANAI", "RACHEL CLOIX", "Yosra MANAI", "4", ""),
        TimelineEvent(issue.issue_id, issue.key, issue.url, issue.summary, issue.project_key, issue.issue_type, issue.status, issue.assignee, issue.sprint_names, parse_jira_datetime("2026-08-06T15:39:24.796+0200"), "Romain CIESLIK", "comment", "Commentaire", "comment", "Commentaire créé", "", "Commentaire de test sur la reproduction du bug", "", "Commentaire de test sur la reproduction du bug", "", "356323"),
    ]
    return [issue], events


def run_analysis(args: argparse.Namespace, log: Callable[[str], None] = print) -> Tuple[Path, Path, int, int]:
    if args.demo:
        issues, events = demo_data(args.base_url)
    else:
        issues, events = fetch_real_data(args, log=log)
    out_xlsx = Path(args.output_xlsx)
    out_html = Path(args.output_html)
    write_workbook(out_xlsx, issues, events, style=args.style)
    write_html(out_html, issues, events, style=args.style)
    log("Terminé.")
    log(f"Excel : {out_xlsx.resolve()}")
    log(f"HTML  : {out_html.resolve()}")
    log(f"Tickets: {len(issues)} | Événements: {len(events)}")
    if args.open_html:
        try:
            webbrowser.open(out_html.resolve().as_uri())
        except Exception:
            pass
    return out_xlsx, out_html, len(issues), len(events)


# =============================================================================
# Simple Tkinter GUI
# =============================================================================


def launch_gui() -> None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except Exception as exc:
        raise SystemExit("Tkinter n'est pas disponible dans cette installation Python.") from exc

    root = tk.Tk()
    root.title("Jira Changelog Analyzer - session navigateur")
    root.geometry("980x720")

    values: Dict[str, tk.StringVar] = {
        "base_url": tk.StringVar(value=DEFAULT_BASE_URL),
        "browser": tk.StringVar(value="edge"),
        "project": tk.StringVar(value=DEFAULT_PROJECT),
        "types": tk.StringVar(value="Story,Bug"),
        "api_sprint_id": tk.StringVar(value=""),
        "api_sprint_name": tk.StringVar(value=""),
        "sprint_contains": tk.StringVar(value=""),
        "issue_keys": tk.StringVar(value=""),
        "limit": tk.StringVar(value="20"),
        "output_xlsx": tk.StringVar(value="jira_changelog_analysis.xlsx"),
        "output_html": tk.StringVar(value="jira_changelog_dashboard.html"),
        "style": tk.StringVar(value="executive"),
    }
    no_comments = tk.BooleanVar(value=False)
    open_html = tk.BooleanVar(value=True)
    demo = tk.BooleanVar(value=False)

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)

    def add_row(row: int, label: str, key: str, width: int = 42, widget: str = "entry", choices: Optional[List[str]] = None) -> None:
        ttk.Label(frm, text=label).grid(row=row, column=0, sticky="w", padx=4, pady=4)
        if widget == "combo":
            cb = ttk.Combobox(frm, textvariable=values[key], values=choices or [], width=width, state="readonly")
            cb.grid(row=row, column=1, sticky="we", padx=4, pady=4)
        else:
            ttk.Entry(frm, textvariable=values[key], width=width).grid(row=row, column=1, sticky="we", padx=4, pady=4)

    add_row(0, "URL Jira", "base_url")
    add_row(1, "Navigateur cookies", "browser", widget="combo", choices=["auto", "edge", "chrome", "firefox"])
    add_row(2, "Projet", "project")
    add_row(3, "Types", "types")
    add_row(4, "Sprint ID API optionnel", "api_sprint_id")
    add_row(5, "Nom sprint exact optionnel", "api_sprint_name")
    add_row(6, "Sprint contient optionnel", "sprint_contains")
    add_row(7, "Tickets explicites optionnel", "issue_keys")
    add_row(8, "Limite test", "limit")
    add_row(9, "Excel sortie", "output_xlsx")
    add_row(10, "HTML sortie", "output_html")
    add_row(11, "Style", "style", widget="combo", choices=["executive", "audit", "kanban"])

    ttk.Checkbutton(frm, text="Ne pas récupérer les commentaires", variable=no_comments).grid(row=12, column=0, sticky="w", padx=4, pady=4)
    ttk.Checkbutton(frm, text="Ouvrir le HTML à la fin", variable=open_html).grid(row=12, column=1, sticky="w", padx=4, pady=4)
    ttk.Checkbutton(frm, text="Mode démo sans Jira", variable=demo).grid(row=13, column=0, sticky="w", padx=4, pady=4)

    def choose_xlsx() -> None:
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            values["output_xlsx"].set(p)

    def choose_html() -> None:
        p = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML", "*.html")])
        if p:
            values["output_html"].set(p)

    ttk.Button(frm, text="Choisir Excel", command=choose_xlsx).grid(row=9, column=2, padx=4, pady=4)
    ttk.Button(frm, text="Choisir HTML", command=choose_html).grid(row=10, column=2, padx=4, pady=4)

    log_box = tk.Text(frm, height=18, wrap="word")
    log_box.grid(row=15, column=0, columnspan=3, sticky="nsew", padx=4, pady=10)
    frm.columnconfigure(1, weight=1)
    frm.rowconfigure(15, weight=1)

    def log(msg: str) -> None:
        log_box.after(0, lambda: (log_box.insert("end", msg + "\n"), log_box.see("end")))

    def build_args() -> argparse.Namespace:
        limit_txt = values["limit"].get().strip()
        limit_val = int(limit_txt) if limit_txt else None
        return argparse.Namespace(
            base_url=values["base_url"].get().strip(),
            browser=values["browser"].get().strip(),
            project=values["project"].get().strip(),
            types=values["types"].get().strip(),
            api_sprint_id=values["api_sprint_id"].get().strip() or None,
            api_sprint_name=values["api_sprint_name"].get().strip() or None,
            sprint_contains=values["sprint_contains"].get().strip() or None,
            issue_keys=values["issue_keys"].get().strip() or None,
            jql=None,
            jql_extra=None,
            sprint_field=DEFAULT_SPRINT_FIELD,
            max_results=100,
            limit=limit_val,
            pause=0.05,
            no_comments=no_comments.get(),
            output_xlsx=values["output_xlsx"].get().strip(),
            output_html=values["output_html"].get().strip(),
            style=values["style"].get().strip(),
            demo=demo.get(),
            open_html=open_html.get(),
        )

    def run_thread() -> None:
        try:
            args = build_args()
            run_analysis(args, log=log)
            log("OK.")
        except Exception as exc:
            log("ERREUR: " + str(exc))
            messagebox.showerror("Erreur", str(exc))

    def run_clicked() -> None:
        log_box.delete("1.0", "end")
        threading.Thread(target=run_thread, daemon=True).start()

    def open_jira_login() -> None:
        webbrowser.open(values["base_url"].get().rstrip("/") + "/rest/api/3/myself")

    btn_frame = ttk.Frame(frm)
    btn_frame.grid(row=14, column=0, columnspan=3, sticky="we", pady=8)
    ttk.Button(btn_frame, text="1. Ouvrir Jira dans mon navigateur", command=open_jira_login).pack(side="left", padx=4)
    ttk.Button(btn_frame, text="2. Lancer extraction", command=run_clicked).pack(side="left", padx=4)

    info = (
        "Mode d'emploi: ouvrez Jira dans votre navigateur avec le bouton 1. "
        "Si /myself affiche du JSON, revenez ici et cliquez sur Lancer extraction. "
        "Le script utilise les cookies locaux de votre navigateur."
    )
    ttk.Label(frm, text=info, wraplength=880).grid(row=16, column=0, columnspan=3, sticky="w", padx=4, pady=4)

    root.mainloop()


# =============================================================================
# CLI
# =============================================================================


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Analyse Jira Cloud changelog avec cookies du navigateur deja connecte.")
    ap.add_argument("--gui", action="store_true", help="Ouvre l'interface graphique")
    ap.add_argument("--base-url", default=DEFAULT_BASE_URL)
    ap.add_argument("--browser", choices=["auto", "edge", "chrome", "firefox"], default="edge")
    ap.add_argument("--project", default=DEFAULT_PROJECT)
    ap.add_argument("--types", default="Story,Bug", help="Types a recuperer, ex: Story,Bug")
    ap.add_argument("--issue-keys", default=None, help="Liste de tickets explicites, ex: PPMG-2185,PPMG-2177. Remplace la recherche JQL.")
    ap.add_argument("--api-sprint-id", default=None, help="Filtre JQL par ID de sprint")
    ap.add_argument("--api-sprint-name", default=None, help="Filtre JQL par nom exact de sprint")
    ap.add_argument("--sprint-contains", default=None, help="Filtre local par texte de sprint, ex: 17")
    ap.add_argument("--jql", default=None, help="JQL complet. Remplace project/types.")
    ap.add_argument("--jql-extra", default=None, help="Clause JQL additionnelle")
    ap.add_argument("--sprint-field", default=DEFAULT_SPRINT_FIELD)
    ap.add_argument("--max-results", type=int, default=100)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--pause", type=float, default=0.05)
    ap.add_argument("--no-comments", action="store_true")
    ap.add_argument("--output-xlsx", default="jira_changelog_analysis.xlsx")
    ap.add_argument("--output-html", default="jira_changelog_dashboard.html")
    ap.add_argument("--style", choices=["executive", "audit", "kanban"], default="executive")
    ap.add_argument("--demo", action="store_true")
    ap.add_argument("--open-html", action="store_true", default=True)
    return ap.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> None:
    if argv is None:
        argv = sys.argv[1:]
    if not argv or "--gui" in argv:
        launch_gui()
        return
    args = parse_args(argv)
    run_analysis(args, log=print)


if __name__ == "__main__":
    main()
