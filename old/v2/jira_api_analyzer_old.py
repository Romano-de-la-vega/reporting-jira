#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Jira Cloud API changelog analyzer.

But:
    Recuperer les tickets Jira via l'API REST, extraire leur changelog complet,
    leurs commentaires, puis produire :
      - un Excel d'analyse avec timelines filtrables ;
      - une interface HTML locale avec filtres et timeline par ticket.

Cas Servier / Jira Cloud :
    Les API tokens personnels peuvent etre bloques. Ce script supporte donc
    une authentification via navigateur Playwright (--auth browser), qui ouvre
    une session Jira/SSO puis utilise les cookies du navigateur pour appeler l'API.

Installation :
    pip install -r requirements_jira_api.txt
    python -m playwright install chromium

Exemple :
    python jira_api_changelog_analyzer.py \
      --base-url https://data-servier.atlassian.net \
      --project PPMG \
      --types Story,Bug \
      --output-xlsx jira_changelog.xlsx \
      --output-html jira_changelog.html

Mode demo sans API :
    python jira_api_changelog_analyzer.py --demo

Notes :
    - Le endpoint de recherche utilise /rest/api/3/search/jql avec nextPageToken.
    - Les details de l'historique sont recuperes via /rest/api/3/issue/{key}/changelog.
    - Les commentaires sont recuperes via /rest/api/3/issue/{key}/comment.
"""

from __future__ import annotations

import argparse
import base64
import getpass
import html
import json
import os
import re
import sys
import time
import webbrowser
from collections import Counter, defaultdict
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import quote

try:
    import xlsxwriter
except Exception:
    xlsxwriter = None


# =============================================================================
# Helpers texte/date
# =============================================================================

CONTROL_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
MAX_EXCEL = 32000


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


def one_line(value: Any, limit: int = 280) -> str:
    text = clean_text(value).replace("\n", " | ")
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 3)].rstrip() + "..."


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
    # Jira: 2026-08-10T10:08:46.228+0200
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
        # Convert +0200 -> +02:00 if needed.
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
    if len(text) > MAX_EXCEL:
        text = text[: MAX_EXCEL - 15] + "... [TRUNCATED]"
    if text.startswith("="):
        text = "'" + text
    return text


def split_csv(value: Optional[str]) -> List[str]:
    if not value:
        return []
    return [v.strip() for v in value.split(",") if v.strip()]


# =============================================================================
# Atlassian Document Format -> texte
# =============================================================================


def adf_to_text(node: Any) -> str:
    """Convertit rapidement l'Atlassian Document Format en texte lisible."""
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
    if ntype == "blockquote":
        return "\n".join(adf_to_text(x) for x in content)
    if ntype == "codeBlock":
        return "\n".join(adf_to_text(x) for x in content)
    if ntype == "mediaSingle":
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
    event_source: str       # changelog / comment
    field: str              # Description / Acceptance Criteria / Statut / Assignation / Commentaire...
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
# Jira clients
# =============================================================================


class JiraClientBase:
    def get_json(self, path: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        raise NotImplementedError

    def close(self) -> None:
        pass


class BrowserJiraClient(JiraClientBase):
    def __init__(self, base_url: str, user_data_dir: str, headless: bool = False, timeout_ms: int = 90000):
        self.base_url = base_url.rstrip("/")
        self.timeout_ms = timeout_ms
        try:
            from playwright.sync_api import sync_playwright
        except Exception as exc:
            raise SystemExit(
                "Playwright est requis pour --auth browser. Installez-le avec:\n"
                "  pip install playwright\n"
                "  python -m playwright install chromium"
            ) from exc
        self._pw = sync_playwright().start()
        self._context = self._pw.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=headless,
            viewport={"width": 1400, "height": 900},
        )
        self._page = self._context.pages[0] if self._context.pages else self._context.new_page()
        self.ensure_logged_in()

    def ensure_logged_in(self) -> None:
        url = f"{self.base_url}/rest/api/3/myself"
        self._page.goto(url, wait_until="domcontentloaded", timeout=self.timeout_ms)
        body = self._page.locator("body").inner_text(timeout=self.timeout_ms)
        if '"accountId"' in body or "accountId" in body:
            return
        print("\nConnexion Jira/SSO requise.")
        print("Une fenetre Chromium est ouverte. Connectez-vous a Jira, puis revenez ici.")
        input("Appuyez sur Entree quand la page /myself affiche du JSON... ")
        self._page.goto(url, wait_until="domcontentloaded", timeout=self.timeout_ms)
        body = self._page.locator("body").inner_text(timeout=self.timeout_ms)
        if '"accountId"' not in body and "accountId" not in body:
            raise SystemExit("Connexion Jira non detectee. Relancez le script apres connexion.")

    def get_json(self, path: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        url = self.base_url + path
        response = self._context.request.get(url, params=params or {}, timeout=self.timeout_ms)
        if response.status >= 400:
            text = response.text()
            raise RuntimeError(f"API Jira error {response.status} on {path}: {text[:1000]}")
        try:
            return response.json()
        except Exception as exc:
            raise RuntimeError(f"Reponse non JSON pour {path}: {response.text()[:500]}") from exc

    def close(self) -> None:
        try:
            self._context.close()
            self._pw.stop()
        except Exception:
            pass


class BasicJiraClient(JiraClientBase):
    def __init__(self, base_url: str, email: str, token: str, timeout: int = 90):
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        try:
            import requests
        except Exception as exc:
            raise SystemExit("requests est requis pour --auth basic. Lancez: pip install requests") from exc
        self.requests = requests
        self.session = requests.Session()
        self.session.auth = (email, token)
        self.session.headers.update({"Accept": "application/json"})

    def get_json(self, path: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        url = self.base_url + path
        response = self.session.get(url, params=params or {}, timeout=self.timeout)
        if response.status_code >= 400:
            raise RuntimeError(f"API Jira error {response.status_code} on {path}: {response.text[:1000]}")
        return response.json()


# =============================================================================
# API fetch
# =============================================================================


def jql_quote(value: str) -> str:
    escaped = value.replace('"', '\\"')
    return f'"{escaped}"'


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


def search_issues(client: JiraClientBase, jql: str, sprint_field: str, max_results: int, limit: Optional[int]) -> Tuple[List[Dict[str, Any]], Dict[str, str]]:
    fields = [
        "summary", "issuetype", "status", "assignee", "reporter", "created", "updated",
        "priority", sprint_field, "parent", "project",
    ]
    issues: List[Dict[str, Any]] = []
    names: Dict[str, str] = {}
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
        names.update(data.get("names") or {})
        batch = data.get("issues") or []
        issues.extend(batch)
        print(f"Issues recuperees: {len(issues)}")
        if limit and len(issues) >= limit:
            issues = issues[:limit]
            break
        if data.get("isLast") is True:
            break
        token = data.get("nextPageToken")
        if not token:
            break
    return issues, names


def fetch_all_changelog(client: JiraClientBase, issue_key_or_id: str) -> List[Dict[str, Any]]:
    histories: List[Dict[str, Any]] = []
    start_at = 0
    while True:
        data = client.get_json(f"/rest/api/3/issue/{quote(str(issue_key_or_id))}/changelog", params={"startAt": start_at, "maxResults": 100})
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


def fetch_all_comments(client: JiraClientBase, issue_key_or_id: str) -> List[Dict[str, Any]]:
    comments: List[Dict[str, Any]] = []
    start_at = 0
    while True:
        data = client.get_json(f"/rest/api/3/issue/{quote(str(issue_key_or_id))}/comment", params={"startAt": start_at, "maxResults": 100, "orderBy": "created"})
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
# Parsing issue/changelog
# =============================================================================


def user_display(user: Any) -> str:
    if not isinstance(user, dict):
        return ""
    return user.get("displayName") or user.get("emailAddress") or user.get("accountId") or ""


def parse_sprint_value(value: Any) -> Tuple[str, str]:
    if not value:
        return "", ""
    if isinstance(value, list):
        names = []
        ids = []
        for item in value:
            if isinstance(item, dict):
                if item.get("name"):
                    names.append(str(item.get("name")))
                if item.get("id") is not None:
                    ids.append(str(item.get("id")))
            else:
                txt = str(item)
                # Old GreenHopper sprint string sometimes contains name=...
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


def count_unique(events: Sequence[TimelineEvent]) -> int:
    return len({e.issue_id for e in events})


def sorted_events(events: Sequence[TimelineEvent]) -> List[TimelineEvent]:
    return sorted(events, key=lambda e: (e.key, e.event_date or datetime.min.replace(tzinfo=timezone.utc), e.field, e.history_id, e.comment_id))


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
        assignee_loop = count_return_loops(assign, value_attr="to_value")
        status_loop = count_return_loops(status, value_attr="to_value")
        first = min((e.event_date for e in evs if e.event_date), default=None)
        last = max((e.event_date for e in evs if e.event_date), default=None)
        churn_score = (
            len(desc) * 4 + len(ac) * 5 + len(status) * 3 +
            len(assign) * 2 + len(sprint) * 2 + len(comments)
        )
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
            "Assignment loops": assignee_loop,
            "Status loops": status_loop,
            "Churn score": churn_score,
        })
    return rows


def count_return_loops(events: Sequence[TimelineEvent], value_attr: str = "to_value") -> int:
    seq: List[str] = []
    for e in sorted(events, key=lambda x: x.event_date or datetime.min.replace(tzinfo=timezone.utc)):
        val = clean_text(getattr(e, value_attr))
        if val:
            seq.append(val)
    loops = 0
    for i in range(2, len(seq)):
        if seq[i] == seq[i - 2] and seq[i] != seq[i - 1]:
            loops += 1
    return loops


def events_by_day(events: Sequence[TimelineEvent]) -> List[Dict[str, Any]]:
    counter: Counter[str] = Counter()
    for e in events:
        if e.event_date:
            counter[e.event_date.astimezone(timezone.utc).date().isoformat()] += 1
    return [{"Date": k, "Événements": v} for k, v in sorted(counter.items())]


def events_by_field(events: Sequence[TimelineEvent]) -> List[Dict[str, Any]]:
    counter: Counter[str] = Counter(e.field for e in events)
    return [{"Champ": k, "Événements": v, "Tickets": len({e.issue_id for e in events if e.field == k})} for k, v in counter.most_common()]


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


def row_from_issue(i: IssueRow) -> Dict[str, Any]:
    d = asdict(i)
    d["created"] = dt_excel(i.created)
    d["updated"] = dt_excel(i.updated)
    return d


def write_dashboard(wb: Any, fmt: Dict[str, Any], p: Dict[str, str], issues: List[IssueRow], events: List[TimelineEvent], by_day: List[Dict[str, Any]], by_field: List[Dict[str, Any]], ticket_rows: List[Dict[str, Any]]) -> None:
    ws = wb.add_worksheet("00_Dashboard")
    ws.hide_gridlines(2)
    ws.merge_range("A1:N2", "Jira changelog analysis - API", fmt["title"])
    ws.set_row(0, 26)
    total_events = len(events)
    cards = [
        ("Tickets", len(issues)),
        ("Événements", total_events),
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

    # Chart source by field
    start_row = 18
    ws.write(start_row, 0, "Champ", fmt["header"])
    ws.write(start_row, 1, "Événements", fmt["header"])
    for idx, row in enumerate(by_field[:10], start=start_row + 1):
        ws.write(idx, 0, row["Champ"])
        ws.write_number(idx, 1, row["Événements"])
    if by_field:
        chart = wb.add_chart({"type": "column"})
        chart.add_series({
            "name": "Événements",
            "categories": ["00_Dashboard", start_row + 1, 0, start_row + min(10, len(by_field)), 0],
            "values": ["00_Dashboard", start_row + 1, 1, start_row + min(10, len(by_field)), 1],
        })
        chart.set_title({"name": "Événements par champ"})
        chart.set_legend({"none": True})
        chart.set_size({"width": 620, "height": 320})
        ws.insert_chart("D18", chart)

    # Chart by day
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


def cell_to_rowcol(cell: str) -> Tuple[int, int]:
    m = re.match(r"([A-Z]+)(\d+)", cell)
    if not m:
        return 0, 0
    letters, row_s = m.groups()
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return int(row_s) - 1, col - 1


def add_timeline_conditional_formats(ws: Any, nrows: int, fmt: Dict[str, Any]) -> None:
    # Column field index in TimelineEvent.as_dict order: field after event_source -> index 14? Safer by full row formula not needed.
    # We color the field column by text content.
    # headers: issue_id,key,url,summary,project_key,issue_type,current_status,current_assignee,current_sprints,event_date,author,event_source,field,...
    field_col = 12
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


# =============================================================================
# HTML dashboard
# =============================================================================


def write_html(path: Path, issues: List[IssueRow], events: List[TimelineEvent], style: str = "executive") -> None:
    palette = style_palette(style)
    issue_payload = [serialize_issue(i) for i in issues]
    event_payload = [serialize_event(e) for e in sorted_events(events)]
    html_text = build_html_template(issue_payload, event_payload, palette, style)
    path.write_text(html_text, encoding="utf-8")


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


def build_html_template(issues: List[Dict[str, Any]], events: List[Dict[str, Any]], palette: Dict[str, str], style: str) -> str:
    data_json = json.dumps({"issues": issues, "events": events}, ensure_ascii=False)
    css_vars = "\n".join([f"--{k}: #{v};" for k, v in palette.items()])
    return f"""<!doctype html>
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
.grid {{ display:grid; grid-template-columns: 380px 1fr; gap:18px; align-items:start; }}
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
<header><h1>Jira changelog timeline</h1><p>Style: {html.escape(style)} · filtres par type, sprint, ticket, champ et auteur</p></header>
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
  renderTimeline(filteredIssues, filteredEvents);
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
function renderTimeline(is, es) {{
  const sorted = [...es].sort((a,b)=>String(b.event_date).localeCompare(String(a.event_date))).slice(0,500);
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


# =============================================================================
# Demo data and orchestration
# =============================================================================


def demo_data(base_url: str) -> Tuple[List[IssueRow], List[TimelineEvent]]:
    issue = IssueRow(
        issue_id="457866", key="PPMG-2185", url=f"{base_url.rstrip('/')}/browse/PPMG-2185",
        summary="[INC0911736] - Le Business Project Status reste à Feasibility Study après passage en phase Réalisation",
        project_key="PPMG", issue_type="Bug", status="Ready for Testing", status_category="A faire",
        priority="Medium", assignee="Yosra MANAI", reporter="Romain CIESLIK", sprint_names="Sprint 16 2026", sprint_ids="16029",
        created=parse_jira_datetime("2026-08-06T15:32:18.633+0200"), updated=parse_jira_datetime("2026-08-10T10:08:46.228+0200"), parent="",
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
    if args.auth == "browser":
        client: JiraClientBase = BrowserJiraClient(
            args.base_url,
            user_data_dir=args.browser_profile,
            headless=args.headless,
        )
    elif args.auth == "basic":
        email = args.email or input("Email Atlassian: ").strip()
        token = args.token or getpass.getpass("API token: ")
        client = BasicJiraClient(args.base_url, email=email, token=token)
    else:
        raise SystemExit("--auth doit etre browser ou basic")

    try:
        jql = build_jql(args)
        print("JQL:", jql)
        issues_raw, names = search_issues(client, jql, args.sprint_field, args.max_results, args.limit)
        issues = [parse_issue_row(i, args.base_url, args.sprint_field) for i in issues_raw]
        sprint_contains = normalize(args.sprint_contains or "")
        if sprint_contains:
            keep_ids = {i.issue_id for i in issues if sprint_contains in normalize(i.sprint_names)}
            issues = [i for i in issues if i.issue_id in keep_ids]
            issues_raw = [i for i in issues_raw if str(i.get("id")) in keep_ids]
            print(f"Filtre local sprint_contains={args.sprint_contains}: {len(issues)} ticket(s)")

        issue_by_id = {i.issue_id: i for i in issues}
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
    finally:
        client.close()


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Analyse Jira Cloud changelog et genere Excel + HTML timeline.")
    ap.add_argument("--base-url", default="https://data-servier.atlassian.net", help="URL Jira Cloud")
    ap.add_argument("--project", default="PPMG", help="Cle projet Jira, defaut PPMG = [DDSI - PRJ] PPM Group (VYP)")
    ap.add_argument("--types", default="Story,Bug", help="Types a recuperer, ex: Story,Bug ou Story")
    ap.add_argument("--api-sprint-id", default=None, help="Filtre JQL par ID de sprint, ex: 17833")
    ap.add_argument("--api-sprint-name", default=None, help="Filtre JQL par nom exact de sprint, ex: Sprint 17 2026")
    ap.add_argument("--sprint-contains", default=None, help="Filtre local si vous donnez juste 15/16/17, ex: 17")
    ap.add_argument("--jql", default=None, help="JQL complet. Remplace --project/--types.")
    ap.add_argument("--jql-extra", default=None, help="Clause JQL additionnelle")
    ap.add_argument("--sprint-field", default="customfield_10020", help="Champ sprint Jira, defaut customfield_10020")
    ap.add_argument("--auth", choices=["browser", "basic"], default="browser", help="browser conseille si API token bloque")
    ap.add_argument("--browser-profile", default=".jira_browser_session", help="Dossier de session Playwright")
    ap.add_argument("--headless", action="store_true", help="Navigateur invisible apres login deja realise")
    ap.add_argument("--email", default=None, help="Email Atlassian pour --auth basic")
    ap.add_argument("--token", default=None, help="API token pour --auth basic")
    ap.add_argument("--max-results", type=int, default=100, help="Taille page search/jql")
    ap.add_argument("--limit", type=int, default=None, help="Limite de tickets pour test")
    ap.add_argument("--pause", type=float, default=0.05, help="Pause entre tickets")
    ap.add_argument("--no-comments", action="store_true", help="Ne pas recuperer les commentaires")
    ap.add_argument("--output-xlsx", default="jira_changelog_analysis.xlsx")
    ap.add_argument("--output-html", default="jira_changelog_dashboard.html")
    ap.add_argument("--style", choices=["executive", "audit", "kanban"], default="executive", help="Style HTML/Excel")
    ap.add_argument("--demo", action="store_true", help="Genere un exemple sans appeler Jira")
    return ap.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = parse_args(argv)
    if args.demo:
        issues, events = demo_data(args.base_url)
    else:
        issues, events = fetch_real_data(args)

    out_xlsx = Path(args.output_xlsx)
    out_html = Path(args.output_html)
    write_workbook(out_xlsx, issues, events, style=args.style)
    write_html(out_html, issues, events, style=args.style)
    print("\nTermine.")
    print(f"Excel : {out_xlsx.resolve()}")
    print(f"HTML  : {out_html.resolve()}")
    print(f"Tickets: {len(issues)} | Evenements: {len(events)}")
    try:
        webbrowser.open(out_html.resolve().as_uri())
    except Exception:
        pass


if __name__ == "__main__":
    main()
