#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Analyse KPI d'un export RSS/XML Jira avec matrice booleenne des donnees manquantes.

Usage:
    python analyse_jira_rss.py "ticket non clot (Jira)(1)(1).xml"
    python analyse_jira_rss.py "jira.xml" --output "jira_kpi_report.xlsx"
    python analyse_jira_rss.py "jira.xml" --output "jira_kpi_report.xlsx" --no-autofilter

Ce script cree un fichier Excel lisible avec:
    - Dashboard avec graphiques;
    - Boolean_Matrix: matrice booleenne triable/filtrable;
    - AC_Missing: tickets sans Acceptance Criteria;
    - regroupements par sprint, assignee, statut, type, projet;
    - Missing_Data_By_Field: champs les plus souvent manquants;
    - Raw_Tickets, CustomFields, Comments, Attachments.

Important:
    Le script ne cree JAMAIS de Table Excel structuree. Donc aucun /xl/tables/table*.xml.
    Les filtres sont des filtres simples de feuille. Si ton Excel/SharePoint pose encore probleme,
    lance avec --no-autofilter.
"""

from __future__ import annotations

import argparse
import csv
import html
import os
import re
import statistics
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None  # type: ignore

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

EMPTY_VALUES = {
    "", "-", "n/a", "na", "none", "null", "non renseigne", "non renseigné",
    "not specified", "ns (not specified)", "ns not specified", "not applicable",
    "n a", "n.a", "n.a.", "undefined",
}

NOT_REAL_VALUES = {
    "n/a", "na", "n a", "n.a", "n.a.", "ns not specified", "ns (not specified)",
    "not specified", "non renseigne", "non renseigné", "not applicable", "undefined",
}

TEMPLATE_WORDS = {
    "context", "as a", "i want", "so that", "acceptance criteria", "test 1",
    "given", "when", "then", "additional information", "cause", "cause description",
    "impact", "impact description", "if yes specify", "if yes, specify", "creation",
    "modification", "deletion", "other", "tad", "mocp", "role matrix", "ddic",
    "data model", "training documentation", "pivot", "talend flow", "business rules",
    "workflows", "regression", "existing functions", "audit trail", "roles permissions",
}

CONTROL_CHARS_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
MAX_EXCEL_TEXT = 32000


@dataclass
class Ticket:
    key: str
    id: str
    link: str
    project_key: str
    project_name: str
    title: str
    summary: str
    issue_type: str
    priority: str
    status: str
    status_category: str
    resolution: str
    assignee: str
    reporter: str
    created: Optional[datetime]
    updated: Optional[datetime]
    resolved: Optional[datetime]
    due: Optional[datetime]
    labels: str
    parent: str
    sprint: str
    theme: str
    target_type: str
    taille_demande: str
    risk_level: str
    risk_priority: str
    horizon: str
    sla_applicable: str
    sync_billable_part: str
    description_text: str
    story_description_text: str
    acceptance_criteria_text: str
    acceptance_criteria_source: str
    requires_acceptance_criteria: bool
    has_acceptance_criteria: bool
    has_description: bool
    has_story_description: bool
    is_unassigned: bool
    is_overdue: bool
    age_days: Optional[float]
    days_since_update: Optional[float]
    days_to_due: Optional[float]
    lead_time_days: Optional[float]
    comments_count: int
    last_comment_date: Optional[datetime]
    last_comment_author: str
    attachments_count: int
    attachments_total_size_mb: float
    issue_links_count: int
    inward_links_count: int
    outward_links_count: int
    subtasks_count: int
    votes: int
    watches: int
    customfields_filled_count: int
    customfields_total_count: int
    customfields_fill_rate: Optional[float]
    bool_ac_missing: bool
    bool_description_missing: bool
    bool_story_description_missing: bool
    bool_assignee_missing: bool
    bool_sprint_missing: bool
    bool_theme_missing: bool
    bool_target_type_missing: bool
    bool_taille_demande_missing: bool
    bool_risk_level_missing: bool
    bool_risk_priority_missing: bool
    bool_sla_missing: bool
    bool_due_date_missing: bool
    bool_no_comment: bool
    bool_no_attachment: bool
    bool_no_jira_link: bool
    bool_no_subtask: bool
    bool_overdue: bool
    bool_stale_7d: bool
    bool_stale_14d: bool
    bool_stale_30d: bool
    bool_stale_60d: bool
    bool_stale_90d: bool
    bool_needs_attention: bool
    missing_fields_count: int
    missing_fields_list: str
    quality_score: int
    quality_flags: str
    attention_level: str


@dataclass
class Metadata:
    source_file: str
    rss_generated_raw: str
    rss_generated_utc: Optional[datetime]
    rss_generated_paris: str
    analysis_date_utc: datetime
    analysis_date_source: str
    channel_title: str
    channel_link: str
    issue_total: str
    jira_version: str
    jira_build_number: str
    jira_build_date: str


def read_text_safely(path: Path) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            return path.read_text(encoding=enc, errors="strict")
        except UnicodeDecodeError:
            pass
    return path.read_text(encoding="utf-8", errors="replace")


def text_of(el: Optional[ET.Element]) -> str:
    return "" if el is None else "".join(el.itertext())


def unescape_repeat(value: str, max_rounds: int = 6) -> str:
    previous = "" if value is None else str(value)
    for _ in range(max_rounds):
        current = html.unescape(previous)
        if current == previous:
            break
        previous = current
    return previous


def clean_rich_text(value: str) -> str:
    value = unescape_repeat(value or "")
    value = re.sub(r"(?i)<\s*br\s*/?\s*>", "\n", value)
    value = re.sub(r"(?i)</\s*(p|li|tr|th|td|div|h\d|ul|ol|table|tbody)\s*>", "\n", value)
    value = re.sub(r"<[^>]+>", " ", value)
    value = unescape_repeat(value).replace("\xa0", " ")
    value = CONTROL_CHARS_RE.sub("", value)
    value = re.sub(r"[ \t\r\f\v]+", " ", value)
    value = re.sub(r"\n\s*\n+", "\n", value)
    return value.strip()


def strip_accents_light(value: str) -> str:
    return value.translate(str.maketrans({
        "à":"a","â":"a","ä":"a","á":"a","ã":"a","å":"a","ç":"c",
        "é":"e","è":"e","ê":"e","ë":"e","î":"i","ï":"i","í":"i","ì":"i",
        "ô":"o","ö":"o","ó":"o","ò":"o","õ":"o","ù":"u","û":"u","ü":"u","ú":"u",
        "ÿ":"y","ñ":"n","œ":"oe","æ":"ae","À":"a","Â":"a","Ä":"a","Á":"a",
        "Ã":"a","Å":"a","Ç":"c","É":"e","È":"e","Ê":"e","Ë":"e","Î":"i",
        "Ï":"i","Í":"i","Ì":"i","Ô":"o","Ö":"o","Ó":"o","Ò":"o","Õ":"o",
        "Ù":"u","Û":"u","Ü":"u","Ú":"u","Ÿ":"y","Ñ":"n","Œ":"oe","Æ":"ae",
    }))


def normalize_for_check(value: str) -> str:
    value = strip_accents_light(clean_rich_text(value).lower())
    value = re.sub(r"[’'`´]", "'", value)
    value = re.sub(r"[^a-z0-9' ]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def is_empty_or_not_real(value: str) -> bool:
    cleaned = normalize_for_check(value)
    return cleaned in EMPTY_VALUES or cleaned in NOT_REAL_VALUES


def is_meaningful_text(value: str, min_chars: int = 25) -> bool:
    cleaned = normalize_for_check(value)
    if cleaned in EMPTY_VALUES:
        return False
    reduced = cleaned
    for word in sorted(TEMPLATE_WORDS, key=len, reverse=True):
        reduced = reduced.replace(strip_accents_light(word.lower()), " ")
    reduced = re.sub(r"\b(test|scenario|scenar|etant donne|alors|quand)\b", " ", reduced)
    reduced = re.sub(r"\s+", " ", reduced).strip()
    words = [w for w in reduced.split() if len(w) >= 3]
    return len(reduced) >= min_chars and len(words) >= 4


def parse_date(value: str) -> Optional[datetime]:
    value = (value or "").strip()
    if not value:
        return None
    try:
        dt = parsedate_to_datetime(value)
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    except Exception:
        pass
    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    except Exception:
        pass
    for fmt in ("%Y-%m-%d %H:%M:%S%z", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            dt = datetime.strptime(value, fmt)
            return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
        except Exception:
            pass
    return None


def days_between(start: Optional[datetime], end: Optional[datetime]) -> Optional[float]:
    if not start or not end:
        return None
    if start.tzinfo is None:
        start = start.replace(tzinfo=timezone.utc)
    if end.tzinfo is None:
        end = end.replace(tzinfo=timezone.utc)
    return round((end.astimezone(timezone.utc) - start.astimezone(timezone.utc)).total_seconds() / 86400, 2)


def dt_for_excel(value: Optional[datetime]) -> Any:
    if value is None:
        return None
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc).replace(tzinfo=None)


def to_paris_string(dt: Optional[datetime]) -> str:
    if not dt:
        return ""
    if ZoneInfo is None:
        return dt.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    return dt.astimezone(ZoneInfo("Europe/Paris")).strftime("%Y-%m-%d %H:%M:%S %Z")


def safe_int(value: str, default: int = 0) -> int:
    try:
        return int(str(value).strip())
    except Exception:
        return default


def safe_float(value: str, default: float = 0.0) -> float:
    try:
        return float(str(value).replace(",", "."))
    except Exception:
        return default


def excel_safe(value: Any) -> Any:
    if isinstance(value, datetime):
        return dt_for_excel(value)
    if isinstance(value, (int, float, bool)) or value is None:
        return value
    text = CONTROL_CHARS_RE.sub("", str(value))
    if len(text) > MAX_EXCEL_TEXT:
        text = text[:MAX_EXCEL_TEXT] + "... [TRUNCATED]"
    return "'" + text if text.startswith("=") else text


def parse_xml(xml_path: Path) -> Tuple[ET.Element, str]:
    raw = read_text_safely(xml_path)
    try:
        root = ET.fromstring(raw.encode("utf-8"))
    except Exception:
        root = ET.parse(xml_path).getroot()
    return root, raw


def extract_rss_generated_date(raw_xml: str) -> Tuple[str, Optional[datetime]]:
    m = re.search(r"RSS generated by JIRA.*? at ([A-Z][a-z]{2} [A-Z][a-z]{2} \d{2} \d{2}:\d{2}:\d{2} UTC \d{4})", raw_xml, re.S)
    if not m:
        return "", None
    raw = m.group(1).strip()
    try:
        return raw, datetime.strptime(raw, "%a %b %d %H:%M:%S UTC %Y").replace(tzinfo=timezone.utc)
    except Exception:
        return raw, None


def parse_as_of(user_value: Optional[str], raw_xml: str) -> Tuple[datetime, str, str, Optional[datetime]]:
    rss_raw, rss_dt = extract_rss_generated_date(raw_xml)
    if user_value:
        parsed = parse_date(user_value)
        if not parsed:
            raise SystemExit(f"Date --as-of non comprise: {user_value}")
        return parsed, "Parametre utilisateur --as-of", rss_raw, rss_dt
    if rss_dt:
        return rss_dt, "Header RSS generated by JIRA", rss_raw, rss_dt
    return datetime.now(timezone.utc), "Date systeme fallback", rss_raw, rss_dt


def get_child_text(item: Optional[ET.Element], tag: str) -> str:
    return "" if item is None else clean_rich_text(text_of(item.find(tag)))


def get_child_raw_text(item: ET.Element, tag: str) -> str:
    return unescape_repeat(text_of(item.find(tag))).strip()


def get_attr(el: Optional[ET.Element], attr: str) -> str:
    return "" if el is None else str(el.attrib.get(attr, ""))


def extract_metadata(xml_path: Path, root: ET.Element, raw_xml: str, as_of: datetime, source: str, rss_raw: str, rss_dt: Optional[datetime]) -> Metadata:
    channel = root.find("channel")
    issue = channel.find("issue") if channel is not None else None
    build = channel.find("build-info") if channel is not None else None
    return Metadata(
        source_file=str(xml_path),
        rss_generated_raw=rss_raw,
        rss_generated_utc=rss_dt,
        rss_generated_paris=to_paris_string(rss_dt),
        analysis_date_utc=as_of,
        analysis_date_source=source,
        channel_title=get_child_text(channel, "title"),
        channel_link=get_child_text(channel, "link"),
        issue_total=issue.attrib.get("total", "") if issue is not None else "",
        jira_version=get_child_text(build, "version"),
        jira_build_number=get_child_text(build, "build-number"),
        jira_build_date=get_child_text(build, "build-date"),
    )


def extract_project(item: ET.Element) -> Tuple[str, str]:
    el = item.find("project")
    return get_attr(el, "key"), clean_rich_text(text_of(el))


def extract_labels(item: ET.Element) -> str:
    labels_el = item.find("labels")
    if labels_el is None:
        return ""
    return "; ".join([clean_rich_text(text_of(x)) for x in labels_el.findall("label") if clean_rich_text(text_of(x))])


def extract_customfields(item: ET.Element) -> Dict[str, Dict[str, Any]]:
    result: Dict[str, Dict[str, Any]] = {}
    for cf in item.findall("./customfields/customfield"):
        name = clean_rich_text(text_of(cf.find("customfieldname"))).strip() or cf.attrib.get("id", "unknown_customfield")
        values = []
        for val_el in cf.findall("./customfieldvalues/customfieldvalue"):
            txt = clean_rich_text(text_of(val_el))
            if txt:
                values.append(txt)
        payload = {"id": cf.attrib.get("id", ""), "key": cf.attrib.get("key", ""), "values": values, "text": "; ".join(values).strip()}
        if name in result and payload["text"]:
            result[name]["values"].extend(values)
            result[name]["text"] = "; ".join([x for x in [result[name].get("text", ""), payload["text"]] if x])
        else:
            result[name] = payload
    return result


def get_customfield(customfields: Dict[str, Dict[str, Any]], *names: str) -> str:
    wanted = {normalize_for_check(n) for n in names}
    for name, payload in customfields.items():
        if normalize_for_check(name) in wanted:
            return payload.get("text", "")
    return ""


def find_story_acceptance_criteria(story_text: str) -> str:
    text = clean_rich_text(story_text)
    if not text:
        return ""
    lowered = strip_accents_light(text.lower())
    starts = [(lowered.find(m), len(m)) for m in ["acceptance criteria", "critere d'acceptation", "criteres d'acceptation", "criteria"] if lowered.find(m) != -1]
    if not starts:
        return ""
    start_idx, marker_len = sorted(starts)[0]
    tail = text[start_idx + marker_len:]
    tail_norm = strip_accents_light(tail.lower())
    positions = [tail_norm.find(m) for m in ["\ntest 1", "\ngiven", "\nwhen", "\nthen", "\nadditional information", "\ninformations complementaires"] if tail_norm.find(m) != -1]
    if positions:
        tail = tail[:min(positions)]
    return tail.strip(" :\n\t")


def extract_acceptance_criteria(customfields: Dict[str, Dict[str, Any]], description_text: str) -> Tuple[str, str, bool]:
    dedicated = get_customfield(customfields, "Acceptance Criteria", "Critères d'acceptation", "Criteres d'acceptation")
    if is_meaningful_text(dedicated):
        return dedicated, "customfield Acceptance Criteria", True
    story = get_customfield(customfields, "Story Description", "Description Story")
    story_ac = find_story_acceptance_criteria(story)
    if is_meaningful_text(story_ac):
        return story_ac, "Story Description / Acceptance criteria", True
    desc_norm = normalize_for_check(description_text)
    markers = ["acceptance criteria", "critere d'acceptation", "criteres d'acceptation", "resultat attendu", "resultats attendus", "expected result", "given", "when", "then", "etant donne"]
    if any(m in desc_norm for m in markers) and is_meaningful_text(description_text, min_chars=80):
        return description_text, "description markers", True
    if dedicated:
        return dedicated, "customfield Acceptance Criteria vide/template", False
    if story:
        return story_ac or story, "Story Description vide/template", False
    return "", "absent", False


def requires_ac(issue_type: str) -> bool:
    return normalize_for_check(issue_type) not in {"test execution", "xray test execution", "sub test execution", "xray test"}


def extract_comments(item: ET.Element) -> List[Dict[str, Any]]:
    return [{"id": c.attrib.get("id", ""), "author": c.attrib.get("author", ""), "created": parse_date(c.attrib.get("created", "")), "body": clean_rich_text(text_of(c))} for c in item.findall("./comments/comment")]


def extract_attachments(item: ET.Element) -> List[Dict[str, Any]]:
    return [{"id": a.attrib.get("id", ""), "name": a.attrib.get("name", ""), "size": safe_float(a.attrib.get("size", "0"), 0.0), "author": a.attrib.get("author", ""), "created": parse_date(a.attrib.get("created", ""))} for a in item.findall("./attachments/attachment")]


def extract_issue_links(item: ET.Element) -> Tuple[int, int, int]:
    inward = 0
    outward = 0
    for linktype in item.findall("./issuelinks/issuelinktype"):
        inward += len(linktype.findall("./inwardlinks/issuelink"))
        outward += len(linktype.findall("./outwardlinks/issuelink"))
    return inward + outward, inward, outward


def compute_quality(missing: Dict[str, bool], issue_type: str) -> Tuple[int, str, int, str, str, bool]:
    field_points = {
        "Acceptance Criteria": ("bool_ac_missing", 25),
        "Description": ("bool_description_missing", 12),
        "Story Description": ("bool_story_description_missing", 8),
        "Assignee": ("bool_assignee_missing", 12),
        "Sprint": ("bool_sprint_missing", 8),
        "Theme": ("bool_theme_missing", 5),
        "Target Type": ("bool_target_type_missing", 5),
        "Taille de la demande": ("bool_taille_demande_missing", 5),
        "Risk level": ("bool_risk_level_missing", 5),
        "Risk priority": ("bool_risk_priority_missing", 5),
        "SLA applicable": ("bool_sla_missing", 3),
    }
    score = 100
    missing_fields = []
    flags = []
    for label, (flag, points) in field_points.items():
        if missing.get(flag):
            missing_fields.append(label)
            flags.append(f"{label} missing")
            score -= points
    if missing.get("bool_overdue"):
        score -= 12; flags.append("Overdue")
    if missing.get("bool_stale_30d"):
        score -= 10; flags.append("No update >=30d")
    if missing.get("bool_no_comment"):
        score -= 4; flags.append("No comment")
    if missing.get("bool_no_attachment") and normalize_for_check(issue_type) == "bug":
        score -= 4; flags.append("Bug without attachment")
    if missing.get("bool_no_jira_link"):
        score -= 3; flags.append("No Jira link")
    score = max(score, 0)
    count = len(missing_fields)
    needs_attention = count >= 2 or bool(missing.get("bool_ac_missing")) or bool(missing.get("bool_overdue")) or bool(missing.get("bool_stale_30d"))
    level = "HIGH" if score < 55 or missing.get("bool_ac_missing") or count >= 5 else ("MEDIUM" if score < 75 or count >= 3 or missing.get("bool_stale_30d") else "LOW")
    return score, "; ".join(flags), count, "; ".join(missing_fields), level, needs_attention


def parse_tickets(root: ET.Element, as_of: datetime) -> Tuple[List[Ticket], List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]]]:
    tickets, comments_flat, attachments_flat, customfields_flat = [], [], [], []
    for item in root.findall("./channel/item"):
        key_el = item.find("key")
        key = clean_rich_text(text_of(key_el))
        project_key = clean_rich_text(text_of(item.find("key")))
        project_name = extract_project(item)[1]
        customfields = extract_customfields(item)
        description_text = clean_rich_text(text_of(item.find("description")))
        story_description_text = get_customfield(customfields, "Story Description", "Description Story")
        ac_text, ac_source, has_ac = extract_acceptance_criteria(customfields, description_text)
        created, updated, resolved, due = [parse_date(text_of(item.find(x))) for x in ("created", "updated", "resolved", "due")]
        comments = extract_comments(item)
        attachments = extract_attachments(item)
        for c in comments:
            comments_flat.append({"Ticket": key, "Comment ID": c["id"], "Author": c["author"], "Created": c["created"], "Body": c["body"]})
        for a in attachments:
            attachments_flat.append({"Ticket": key, "Attachment ID": a["id"], "Name": a["name"], "Size MB": round(a["size"]/(1024*1024), 3), "Author": a["author"], "Created": a["created"]})
        for cf_name, payload in customfields.items():
            val = payload.get("text", "")
            customfields_flat.append({"Ticket": key, "Custom field": cf_name, "Custom field id": payload.get("id", ""), "Value": val, "Is filled": is_meaningful_text(val, 2) and not is_empty_or_not_real(val)})
        issue_links_count, inward_links_count, outward_links_count = extract_issue_links(item)
        cf_total = len(customfields)
        cf_filled = sum(1 for p in customfields.values() if is_meaningful_text(p.get("text", ""), 2) and not is_empty_or_not_real(p.get("text", "")))
        issue_type = get_child_text(item, "type")
        priority = get_child_text(item, "priority")
        status = get_child_text(item, "status")
        status_cat_el = item.find("statusCategory")
        status_category = status_cat_el.attrib.get("key", "") if status_cat_el is not None else ""
        assignee = get_child_text(item, "assignee")
        reporter = get_child_text(item, "reporter")
        sprint = get_customfield(customfields, "Sprint")
        theme = get_customfield(customfields, "Theme")
        target_type = get_customfield(customfields, "Target Type")
        taille_demande = get_customfield(customfields, "Taille de la demande")
        risk_level = get_customfield(customfields, "Risk level")
        risk_priority = get_customfield(customfields, "Risk priority")
        horizon = get_customfield(customfields, "Horizon")
        sla_applicable = get_customfield(customfields, "SLA applicable")
        sync_billable_part = get_customfield(customfields, "Sync billable part")
        req_ac = requires_ac(issue_type)
        has_desc = is_meaningful_text(description_text, 20)
        has_story = is_meaningful_text(story_description_text, 25)
        is_unassigned = normalize_for_check(assignee) in {"", "non assigne", "unassigned"} or get_attr(item.find("assignee"), "accountid") == "-1"
        days_update = days_between(updated, as_of)
        is_overdue = bool(due and due.date() < as_of.date() and status_category != "done")
        subtasks_count = len(item.findall("./subtasks/subtask"))
        last_comment_date = max([c["created"] for c in comments if c.get("created")], default=None)
        last_comment_author = next((c["author"] for c in reversed(comments) if c.get("created") == last_comment_date), "") if last_comment_date else ""
        missing = {
            "bool_ac_missing": bool(req_ac and not has_ac),
            "bool_description_missing": not has_desc,
            "bool_story_description_missing": bool(req_ac and not has_story),
            "bool_assignee_missing": is_unassigned,
            "bool_sprint_missing": is_empty_or_not_real(sprint),
            "bool_theme_missing": is_empty_or_not_real(theme),
            "bool_target_type_missing": is_empty_or_not_real(target_type),
            "bool_taille_demande_missing": is_empty_or_not_real(taille_demande),
            "bool_risk_level_missing": is_empty_or_not_real(risk_level),
            "bool_risk_priority_missing": is_empty_or_not_real(risk_priority),
            "bool_sla_missing": is_empty_or_not_real(sla_applicable),
            "bool_due_date_missing": due is None,
            "bool_no_comment": len(comments) == 0,
            "bool_no_attachment": len(attachments) == 0,
            "bool_no_jira_link": issue_links_count == 0,
            "bool_no_subtask": subtasks_count == 0,
            "bool_overdue": is_overdue,
            "bool_stale_7d": bool(days_update is not None and days_update >= 7),
            "bool_stale_14d": bool(days_update is not None and days_update >= 14),
            "bool_stale_30d": bool(days_update is not None and days_update >= 30),
            "bool_stale_60d": bool(days_update is not None and days_update >= 60),
            "bool_stale_90d": bool(days_update is not None and days_update >= 90),
        }
        score, flags, miss_count, miss_list, attention_level, needs_attention = compute_quality(missing, issue_type)
        tickets.append(Ticket(
            key=key, id=get_attr(key_el, "id"), link=get_child_raw_text(item, "link"), project_key=project_key, project_name=project_name,
            title=get_child_text(item, "title"), summary=get_child_text(item, "summary"), issue_type=issue_type, priority=priority, status=status,
            status_category=status_category, resolution=get_child_text(item, "resolution"), assignee=assignee, reporter=reporter,
            created=created, updated=updated, resolved=resolved, due=due, labels=extract_labels(item), parent=get_child_text(item, "parent"),
            sprint=sprint, theme=theme, target_type=target_type, taille_demande=taille_demande, risk_level=risk_level,
            risk_priority=risk_priority, horizon=horizon, sla_applicable=sla_applicable, sync_billable_part=sync_billable_part,
            description_text=description_text, story_description_text=story_description_text, acceptance_criteria_text=clean_rich_text(ac_text), acceptance_criteria_source=ac_source,
            requires_acceptance_criteria=req_ac, has_acceptance_criteria=has_ac, has_description=has_desc, has_story_description=has_story,
            is_unassigned=is_unassigned, is_overdue=is_overdue, age_days=days_between(created, as_of), days_since_update=days_update,
            days_to_due=days_between(as_of, due) if due else None, lead_time_days=days_between(created, resolved) if resolved else None,
            comments_count=len(comments), last_comment_date=last_comment_date, last_comment_author=last_comment_author,
            attachments_count=len(attachments), attachments_total_size_mb=round(sum(a["size"] for a in attachments)/(1024*1024), 3),
            issue_links_count=issue_links_count, inward_links_count=inward_links_count, outward_links_count=outward_links_count,
            subtasks_count=subtasks_count, votes=safe_int(get_child_text(item, "votes")), watches=safe_int(get_child_text(item, "watches")),
            customfields_filled_count=cf_filled, customfields_total_count=cf_total, customfields_fill_rate=round(cf_filled/cf_total, 4) if cf_total else None,
            **missing, bool_needs_attention=needs_attention, missing_fields_count=miss_count, missing_fields_list=miss_list,
            quality_score=score, quality_flags=flags, attention_level=attention_level
        ))
    return tickets, comments_flat, attachments_flat, customfields_flat


def pct(n: int, d: int) -> float:
    return round(n / d, 4) if d else 0.0


def avg(values: Iterable[Optional[float]]) -> Optional[float]:
    vals = [float(v) for v in values if v is not None]
    return round(sum(vals) / len(vals), 2) if vals else None


def median(values: Iterable[Optional[float]]) -> Optional[float]:
    vals = sorted(float(v) for v in values if v is not None)
    return round(statistics.median(vals), 2) if vals else None


def p90(values: Iterable[Optional[float]]) -> Optional[float]:
    vals = sorted(float(v) for v in values if v is not None)
    if not vals:
        return None
    return round(vals[int(round(0.9 * (len(vals)-1)))], 2)


def group_summary(tickets: List[Ticket], field: str) -> List[Dict[str, Any]]:
    groups: Dict[str, List[Ticket]] = defaultdict(list)
    for t in tickets:
        val = getattr(t, field) or "(vide)"
        groups[str(val)].append(t)
    rows = []
    for val, items in sorted(groups.items(), key=lambda kv: (-sum(t.bool_ac_missing for t in kv[1]), -len(kv[1]), kv[0].lower())):
        req = [t for t in items if t.requires_acceptance_criteria]
        rows.append({
            field: val,
            "Tickets": len(items),
            "AC missing": sum(t.bool_ac_missing for t in items),
            "% AC missing": pct(sum(t.bool_ac_missing for t in items), len(req)),
            "Needs attention": sum(t.bool_needs_attention for t in items),
            "Avg missing fields": avg(t.missing_fields_count for t in items),
            "Unassigned": sum(t.bool_assignee_missing for t in items),
            "Sprint missing": sum(t.bool_sprint_missing for t in items),
            "No update >=30d": sum(t.bool_stale_30d for t in items),
            "Avg age days": avg(t.age_days for t in items),
            "Avg quality score": avg(t.quality_score for t in items),
        })
    return rows


def missing_by_field(tickets: List[Ticket]) -> List[Dict[str, Any]]:
    fields = [
        ("Acceptance Criteria", "bool_ac_missing"), ("Description", "bool_description_missing"),
        ("Story Description", "bool_story_description_missing"), ("Assignee", "bool_assignee_missing"),
        ("Sprint", "bool_sprint_missing"), ("Theme", "bool_theme_missing"),
        ("Target Type", "bool_target_type_missing"), ("Taille de la demande", "bool_taille_demande_missing"),
        ("Risk level", "bool_risk_level_missing"), ("Risk priority", "bool_risk_priority_missing"),
        ("SLA applicable", "bool_sla_missing"), ("Due date", "bool_due_date_missing"),
        ("Comment", "bool_no_comment"), ("Attachment", "bool_no_attachment"),
        ("Jira link", "bool_no_jira_link"), ("Subtask", "bool_no_subtask"),
    ]
    total = len(tickets)
    rows = []
    for label, attr in fields:
        n = sum(bool(getattr(t, attr)) for t in tickets)
        rows.append({"Field": label, "Missing count": n, "% missing": pct(n, total)})
    return sorted(rows, key=lambda r: (-r["Missing count"], r["Field"]))


def field_coverage(customfields_flat: List[Dict[str, Any]], tickets_count: int) -> List[Dict[str, Any]]:
    by_field: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for row in customfields_flat:
        by_field[row["Custom field"]].append(row)
    rows = []
    for name, items in sorted(by_field.items(), key=lambda kv: kv[0].lower()):
        filled = sum(1 for r in items if r["Is filled"])
        rows.append({
            "Custom field": name,
            "Present on tickets": len(items),
            "Filled count": filled,
            "Missing/empty count": len(items)-filled,
            "% filled vs present": pct(filled, len(items)),
            "% filled vs all tickets": pct(filled, tickets_count),
        })
    return rows


def ticket_rows(tickets: List[Ticket], compact: bool = False) -> List[Dict[str, Any]]:
    rows = []
    for t in tickets:
        base = asdict(t)
        if compact:
            keep = [
                "key", "link", "summary", "project_key", "issue_type", "status", "priority", "assignee", "reporter", "sprint", "theme",
                "target_type", "taille_demande", "created", "updated", "age_days", "days_since_update", "requires_acceptance_criteria",
                "has_acceptance_criteria", "bool_ac_missing", "bool_description_missing", "bool_story_description_missing", "bool_assignee_missing",
                "bool_sprint_missing", "bool_theme_missing", "bool_target_type_missing", "bool_taille_demande_missing", "bool_risk_level_missing",
                "bool_risk_priority_missing", "bool_sla_missing", "bool_due_date_missing", "bool_no_comment", "bool_no_attachment",
                "bool_no_jira_link", "bool_no_subtask", "bool_overdue", "bool_stale_7d", "bool_stale_14d", "bool_stale_30d",
                "bool_stale_60d", "bool_stale_90d", "bool_needs_attention", "missing_fields_count", "missing_fields_list",
                "quality_score", "quality_flags", "attention_level"
            ]
            rows.append({k: base[k] for k in keep})
        else:
            rows.append(base)
    return rows


def metadata_rows(meta: Metadata) -> List[Dict[str, Any]]:
    return [
        {"Field": "Source file", "Value": meta.source_file},
        {"Field": "RSS generated raw", "Value": meta.rss_generated_raw},
        {"Field": "RSS generated UTC", "Value": meta.rss_generated_utc},
        {"Field": "RSS generated Europe/Paris", "Value": meta.rss_generated_paris},
        {"Field": "Analysis date UTC", "Value": meta.analysis_date_utc},
        {"Field": "Analysis date source", "Value": meta.analysis_date_source},
        {"Field": "Channel title", "Value": meta.channel_title},
        {"Field": "Channel link", "Value": meta.channel_link},
        {"Field": "Issue total", "Value": meta.issue_total},
        {"Field": "Jira version", "Value": meta.jira_version},
        {"Field": "Jira build number", "Value": meta.jira_build_number},
        {"Field": "Jira build date", "Value": meta.jira_build_date},
        {"Field": "Excel structured tables", "Value": "Disabled: no /xl/tables/table*.xml"},
    ]


def dashboard_rows(tickets: List[Ticket], meta: Metadata) -> List[Dict[str, Any]]:
    total = len(tickets)
    req = [t for t in tickets if t.requires_acceptance_criteria]
    open_t = [t for t in tickets if t.status_category != "done"]
    rows = [
        {"KPI": "Date export Jira UTC", "Value": meta.rss_generated_utc, "Comment": "Header RSS generated by JIRA"},
        {"KPI": "Date export Jira Europe/Paris", "Value": meta.rss_generated_paris, "Comment": "Conversion locale"},
        {"KPI": "Tickets total", "Value": total, "Comment": "Nombre de balises item"},
        {"KPI": "Tickets ouverts/non done", "Value": len(open_t), "Comment": "statusCategory different de done"},
        {"KPI": "Tickets necessitant AC", "Value": len(req), "Comment": "Hors types de test execution"},
        {"KPI": "Tickets sans AC", "Value": sum(t.bool_ac_missing for t in tickets), "Comment": "AC absent/vide/template"},
        {"KPI": "% sans AC", "Value": pct(sum(t.bool_ac_missing for t in tickets), len(req)), "Comment": "Base: tickets necessitant AC"},
        {"KPI": "Tickets avec donnees manquantes", "Value": sum(t.bool_needs_attention for t in tickets), "Comment": "Needs attention = TRUE"},
        {"KPI": "Non assignes", "Value": sum(t.bool_assignee_missing for t in tickets), "Comment": "Assignee vide/non assigne"},
        {"KPI": "Sprint manquant", "Value": sum(t.bool_sprint_missing for t in tickets), "Comment": "Sprint vide ou non exploitable"},
        {"KPI": "Sans mise a jour >=30j", "Value": sum(t.bool_stale_30d for t in tickets), "Comment": "updated ancien"},
        {"KPI": "Age moyen tickets ouverts", "Value": avg(t.age_days for t in open_t), "Comment": "En jours"},
        {"KPI": "P90 age tickets ouverts", "Value": p90(t.age_days for t in open_t), "Comment": "En jours"},
        {"KPI": "Score qualite moyen", "Value": avg(t.quality_score for t in tickets), "Comment": "100 = complet"},
    ]
    return rows


def write_csvs(output_base: Path, sheets: Dict[str, List[Dict[str, Any]]]) -> None:
    out_dir = output_base.with_suffix("").as_posix() + "_csv"
    os.makedirs(out_dir, exist_ok=True)
    for name, rows in sheets.items():
        path = Path(out_dir) / f"{name}.csv"
        headers = list(rows[0].keys()) if rows else ["empty"]
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
            w.writeheader()
            for row in rows:
                w.writerow({k: excel_safe(v) for k, v in row.items()})
    print(f"XlsxWriter non installe: CSV ecrits dans {out_dir}")


def write_table_sheet(workbook: Any, sheet_name: str, rows: List[Dict[str, Any]], formats: Dict[str, Any], autofilter: bool = True) -> Any:
    ws = workbook.add_worksheet(sheet_name[:31])
    if not rows:
        ws.write(0, 0, "No data", formats["note"])
        return ws
    headers = list(rows[0].keys())
    for c, h in enumerate(headers):
        ws.write(0, c, h, formats["header"])
    for r, row in enumerate(rows, start=1):
        for c, h in enumerate(headers):
            val = excel_safe(row.get(h))
            fmt = None
            if isinstance(val, bool):
                fmt = formats["bool"]
            elif isinstance(val, datetime):
                fmt = formats["date"]
            elif h.lower().startswith("%") or h.lower().endswith("%") or h.lower().startswith("pct") or "%" in h:
                fmt = formats["pct"]
            elif "score" in h.lower() or "count" in h.lower() or "tickets" in h.lower() or "age" in h.lower() or "days" in h.lower():
                fmt = formats["number"]
            ws.write(r, c, val, fmt)
    ws.freeze_panes(1, 0)
    if autofilter:
        ws.autofilter(0, 0, max(0, len(rows)), max(0, len(headers)-1))
    widths = {}
    for idx, h in enumerate(headers):
        max_len = len(str(h))
        for row in rows[:300]:
            max_len = max(max_len, len(str(row.get(h, "") or "")))
        width = min(max(max_len + 2, 10), 45)
        if any(x in h.lower() for x in ["summary", "description", "criteria", "flags", "fields", "link", "value", "body"]):
            width = min(max(width, 25), 65)
        widths[idx] = width
        ws.set_column(idx, idx, width)
    ws.conditional_format(1, 0, len(rows), len(headers)-1, {"type": "formula", "criteria": '=$A2=""', "format": formats["blank"]})
    for c, h in enumerate(headers):
        if h.startswith("bool_") or h in {"has_acceptance_criteria", "requires_acceptance_criteria", "Is filled"}:
            ws.conditional_format(1, c, len(rows), c, {"type": "cell", "criteria": "==", "value": True, "format": formats["true"]})
            ws.conditional_format(1, c, len(rows), c, {"type": "cell", "criteria": "==", "value": False, "format": formats["false"]})
        if h == "attention_level":
            ws.conditional_format(1, c, len(rows), c, {"type": "text", "criteria": "containing", "value": "HIGH", "format": formats["high"]})
            ws.conditional_format(1, c, len(rows), c, {"type": "text", "criteria": "containing", "value": "MEDIUM", "format": formats["medium"]})
            ws.conditional_format(1, c, len(rows), c, {"type": "text", "criteria": "containing", "value": "LOW", "format": formats["low"]})
        if h == "quality_score":
            ws.conditional_format(1, c, len(rows), c, {"type": "3_color_scale"})
        if h == "missing_fields_count":
            ws.conditional_format(1, c, len(rows), c, {"type": "data_bar", "bar_color": "#C00000"})
    return ws


def add_dashboard(workbook: Any, sheets: Dict[str, List[Dict[str, Any]]], formats: Dict[str, Any], autofilter: bool) -> None:
    ws = workbook.add_worksheet("Dashboard")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 26)
    ws.set_column("B:B", 18)
    ws.set_column("C:C", 62)
    ws.merge_range("A1:H1", "Jira RSS - Dashboard donnees manquantes", formats["title"])
    ws.write("A3", "KPI", formats["header"])
    ws.write("B3", "Value", formats["header"])
    ws.write("C3", "Comment", formats["header"])
    for r, row in enumerate(sheets["Dashboard_Data"], start=3):
        ws.write(r, 0, row["KPI"], formats["kpi_label"])
        val = excel_safe(row["Value"])
        fmt = formats["pct"] if "%" in row["KPI"] else (formats["date"] if isinstance(val, datetime) else formats["kpi_value"])
        ws.write(r, 1, val, fmt)
        ws.write(r, 2, row["Comment"], formats["note"])
    ws.write("E3", "Lecture rapide", formats["header"])
    notes = [
        "1. La feuille Boolean_Matrix contient les colonnes bool_* pour trier/filtrer les manques.",
        "2. La feuille AC_Missing isole directement les tickets sans Acceptance Criteria.",
        "3. Les feuilles By_Sprint / By_Assignee permettent de prioriser les relances.",
        "4. Aucune Table Excel structuree n'est creee: pas de /xl/tables/table*.xml.",
    ]
    for i, note in enumerate(notes, start=4):
        ws.write(i, 4, note, formats["note_box"])


def add_chart_to_dashboard(workbook: Any, sheet_name: str, category_col: int, value_col: int, rows_count: int, chart_title: str, pos: str) -> None:
    if rows_count < 2:
        return
    dash = workbook.get_worksheet_by_name("Dashboard")
    chart = workbook.add_chart({"type": "column"})
    chart.add_series({
        "name": chart_title,
        "categories": [sheet_name, 1, category_col, min(rows_count, 12), category_col],
        "values": [sheet_name, 1, value_col, min(rows_count, 12), value_col],
        "data_labels": {"value": True},
        "fill": {"color": "#C00000"},
    })
    chart.set_title({"name": chart_title})
    chart.set_legend({"none": True})
    chart.set_y_axis({"major_gridlines": {"visible": False}})
    chart.set_style(10)
    dash.insert_chart(pos, chart, {"x_scale": 1.15, "y_scale": 1.05})


def make_formats(workbook: Any) -> Dict[str, Any]:
    return {
        "title": workbook.add_format({"bold": True, "font_size": 18, "font_color": "white", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter"}),
        "header": workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True}),
        "kpi_label": workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1}),
        "kpi_value": workbook.add_format({"bold": True, "font_size": 12, "bg_color": "#F2F2F2", "border": 1}),
        "note": workbook.add_format({"text_wrap": True, "valign": "top"}),
        "note_box": workbook.add_format({"text_wrap": True, "valign": "top", "bg_color": "#FFF2CC", "border": 1}),
        "bool": workbook.add_format({"align": "center"}),
        "true": workbook.add_format({"bg_color": "#F4CCCC", "font_color": "#990000", "bold": True, "align": "center"}),
        "false": workbook.add_format({"bg_color": "#D9EAD3", "font_color": "#274E13", "align": "center"}),
        "high": workbook.add_format({"bg_color": "#C00000", "font_color": "white", "bold": True}),
        "medium": workbook.add_format({"bg_color": "#F4B183", "font_color": "#7F3F00", "bold": True}),
        "low": workbook.add_format({"bg_color": "#C6E0B4", "font_color": "#274E13", "bold": True}),
        "date": workbook.add_format({"num_format": "yyyy-mm-dd hh:mm", "border": 0}),
        "pct": workbook.add_format({"num_format": "0.0%"}),
        "number": workbook.add_format({"num_format": "0.00"}),
        "blank": workbook.add_format({"bg_color": "#FCE4D6"}),
    }


def write_excel(path: Path, sheets: Dict[str, List[Dict[str, Any]]], autofilter: bool = True) -> None:
    if xlsxwriter is None:
        write_csvs(path, sheets)
        return
    workbook = xlsxwriter.Workbook(str(path), {"strings_to_urls": False, "constant_memory": False})
    formats = make_formats(workbook)
    add_dashboard(workbook, sheets, formats, autofilter)
    for name in ["Metadata", "Boolean_Matrix", "AC_Missing", "By_Sprint", "By_Assignee", "By_Status", "By_Type", "By_Project", "Missing_Data_By_Field", "Field_Coverage", "Raw_Tickets", "CustomFields", "Comments", "Attachments"]:
        write_table_sheet(workbook, name, sheets.get(name, []), formats, autofilter)
    add_chart_to_dashboard(workbook, "By_Sprint", 0, 2, len(sheets.get("By_Sprint", [])), "AC missing by sprint", "E9")
    add_chart_to_dashboard(workbook, "By_Assignee", 0, 2, len(sheets.get("By_Assignee", [])), "AC missing by assignee", "E25")
    add_chart_to_dashboard(workbook, "Missing_Data_By_Field", 0, 1, len(sheets.get("Missing_Data_By_Field", [])), "Missing data by field", "A22")
    workbook.close()
    verify_xlsx(path)
    print(f"Rapport ecrit: {path}")


def verify_xlsx(path: Path) -> None:
    with zipfile.ZipFile(path, "r") as zf:
        tables = [n for n in zf.namelist() if n.startswith("xl/tables/")]
        if tables:
            raise RuntimeError("ERREUR: tables Excel structurees detectees: " + ", ".join(tables))


def build_sheets(xml_path: Path, as_of_arg: Optional[str]) -> Tuple[Dict[str, List[Dict[str, Any]]], Metadata, List[Ticket]]:
    root, raw_xml = parse_xml(xml_path)
    as_of, as_of_source, rss_raw, rss_dt = parse_as_of(as_of_arg, raw_xml)
    meta = extract_metadata(xml_path, root, raw_xml, as_of, as_of_source, rss_raw, rss_dt)
    tickets, comments, attachments, customfields = parse_tickets(root, as_of)
    ac_missing = sorted([t for t in tickets if t.bool_ac_missing], key=lambda t: ((t.sprint or ""), (t.assignee or ""), -(t.age_days or 0), t.key))
    sheets = {
        "Metadata": metadata_rows(meta),
        "Dashboard_Data": dashboard_rows(tickets, meta),
        "Boolean_Matrix": ticket_rows(sorted(tickets, key=lambda t: (not t.bool_ac_missing, t.sprint or "", t.assignee or "", t.key)), compact=True),
        "AC_Missing": ticket_rows(ac_missing, compact=True),
        "By_Sprint": group_summary(tickets, "sprint"),
        "By_Assignee": group_summary(tickets, "assignee"),
        "By_Status": group_summary(tickets, "status"),
        "By_Type": group_summary(tickets, "issue_type"),
        "By_Project": group_summary(tickets, "project_key"),
        "Missing_Data_By_Field": missing_by_field(tickets),
        "Field_Coverage": field_coverage(customfields, len(tickets)),
        "Raw_Tickets": ticket_rows(tickets, compact=False),
        "CustomFields": customfields,
        "Comments": comments,
        "Attachments": attachments,
    }
    return sheets, meta, tickets


def print_summary(tickets: List[Ticket], meta: Metadata) -> None:
    total = len(tickets)
    req = [t for t in tickets if t.requires_acceptance_criteria]
    without_ac = sum(t.bool_ac_missing for t in tickets)
    print("\nResume")
    print("------")
    print(f"Date export Jira UTC          : {meta.rss_generated_utc or 'non trouvee'}")
    print(f"Date export Jira Paris        : {meta.rss_generated_paris or 'non trouvee'}")
    print(f"Tickets analyses              : {total}")
    print(f"Tickets necessitant AC         : {len(req)}")
    print(f"Sans Acceptance Criteria       : {without_ac} ({pct(without_ac, len(req)):.1%})")
    print(f"A traiter / donnees manquantes : {sum(t.bool_needs_attention for t in tickets)} ({pct(sum(t.bool_needs_attention for t in tickets), total):.1%})")
    print(f"Non assignes                   : {sum(t.bool_assignee_missing for t in tickets)} ({pct(sum(t.bool_assignee_missing for t in tickets), total):.1%})")
    print(f"Sprint manquant                : {sum(t.bool_sprint_missing for t in tickets)} ({pct(sum(t.bool_sprint_missing for t in tickets), total):.1%})")
    print("\nNote: le RSS Jira ne contient pas le changelog complet champ par champ.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyse KPI d'un export RSS/XML Jira")
    parser.add_argument("xml", help="Chemin du fichier XML RSS Jira")
    parser.add_argument("--output", "-o", default="jira_kpi_report.xlsx", help="Fichier Excel de sortie")
    parser.add_argument("--as-of", default=None, help="Date d'analyse optionnelle, ex: 2026-06-05 13:55:18+00:00")
    parser.add_argument("--no-autofilter", action="store_true", help="Desactive les filtres simples de feuille Excel")
    args = parser.parse_args()
    xml_path = Path(args.xml)
    if not xml_path.exists():
        raise SystemExit(f"Fichier introuvable: {xml_path}")
    sheets, meta, tickets = build_sheets(xml_path, args.as_of)
    write_excel(Path(args.output), sheets, autofilter=not args.no_autofilter)
    print_summary(tickets, meta)


if __name__ == "__main__":
    main()
