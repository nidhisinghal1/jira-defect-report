#!/usr/bin/env python3
"""
Defect report for open Viasat Jira defects tagged VS3, MCIS, R4.5 or R4.4.1.

The four tags are stored inconsistently in Jira (some are labels, some are
Fix Versions, some may be components), so the script probes each value against
each candidate field and builds its JQL from only the combinations Jira
accepts. An issue is included if it matches ANY of the four (OR).

Usage:
    set JIRA_PAT=<your personal access token>
    python jira_defect_report.py
    python jira_defect_report.py --discover          # just show the field mapping
    python jira_defect_report.py --out report.xlsx --csv

Auth is a Jira Server/Data Center Personal Access Token:
    Jira > Profile > Personal Access Tokens > Create token
"""

from __future__ import annotations

import argparse
import csv
import getpass
import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

BASE_URL = os.environ.get("JIRA_URL", "https://jira.viasat.com").rstrip("/")

# Account the report runs as. Basic auth still needs a secret alongside this -
# an email address alone cannot authenticate.
DEFAULT_EMAIL = os.environ.get("JIRA_EMAIL", "nidhi.singhal@viasat.com")
DEFAULT_AUTH = os.environ.get("JIRA_AUTH_MODE", "bearer")

# The release / stream tags to report on.
TAGS = ["VS3", "MCIS", "R4.5", "R4.4.1"]

# Fields each tag might live in, in probe order. `component`, `fixVersion` and
# `affectedVersion` are the JQL field names (singular), not the REST names.
CANDIDATE_FIELDS = ["labels", "fixVersion", "component", "affectedVersion"]

# What counts as a defect. Validated against the instance before use, so a name
# that does not exist on this Jira is dropped rather than causing a 400.
DEFECT_ISSUE_TYPES = ["Bug", "Defect"]

PAGE_SIZE = 100
REQUEST_TIMEOUT = 60
MAX_RETRIES = 4

ISSUE_FIELDS = [
    "key", "summary", "description", "issuetype", "status", "priority",
    "resolution", "assignee", "reporter", "created", "updated",
    "resolutiondate", "duedate", "labels", "components", "fixVersions",
    "versions", "project", "parent",
]

# Jira descriptions run to thousands of characters and Excel caps a cell at
# 32767, so truncate by default; --full-description keeps everything.
DESCRIPTION_LIMIT = 1000
EXCEL_CELL_LIMIT = 32000

AGE_BUCKETS = [(0, 7, "0-7 days"), (8, 30, "8-30 days"), (31, 90, "31-90 days"),
               (91, 180, "91-180 days"), (181, 10**6, "180+ days")]

PRIORITY_FILLS = {
    "blocker": "FFC7CE", "highest": "FFC7CE", "critical": "FFC7CE",
    "high": "FFD9A0", "major": "FFD9A0",
    "medium": "FFEB9C", "normal": "FFEB9C",
    "low": "C6EFCE", "minor": "C6EFCE", "lowest": "C6EFCE", "trivial": "C6EFCE",
}

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F3864")
SECTION_FONT = Font(bold=True, size=11, color="1F3864")
THIN_BORDER = Border(*[Side(style="thin", color="D9D9D9")] * 4)


# --------------------------------------------------------------------------
# Jira client
# --------------------------------------------------------------------------

class JiraError(RuntimeError):
    pass


class Jira:
    def __init__(self, base_url: str, token: str, verify: bool | str = True,
                 auth_mode: str = "bearer", username: str | None = None):
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update({
            "Accept": "application/json",
            "Content-Type": "application/json",
        })
        if auth_mode == "bearer":
            self.session.headers["Authorization"] = f"Bearer {token}"
        elif auth_mode == "basic":
            # Jira Cloud wants email:api-token; some Server setups accept
            # username:password or token-as-password.
            self.session.auth = (username or "", token)
        else:
            raise JiraError(f"Unknown auth mode: {auth_mode}")
        self.session.verify = verify

    def _request(self, method: str, path: str, **kwargs) -> requests.Response:
        url = f"{self.base_url}{path}"
        last_exc = None
        for attempt in range(MAX_RETRIES):
            try:
                resp = self.session.request(method, url, timeout=REQUEST_TIMEOUT, **kwargs)
            except requests.RequestException as exc:
                last_exc = exc
                time.sleep(2 ** attempt)
                continue
            if resp.status_code in (429, 500, 502, 503, 504) and attempt < MAX_RETRIES - 1:
                wait = int(resp.headers.get("Retry-After", 2 ** attempt))
                time.sleep(min(wait, 30))
                continue
            return resp
        raise JiraError(f"{method} {url} failed after {MAX_RETRIES} attempts: {last_exc}")

    def check_auth(self) -> str:
        resp = self._request("GET", "/rest/api/2/myself")
        if resp.status_code == 401:
            # Jira Server reports *why* it refused in these headers. Worth
            # surfacing: a CAPTCHA lockout needs a browser login, not a new token.
            reason = resp.headers.get("X-Seraph-LoginReason", "")
            denied = resp.headers.get("X-Authentication-Denied-Reason", "")
            hint = ""
            if reason:
                hint += f"\n  Jira says: X-Seraph-LoginReason = {reason}"
            if denied:
                hint += f"\n  Denied reason: {denied}"
            if "CAPTCHA" in (reason + denied).upper():
                hint += ("\n  >> Your account is CAPTCHA-locked after failed logins."
                         "\n     Log into Jira in a browser, solve the CAPTCHA, then retry.")
            raise JiraError(
                "401 Unauthorized - the credential was rejected. Either:\n"
                "  * the secret is wrong, expired or revoked - refresh it\n"
                "  * the email/username does not match the secret\n"
                "  * SSO is enforced, so a password will not work - create a token at\n"
                f"    {self.base_url}/secure/ViewProfile.jspa > Personal Access Tokens\n"
                "    and retry with --auth bearer"
                + hint +
                "\n  NOTE: repeated failures CAPTCHA-lock the account. If the next "
                "attempt\n  also fails, log into Jira in a browser before retrying.")
        if resp.status_code == 403:
            raise JiraError("403 Forbidden - the token is valid but lacks access, or a "
                            "proxy/CAPTCHA is blocking it. Try logging into Jira in a browser first.")
        if resp.status_code != 200:
            raise JiraError(f"Could not verify credentials ({resp.status_code}): {resp.text[:300]}")
        me = resp.json()
        return me.get("displayName") or me.get("name") or "unknown"

    def issue_type_names(self) -> set[str]:
        resp = self._request("GET", "/rest/api/2/issuetype")
        if resp.status_code != 200:
            return set()
        return {t["name"].lower() for t in resp.json()}

    def count(self, jql: str) -> int | None:
        """Return issue count for `jql`, or None if Jira rejects the query."""
        resp = self._request("POST", "/rest/api/2/search",
                             data=json.dumps({"jql": jql, "maxResults": 0,
                                              "fields": ["key"]}))
        if resp.status_code == 400:
            return None
        if resp.status_code != 200:
            raise JiraError(f"Count failed ({resp.status_code}) for `{jql}`: {resp.text[:300]}")
        return resp.json().get("total", 0)

    def search(self, jql: str) -> list[dict]:
        """Page through every issue matching `jql`."""
        issues: list[dict] = []
        start_at, total = 0, None
        while True:
            payload = {"jql": jql, "startAt": start_at, "maxResults": PAGE_SIZE,
                       "fields": ISSUE_FIELDS}
            resp = self._request("POST", "/rest/api/2/search", data=json.dumps(payload))
            if resp.status_code != 200:
                raise JiraError(f"Search failed ({resp.status_code}): {resp.text[:500]}")
            data = resp.json()
            batch = data.get("issues", [])
            issues.extend(batch)
            total = data.get("total", len(issues))
            start_at += len(batch)
            print(f"  fetched {len(issues)}/{total}", end="\r", flush=True)
            if not batch or start_at >= total:
                break
        print(f"  fetched {len(issues)}/{total}    ")
        return issues


# --------------------------------------------------------------------------
# JQL construction
# --------------------------------------------------------------------------

def jql_quote(value: str) -> str:
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


def discover_tag_fields(jira: Jira, tags: list[str],
                        fields: list[str]) -> dict[str, list[str]]:
    """Probe each (tag, field) pair and keep the ones Jira accepts with >0 hits.

    `fixVersion = "VS3"` is a hard 400 when no version of that name exists, so
    the probe is what keeps the real query from blowing up. A field that is
    valid but empty is also dropped - it only widens the query for nothing.
    """
    mapping: dict[str, list[str]] = {}
    print("Resolving how each tag is stored in Jira:")
    for tag in tags:
        matched = []
        for field in fields:
            total = jira.count(f"{field} = {jql_quote(tag)}")
            if total is None:
                status = "not a valid value for this field"
            elif total == 0:
                status = "valid but 0 issues"
            else:
                status = f"{total} issues"
                matched.append(field)
            print(f"  {tag:<9} {field:<16} {status}")
        mapping[tag] = matched
        if not matched:
            print(f"  !! {tag}: no field matched - it will be skipped")
    return mapping


def build_jql(mapping: dict[str, list[str]], issue_types: list[str],
              include_resolved: bool, extra: str | None) -> str:
    tag_clauses = []
    for tag, fields in mapping.items():
        if not fields:
            continue
        ors = " OR ".join(f"{f} = {jql_quote(tag)}" for f in fields)
        tag_clauses.append(f"({ors})")
    if not tag_clauses:
        raise JiraError("None of the tags matched any Jira field - nothing to report. "
                        "Re-run with --discover to see the probe results.")

    parts = ["(" + " OR ".join(tag_clauses) + ")"]
    if issue_types:
        parts.append("issuetype in (" + ", ".join(jql_quote(t) for t in issue_types) + ")")
    if not include_resolved:
        parts.append("statusCategory != Done")
    if extra:
        parts.append(f"({extra})")
    return " AND ".join(parts) + " ORDER BY priority DESC, created ASC"


# --------------------------------------------------------------------------
# Row shaping
# --------------------------------------------------------------------------

def parse_jira_dt(value: str | None) -> datetime | None:
    if not value:
        return None
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def _name(obj: dict | None, key: str = "name") -> str:
    return (obj or {}).get(key) or ""


def matched_tags(fields: dict, tags: list[str]) -> list[str]:
    """Which of the report tags this issue actually carries, in any field."""
    present = set(fields.get("labels") or [])
    present |= {c["name"] for c in fields.get("components") or []}
    present |= {v["name"] for v in fields.get("fixVersions") or []}
    present |= {v["name"] for v in fields.get("versions") or []}
    lowered = {p.lower() for p in present}
    return [t for t in tags if t.lower() in lowered]


COLUMNS = [
    "Key", "Summary", "Description", "Type", "Status", "Status Category", "Priority",
    "Matched Tags", "Project", "Assignee", "Reporter", "Created", "Updated",
    "Age (days)", "Days Since Update", "Due Date", "Labels", "Components",
    "Fix Versions", "Affects Versions", "Parent", "URL",
]


def clean_description(text: str | None, limit: int) -> str:
    """Flatten a Jira description for a spreadsheet cell.

    Collapses the blank lines and carriage returns that wiki markup leaves
    behind, then clips to `limit` so one verbose defect cannot blow past
    Excel's per-cell ceiling.
    """
    if not text:
        return ""
    lines = [ln.strip() for ln in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    flat = " ".join(ln for ln in lines if ln)
    cap = min(limit, EXCEL_CELL_LIMIT)
    return flat if len(flat) <= cap else flat[:cap - 3].rstrip() + "..."


def to_row(issue: dict, tags: list[str], base_url: str, now: datetime,
           desc_limit: int = DESCRIPTION_LIMIT) -> dict:
    f = issue.get("fields", {})
    created = parse_jira_dt(f.get("created"))
    updated = parse_jira_dt(f.get("updated"))
    due = parse_jira_dt(f.get("duedate"))
    parent = f.get("parent") or {}

    return {
        "Key": issue.get("key", ""),
        "Summary": f.get("summary") or "",
        "Description": clean_description(f.get("description"), desc_limit),
        "Type": _name(f.get("issuetype")),
        "Status": _name(f.get("status")),
        "Status Category": ((f.get("status") or {}).get("statusCategory") or {}).get("name", ""),
        "Priority": _name(f.get("priority")),
        "Matched Tags": ", ".join(matched_tags(f, tags)),
        "Project": _name(f.get("project")),
        "Assignee": _name(f.get("assignee"), "displayName") or "Unassigned",
        "Reporter": _name(f.get("reporter"), "displayName") or "",
        "Created": created.replace(tzinfo=None) if created else None,
        "Updated": updated.replace(tzinfo=None) if updated else None,
        "Age (days)": (now - created).days if created else None,
        "Days Since Update": (now - updated).days if updated else None,
        "Due Date": due.replace(tzinfo=None) if due else None,
        "Labels": ", ".join(f.get("labels") or []),
        "Components": ", ".join(c["name"] for c in f.get("components") or []),
        "Fix Versions": ", ".join(v["name"] for v in f.get("fixVersions") or []),
        "Affects Versions": ", ".join(v["name"] for v in f.get("versions") or []),
        "Parent": parent.get("key", ""),
        "URL": f"{base_url}/browse/{issue.get('key', '')}",
    }


# --------------------------------------------------------------------------
# Excel output
# --------------------------------------------------------------------------

def _style_header(ws, ncols: int, row: int = 1) -> None:
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", horizontal="left")
    ws.row_dimensions[row].height = 22


def _autosize(ws, rows: list[list], headers: list[str], cap: int = 55) -> None:
    for idx, header in enumerate(headers, start=1):
        widest = len(str(header))
        for row in rows:
            val = row[idx - 1] if idx - 1 < len(row) else ""
            widest = max(widest, len(str(val)) if val is not None else 0)
        ws.column_dimensions[get_column_letter(idx)].width = min(max(widest + 2, 10), cap)


def write_detail_sheet(ws, rows: list[dict], base_url: str,
                       table_name: str = "Defects") -> None:
    ws.append(COLUMNS)
    _style_header(ws, len(COLUMNS))

    for r in rows:
        ws.append([r[c] for c in COLUMNS])

    key_col = COLUMNS.index("Key") + 1
    url_col = COLUMNS.index("URL") + 1
    prio_col = COLUMNS.index("Priority") + 1
    age_col = COLUMNS.index("Age (days)") + 1

    for excel_row, r in enumerate(rows, start=2):
        key_cell = ws.cell(row=excel_row, column=key_col)
        key_cell.hyperlink = r["URL"]
        key_cell.font = Font(color="0563C1", underline="single")

        fill = PRIORITY_FILLS.get((r["Priority"] or "").lower())
        if fill:
            ws.cell(row=excel_row, column=prio_col).fill = PatternFill("solid", fgColor=fill)

        age = r["Age (days)"]
        if age is not None and age > 180:
            ws.cell(row=excel_row, column=age_col).font = Font(color="C00000", bold=True)

        for col in (COLUMNS.index("Created") + 1, COLUMNS.index("Updated") + 1,
                    COLUMNS.index("Due Date") + 1):
            ws.cell(row=excel_row, column=col).number_format = "yyyy-mm-dd"

        ws.cell(row=excel_row, column=COLUMNS.index("Summary") + 1).alignment = \
            Alignment(vertical="top", wrap_text=False)
        ws.cell(row=excel_row, column=url_col).font = Font(color="808080", size=9)

    if rows:
        ref = f"A1:{get_column_letter(len(COLUMNS))}{len(rows) + 1}"
        table = Table(displayName=table_name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(table)

    ws.freeze_panes = "C2"
    _autosize(ws, [[r[c] for c in COLUMNS] for r in rows], COLUMNS)
    ws.column_dimensions[get_column_letter(COLUMNS.index("Summary") + 1)].width = 70
    # Autosize would blow this out to the cap; a fixed width keeps the sheet
    # navigable and the full text is still in the cell.
    ws.column_dimensions[get_column_letter(COLUMNS.index("Description") + 1)].width = 80


def _tally(rows: list[dict], key: str) -> list[tuple[str, int]]:
    counts: dict[str, int] = {}
    for r in rows:
        counts[r[key] or "(none)"] = counts.get(r[key] or "(none)", 0) + 1
    return sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))


def _age_tally(rows: list[dict]) -> list[tuple[str, int]]:
    counts = {label: 0 for _, _, label in AGE_BUCKETS}
    for r in rows:
        age = r["Age (days)"]
        if age is None:
            continue
        for lo, hi, label in AGE_BUCKETS:
            if lo <= age <= hi:
                counts[label] += 1
                break
    return [(label, counts[label]) for _, _, label in AGE_BUCKETS]


def write_summary_sheet(ws, rows: list[dict], tags: list[str], jql: str,
                        generated: datetime) -> None:
    ws["A1"] = "Open Defect Report"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Tags: {', '.join(tags)}   (issue matches ANY tag)"
    ws["A3"] = f"Generated: {generated:%Y-%m-%d %H:%M} UTC"
    ws["A4"] = f"Total open defects: {len(rows)}"
    ws["A4"].font = Font(bold=True, size=12)
    ws["A5"] = "JQL:"
    ws["A6"] = jql
    ws["A6"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A6:F6")
    ws.row_dimensions[6].height = 46

    # Per-tag counts are computed from the issues themselves, so an issue
    # carrying two tags is counted under both - these will not sum to the total.
    tag_counts = [(t, sum(1 for r in rows if t in r["Matched Tags"].split(", ")))
                  for t in tags]

    sections = [
        ("Defects by tag", tag_counts),
        ("Defects by status", _tally(rows, "Status")),
        ("Defects by priority", _tally(rows, "Priority")),
        ("Defects by project", _tally(rows, "Project")),
        ("Defects by age", _age_tally(rows)),
        ("Defects by assignee (top 20)", _tally(rows, "Assignee")[:20]),
    ]

    row = 8
    for title, data in sections:
        ws.cell(row=row, column=1, value=title).font = SECTION_FONT
        row += 1
        for header_idx, header in enumerate(("Value", "Count", "% of total"), start=1):
            cell = ws.cell(row=row, column=header_idx, value=header)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
        row += 1
        for label, count in data:
            ws.cell(row=row, column=1, value=label).border = THIN_BORDER
            ws.cell(row=row, column=2, value=count).border = THIN_BORDER
            pct = ws.cell(row=row, column=3,
                          value=(count / len(rows) if rows else 0))
            pct.number_format = "0.0%"
            pct.border = THIN_BORDER
            row += 1
        row += 2

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12


def write_workbook(path: Path, rows: list[dict], tags: list[str], jql: str,
                   base_url: str, generated: datetime) -> None:
    wb = Workbook()
    write_summary_sheet(wb.active, rows, tags, jql, generated)
    wb.active.title = "Summary"
    write_detail_sheet(wb.create_sheet("Open Defects"), rows, base_url, "AllDefects")

    for idx, tag in enumerate(tags):
        subset = [r for r in rows if tag in r["Matched Tags"].split(", ")]
        if not subset:
            continue
        # Sheet names cannot contain [ ] : * ? / \ and cap at 31 chars; Excel
        # table names must additionally be unique and identifier-safe.
        sheet = "".join("-" if ch in r"[]:*?/\\" else ch for ch in tag)[:31]
        safe_name = "".join(ch if ch.isalnum() else "_" for ch in tag)
        write_detail_sheet(wb.create_sheet(sheet), subset, base_url,
                           f"Tag{idx}_{safe_name}")

    wb.save(path)


def write_csv(path: Path, rows: list[dict]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=COLUMNS)
        writer.writeheader()
        for r in rows:
            writer.writerow({
                k: (v.strftime("%Y-%m-%d") if isinstance(v, datetime) else v)
                for k, v in r.items()
            })


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------

def get_token(cli_token: str | None, auth_mode: str = "basic",
              user: str | None = None) -> str:
    """Resolve the secret. Basic auth pairs it with an email; bearer uses a PAT.

    An email address is not itself a credential, so a secret is required either
    way - it is just a password / API token under basic rather than a PAT.
    """
    token = cli_token
    if not token:
        for var in ("JIRA_PAT", "JIRA_TOKEN", "JIRA_PASSWORD", "MY_JIRA_PAT"):
            if os.environ.get(var):
                token = os.environ[var]
                print(f"Using secret from ${var}")
                break
    if not token:
        label = (f"Jira API token or password for {user}: " if auth_mode == "basic"
                 else "Jira Personal Access Token: ")
        token = getpass.getpass(label)
    if not token.strip():
        raise JiraError(
            "No secret supplied. An email address alone cannot authenticate to "
            "Jira - set JIRA_PAT (or JIRA_PASSWORD) to the API token or password "
            f"for {user or 'your account'}, or pass --token.")
    # A token pasted with a trailing newline or stray quotes fails as a silent
    # 401, which is indistinguishable from an expired token - strip both.
    return token.strip().strip('"').strip("'")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--url", default=BASE_URL, help=f"Jira base URL (default {BASE_URL})")
    parser.add_argument("--token", help="PAT; defaults to $JIRA_PAT / $JIRA_TOKEN "
                                        "/ $MY_JIRA_PAT, else prompts")
    parser.add_argument("--auth", choices=["bearer", "basic"], default=DEFAULT_AUTH,
                        help=f"basic = email + token/password (default: {DEFAULT_AUTH}); "
                             "bearer = Server/DC Personal Access Token")
    parser.add_argument("--user", default=DEFAULT_EMAIL,
                        help=f"Email/username for --auth basic (default: {DEFAULT_EMAIL})")
    parser.add_argument("--tags", nargs="+", default=TAGS, help="Tag values to report on")
    parser.add_argument("--fields", nargs="+", default=CANDIDATE_FIELDS,
                        help="JQL fields to probe for each tag")
    parser.add_argument("--issue-types", nargs="+", default=DEFECT_ISSUE_TYPES,
                        help="Issue types treated as defects")
    parser.add_argument("--jql", help="Use this JQL verbatim and skip discovery")
    parser.add_argument("--extra-jql", help="Extra JQL ANDed onto the generated query")
    parser.add_argument("--include-resolved", action="store_true",
                        help="Include Done issues (default: open only)")
    parser.add_argument("--discover", action="store_true",
                        help="Only print the tag-to-field mapping, then exit")
    parser.add_argument("--description-limit", type=int, default=DESCRIPTION_LIMIT,
                        help=f"Characters of Description to keep (default {DESCRIPTION_LIMIT})")
    parser.add_argument("--full-description", action="store_true",
                        help="Keep the whole Description (capped at Excel's cell limit)")
    parser.add_argument("--out", help="Output .xlsx path")
    parser.add_argument("--csv", action="store_true", help="Also write a .csv")
    parser.add_argument("--insecure", action="store_true",
                        help="Skip TLS verification (corporate MITM proxies)")
    parser.add_argument("--ca-bundle", help="Path to a corporate CA bundle .pem")
    args = parser.parse_args(argv)

    verify: bool | str = True
    if args.insecure:
        verify = False
        requests.packages.urllib3.disable_warnings()  # type: ignore[attr-defined]
    elif args.ca_bundle:
        verify = args.ca_bundle

    try:
        secret = get_token(args.token, args.auth, args.user)
        jira = Jira(args.url, secret, verify=verify,
                    auth_mode=args.auth, username=args.user)
        who = f"{args.auth} as {args.user}" if args.auth == "basic" else args.auth
        print(f"Connecting to {args.url} (auth: {who}) ...")
        print(f"Authenticated as {jira.check_auth()}\n")

        if args.jql:
            jql = args.jql
            print(f"Using supplied JQL:\n  {jql}\n")
        else:
            mapping = discover_tag_fields(jira, args.tags, args.fields)
            if args.discover:
                print("\nMapping:")
                for tag, fields in mapping.items():
                    print(f"  {tag:<9} -> {', '.join(fields) or '(no match)'}")
                return 0

            available = jira.issue_type_names()
            types = [t for t in args.issue_types if not available or t.lower() in available]
            dropped = [t for t in args.issue_types if t not in types]
            if dropped:
                print(f"\nNote: issue type(s) {', '.join(dropped)} do not exist here - ignored.")

            jql = build_jql(mapping, types, args.include_resolved, args.extra_jql)
            print(f"\nJQL:\n  {jql}\n")

        print("Fetching issues ...")
        issues = jira.search(jql)
    except JiraError as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        return 1
    except requests.exceptions.SSLError as exc:
        print(f"\nERROR: TLS verification failed ({exc}).\n"
              "Behind a corporate proxy, retry with --ca-bundle <corp-ca.pem> "
              "or --insecure.", file=sys.stderr)
        return 1

    now = datetime.now(timezone.utc)
    desc_limit = EXCEL_CELL_LIMIT if args.full_description else args.description_limit
    rows = [to_row(i, args.tags, args.url, now, desc_limit) for i in issues]

    stamp = now.strftime("%Y%m%d-%H%M")
    out = Path(args.out) if args.out else Path(f"open-defects-{stamp}.xlsx")
    out.parent.mkdir(parents=True, exist_ok=True)
    write_workbook(out, rows, args.tags, jql, args.url, now)
    print(f"\nWrote {len(rows)} defects to {out.resolve()}")

    if args.csv:
        csv_path = out.with_suffix(".csv")
        write_csv(csv_path, rows)
        print(f"Wrote CSV to {csv_path.resolve()}")

    if rows:
        print("\nBy tag:")
        for tag in args.tags:
            print(f"  {tag:<9} {sum(1 for r in rows if tag in r['Matched Tags'].split(', '))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
