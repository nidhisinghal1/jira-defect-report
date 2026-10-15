# Open Defect Report — VS3 / MCIS / R4.5 / R4.4.1

Pulls every **open defect** from https://jira.viasat.com tagged with any of
`VS3`, `MCIS`, `R4.5`, `R4.4.1` and writes a formatted Excel report.

## Setup

Dependencies are already installed (`requests`, `openpyxl`). If you move this
to another machine:

```bash
pip install requests openpyxl
```

Create a Personal Access Token in Jira: **Profile → Personal Access Tokens →
Create token**, then:

```bash
# Git Bash
export JIRA_PAT=<token>

# PowerShell
$env:JIRA_PAT = "<token>"
```

The script defaults to **basic auth as `nidhi.singhal@viasat.com`**. Note that
an email address is not itself a credential — Jira always needs a secret
alongside it. The script reads `JIRA_PAT`, `JIRA_TOKEN`, `JIRA_PASSWORD`, then
`MY_JIRA_PAT`; if none is set it prompts (input hidden).

Change the account without editing the file via `--user`, or `$JIRA_EMAIL`.

If your Jira enforces SSO, a password will not authenticate over the API. Use
a Personal Access Token instead:

```bash
JIRA_PAT=<token> python jira_defect_report.py --auth bearer
```

Set `$JIRA_AUTH_MODE=bearer` to make that the default.

> **Account lockout:** Jira Server CAPTCHA-locks an account after a handful of
> failed logins. If you get two 401s in a row, stop and log into Jira in a
> browser before retrying — otherwise you lose browser access too. The script
> reports Jira's `X-Seraph-LoginReason` on 401 so you can tell the cases apart:
> `AUTHENTICATED_FAILED` is a wrong secret, `AUTHENTICATION_DENIED` is a lock.

## Run

```bash
python jira_defect_report.py                      # -> open-defects-<timestamp>.xlsx
python jira_defect_report.py --out weekly.xlsx --csv
python jira_defect_report.py --discover           # show the tag->field mapping only
```

## How the four tags are resolved

They are stored inconsistently in Jira, so the script does not guess. On each
run it probes every tag against `labels`, `fixVersion`, `component` and
`affectedVersion`, and builds its JQL from only the combinations Jira accepts
with a non-zero hit count. This matters: `fixVersion = "VS3"` is a hard
**400 error** when no version of that name exists, which would otherwise kill
the whole query.

`--discover` prints the probe table so you can confirm the mapping:

```
VS3       labels           412 issues
VS3       fixVersion       not a valid value for this field
R4.5      labels           valid but 0 issues
R4.5      fixVersion       228 issues
```

Once you know the real mapping you can pin it and skip discovery:

```bash
python jira_defect_report.py --jql '(labels = "VS3" OR labels = "MCIS" OR fixVersion in ("R4.5","R4.4.1")) AND issuetype = Bug AND statusCategory != Done'
```

## What "open" means

`statusCategory != Done` — everything not in a green/done status, regardless of
how the workflows differ between projects. Use `--include-resolved` to drop
that filter.

Issue types default to `Bug` and `Defect`; the script checks which of those
actually exist on this Jira and silently ignores the rest, so it works whether
your projects call them Bugs or Defects. Override with `--issue-types`.

## Output

**Summary** sheet — total, plus breakdowns by tag, status, priority, project,
age bucket and top-20 assignee, with the exact JQL used recorded for audit.

**Open Defects** sheet — one row per defect: key (hyperlinked to Jira),
summary, type, status, priority, matched tags, project, assignee, reporter,
created/updated, age in days, days since update, due date, labels, components,
fix/affects versions, parent, URL. Filterable table, frozen panes, priority
colour-coded, and ages over 180 days flagged red.

**One sheet per tag** — `VS3`, `MCIS`, `R4.5`, `R4.4.1`, same columns.

Per-tag counts intentionally sum to more than the total: an issue carrying
both `R4.5` and `MCIS` is counted under both.

## Options

| Flag | Purpose |
|---|---|
| `--url` | Jira base URL (default `https://jira.viasat.com`, or `$JIRA_URL`) |
| `--token` | Secret inline; defaults to `$JIRA_PAT` / `$JIRA_TOKEN` / `$JIRA_PASSWORD` / `$MY_JIRA_PAT`, else prompts |
| `--auth` | `basic` (email + token/password, default) or `bearer` (Server PAT) |
| `--user` | Email/username for basic auth (default `nidhi.singhal@viasat.com`, or `$JIRA_EMAIL`) |
| `--tags` | Override the four tag values |
| `--fields` | Which JQL fields to probe |
| `--issue-types` | Issue types treated as defects |
| `--jql` | Use this JQL verbatim, skip discovery |
| `--extra-jql` | Extra clause ANDed on, e.g. `project = ABC` |
| `--include-resolved` | Include Done issues |
| `--discover` | Print the tag→field mapping and exit |
| `--out` / `--csv` | Output path / also emit CSV |
| `--ca-bundle` / `--insecure` | Corporate TLS interception |

## Troubleshooting

- **401** — token rejected or expired; generate a new one.
- **403** — token valid but blocked; log into Jira in a browser once (Jira
  Server throws a CAPTCHA after failed logins) and retry.
- **SSLError** — corporate TLS proxy. Use `--ca-bundle <corp-ca.pem>`, or
  `--insecure` as a last resort.
- **"None of the tags matched any Jira field"** — run `--discover`; the tag
  names likely differ from what is in Jira (e.g. `R4.5.0` rather than `R4.5`).

Paging, retries on 429/5xx with backoff, and unassigned/empty fields are all
handled.
