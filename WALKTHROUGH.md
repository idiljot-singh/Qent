# Qent Walkthrough

How Qent turns "too many threat reports" into "this PIR is supported by two independent sources", and how your own logs add evidence nobody else has. Everything below uses the bundled general demo PIRs and the synthetic logs in `samples/`.

## 1. The problem

A security team gets hundreds of advisories, vendor blogs and feeds a week. Without written priority intelligence requirements (PIRs) nobody can say which reports matter, whether a question has really been answered, or whether collection is even working. Qent starts from the questions and works backwards.

## 2. Glossary

| Term | Meaning |
|---|---|
| **PIR** | Priority intelligence requirement: a question the business needs answered, with the decision it supports |
| **Source reliability** | A to F grade of a feed (CISA KEV is A, a vendor blog is C) |
| **Item** | One feed entry that matched a PIR's keywords |
| **Indicator** | CVE, IP, hash, URL, domain, ATT&CK technique or actor / malware name extracted from an item |
| **Status** | OPEN, PARTIAL, SUPPORTED or ANSWERED (section 5) |
| **Finding** | A deterministic detection in your own logs (section 6) |
| **Sighting** | An indicator seen in your logs that is also known to a feed or your MISP |
| **STIX / TAXII** | Standard formats to share intelligence (bundle file / read-only server) |

## 3. The flow at a glance

```
PIRs  ->  Collect  ->  Match  ->  Extract  ->  Assess  ->  Answer
(questions) (feeds)   (keywords)  (indicators)  (status)   (analyst)

Your logs -> Findings -> Sightings (vs feeds + MISP) -> Share (STIX / CSV / TAXII / MISP)
```

## 4. Direction and collection

**PIRs** are loaded on first start from `pir_workbook.xlsx` (or created in the Requirements tab). Each has a question, the decision it supports, a level (tactical / strategic), match keywords and a staleness limit in days.

| ID | Question (short) |
|---|---|
| PIR-01 | Phishing, ClickFix and account takeover against cloud and email accounts |
| PIR-02 | Exploited vulnerabilities in widely deployed edge and enterprise products |
| PIR-03 | Ransomware and extortion groups and their entry methods |
| PIR-04 | State-aligned actors targeting our sector and supply chain |
| PIR-05 | Staff actions that put company data at risk |

**Sources** (Sources tab) are RSS feeds, CISA KEV, abuse.ch Feodo and MISP. Each is *tasked* to the PIRs it can help with and carries a reliability grade. **Collect now** (or `python qent.py collect` on a schedule) fetches them in parallel with size caps and a public-address check, so a feed cannot point Qent at your internal network.

## 5. Matching, assessment and status

1. **Match.** An item is kept for a PIR only if its source is tasked to that PIR, it is newer than 30 days, and one of the PIR's keywords hits at a word start (`entra` does not hit `central`).
2. **Extract.** Regular expressions pull CVEs, IPs, hashes, URLs, domains, ATT&CK techniques and actor / malware names. Defanged forms (`hxxp`, `[.]`) are understood.
3. **Score.** Each item scores keyword hits (max 3) + context match for your estate (`ORG_CONTEXT`, default Microsoft 365, VPN, supply chain, ...) + source reliability + analyst verdict.
4. **Status** per PIR:

| Status | Rule |
|---|---|
| OPEN | no matching item |
| PARTIAL | some evidence, not corroborated |
| SUPPORTED | 2 or more independent sources and a strong item (automatic) |
| ANSWERED | an analyst marked an item as answering the question |

5. **Confidence** is the number of independent sources: none, low, medium, high. One source never rates higher than "low".
6. **Gaps are shown, not hidden.** A PIR with no items, a feed that failed, a feed that returned nothing for it, or evidence older than its staleness limit raises a flag.

An optional AI summary (Azure OpenAI in an EU data zone for cloud, Gemini for development) writes a short narrative per PIR from **public feed items only**. Without a key, the rules above still run.

## 6. Your own logs (beta)

On the **Internal** tab, upload a firewall, SSL-VPN or endpoint CSV (`samples/`).

**On upload, before anything is stored:** usernames and every unrecognised column are dropped; internal IPs and host names become keyed pseudonyms (`asset-1a2b3c4d`); command lines and paths are redacted of secrets, accounts and addresses; rows are deleted after 30 days. The file is parsed in memory and never written to disk. **No AI ever sees logs.**

**Detections** (rule-based, mapped to MITRE ATT&CK):

| Finding | ATT&CK |
|---|---|
| Port scan / sweep | T1046 |
| Repeated failures, then a success | T1110 |
| Password spraying / credential stuffing | T1110.003 |
| Regular intervals (beaconing) | T1071 |
| Encoded PowerShell, certutil download | T1059.001, T1105 |
| Exploit probes: Fortinet, Citrix, Ivanti, Cisco ASA paths | T1190 |
| Scanner and automated clients (curl, sqlmap, python-requests) | T1595.002 |
| Same account from two countries within hours | T1078 |
| Rare process on one asset | - |

**Sightings.** Internal indicators are compared with the collected feeds and, if a MISP source has lookup on, searched in your MISP (`attributes/restSearch`).

## 7. Sharing

| Output | Notes |
|---|---|
| STIX 2.1 | Indicators, vulnerabilities, attack patterns, actors and one report per PIR. Stable ids, so MISP de-duplicates re-imports |
| Internal STIX / CSV | Sightings and ATT&CK techniques, TLP:AMBER, no internal identifiers |
| PDF report | Report with the PIR status, evidence and indicators |
| TAXII 2.1 | Read-only, off by default (`TAXII=1`) |
| MISP push | Events and sightings back to your MISP, off by default (`MISP_PUSH=1`) |

Everything that leaves passes `share_safe()`: no pseudonyms, no private IPs.

## 8. Roles, audit and health

- **Roles:** Viewer, Analyst, Admin. Local mode is Admin on localhost; cloud mode requires Entra sign-in.
- **Audit:** every mark, grade, upload, delete and export is recorded with who and when.
- **Health tab:** per feed (disabled, key missing, never run, failed, returned 0, none matched), per PIR (gap, stale) and for AI and the analyst loop. Each check says how to fix it.

## 9. Security of Qent itself

Host-header checks, Origin checks on every write, size caps on feeds and uploads, public-address-only fetching (private hosts only by explicit opt-in over https), self-signed MISP certificates pinned by file rather than disabling verification, secrets only from the environment, spreadsheet-formula neutralising in CSV exports. Details in [SECURITY.md](SECURITY.md).

## 10. Known limitations

- A proof of concept: one process, one user at a time in local mode.
- Detections are heuristics; regular intervals are also normal for monitoring and updates. Confirm before acting.
- Keyword matching misses synonyms; tune each PIR's keywords.
- MISP warninglists may be off on a lab instance, so common addresses (for example public DNS) can appear as hits.
- The MISP push and the Azure deploy have not been run against production.
- Upload cap: 20 MB per file, 200k rows per file, 1M rows in total.

## 11. Quick reference

```
python qent.py                       serve on http://localhost:8080
python qent.py --test                self-test (prints "selftest ok")
python qent.py collect               one collection run (for cron)
python qent.py import <workbook>     load PIRs / sources from a workbook
```
