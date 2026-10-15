# Security policy

Qent processes internal firewall, VPN and endpoint logs, so its own security matters.

## Reporting a vulnerability
Please **don't open a public issue**. Report it privately through GitHub:
**Security → Report a vulnerability** on this repository.

## Built-in controls
- Usernames and every unrecognised log column are dropped on upload; internal IPs and host names become keyed pseudonyms; command lines and paths are redacted.
- Logs never reach an AI model. Only public feed items are summarised.
- Everything that leaves (STIX, CSV, TAXII, MISP push) passes `share_safe()`: no pseudonyms, no private IPs.
- Writing back to MISP is off unless `MISP_PUSH=1`. Log rows are deleted after 30 days.
- Local mode binds to localhost with Host checking. Cloud mode requires sign-in (Entra) and roles Viewer / Analyst / Admin.
- Secrets come from environment variables or a git-ignored `.env`; none are in this repository.
