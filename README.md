# Leasing automations — Prudent Growth

Home for leasing-department automations owned by Alexis Pattison (CSO).

## `snap_shot/` — Leasing Snap Shot rebuild

Rebuilds Ann's 72-property Leasing Snap Shot workbook nightly, weekly, and on demand.

- Nightly (`.github/workflows/snap-shot-nightly.yml`) — full rebuild at 4:00 AM ET
- Weekly (`.github/workflows/snap-shot-weekly.yml`) — nightly + Broker Active Interest refresh, Mondays at 5:00 AM ET
- Poll (`.github/workflows/snap-shot-poll.yml`) — every 15 min 7 AM–6 PM ET Mon–Fri, rebuilds if the "Rebuild Snap Shot" checkbox is checked on the ClickUp control task

See `snap_shot/rebuild.py --help` and the `prudent-snap-shot-rebuild` skill (personal skill library) for full docs.

## Required repo secrets

| Secret | Purpose |
|---|---|
| `APATTISON_MS_REFRESH_TOKEN` | Microsoft Graph delegated refresh token for `apattison@prudentgrowth.com` — used for SharePoint round-trip + error emails |
| `GRAPH_CLIENT_ID` | Azure app registration client ID |
| `GRAPH_CLIENT_SECRET` | Azure app registration client secret |
| `GRAPH_TENANT_ID` | Azure tenant ID |
| `APPFOLIO_API_SECRET` | AppFolio Read-Only API password |
| `CLICKUP_API_TOKEN` | ClickUp personal API token |

## Change management

All changes follow the `automation-change-management` org skill. Every change updates `CHANGELOG.md` at repo root with a full WHY / WHAT CHANGED / WHAT DID NOT CHANGE / RISK / VERIFICATION section.
