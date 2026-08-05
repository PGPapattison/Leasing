# CHANGELOG

All notable changes to the Leasing automations repo. Newest at the top.

## 2026-08-05 — First successful end-to-end run + add Pillow for logo (L1)

- **File(s):** `snap_shot/requirements.txt`
- **Author:** Alexis Pattison

**Why.** After yesterday's auth debugging (see follow-up entry below), the workflow now runs green end-to-end on Jerry's shared `PrudentGrowth Email Automation` app registration (client `d4aec1ec-...`, tenant `b2a05ba0-...`). Run 31010581786 succeeded in 10 seconds: pulled 1060 AppFolio units into 84 properties, wrote 72 broker-managed properties with 905 unit rows, and uploaded `Leasing-Snap-Shot.xlsx` to SharePoint. Only cosmetic issue: `Logo image skipped: You must install Pillow to fetch image objects` — the workbook shipped without the PG logo because Pillow wasn't listed in requirements.

**What changed.** Added `Pillow>=10.0` to `snap_shot/requirements.txt`.

**What did not change.** No code changes. Auth stack unchanged. Workbook logic unchanged.

**Risk / rollback.** Risk: none. Adding an image library. Rollback: revert this commit; workbook will lose the logo again.

**Verification.** Re-fire nightly. Expect the same 10‑second run, but with no `Pillow` warning in the log and the PG logo visible on the workbook cover sheet in SharePoint.

## 2026-08-05 — Reset Graph auth to Jerry's shared app registration (L2)

- **File(s):** `.github/workflows/snap-shot-*.yml`, GitHub secrets (deleted GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_TENANT_ID)
- **Author:** Alexis Pattison

**Why.** Yesterday I stood up a brand-new Azure app registration to power Snap Shot Graph auth. It never worked — kept returning `AADSTS65001: not consented` despite both user and admin consent flows. After Alexis's GitHub connected to the `PrudentGrowth` org, I could read the working `PrudentGrowth/delinquency` scripts and found the pattern: every existing PG automation uses the same shared app (`PrudentGrowth Email Automation`, client `d4aec1ec-...`, tenant `b2a05ba0-...`) with a public‑client refresh‑token flow (no client secret). Our `rebuild.py` was already coded for exactly that pattern — the mismatch was that yesterday's GitHub secrets pointed at Alexis's custom app instead.

**What changed.**
- Removed `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET`, `GRAPH_TENANT_ID` env passthroughs from all three workflows (they were never read anyway).
- Deleted those three secrets from the GitHub Secrets store.
- Regenerated `APATTISON_MS_REFRESH_TOKEN` against the shared PG app via a new helper (`get_refresh_token_v2.py`, not committed — personal helper).
- Extended app scopes with `Files.ReadWrite.All` and `Sites.ReadWrite.All` (Jerry's app previously only had `Mail.Send`); admin‑consented at the tenant level.

**What did not change.** No `rebuild.py` changes. Auth code was already correct.

**Risk / rollback.** Risk: low. Rollback: re-add the three deleted secrets and re-add the env lines to the workflows if Jerry ever revokes tenant‑wide access on the shared app.

**Verification.** Manual run 31010581786 completed successfully in 10 seconds with SharePoint upload confirmed.

## 2026-08-04 — Fix f-string SyntaxError on Python 3.11, part 2 (L1)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison

**Why.** Second live nightly run (30950269822) failed at import with the same `SyntaxError: f-string expression part cannot include a backslash` — different site, line 1400. The first fix used an AST walk to hunt for issues, but `ast.parse()` runs the Python 3.14 tokenizer which normalizes string escapes before the walk sees them, hiding backslashes inside string-literal `.join()` arguments. Result: the AST audit missed line 1400.

**What changed.**
- Line 1400: same middle-dot hoisting treatment as line 1361 (extract `\u00b7` into a variable outside the f-string).
- Replaced the AST audit approach with a raw-source regex scan and `ast.parse(feature_version=(3,11))` — both now report clean.

**What did not change.** Rendered string still byte-identical. No workbook output changes.

**Risk / rollback.** Risk: low, syntactic-only fix. Rollback: revert the commit.

**Verification.** Re-fire nightly workflow. Should progress past import. If it fails at a later runtime step (Graph auth, ClickUp query, SharePoint write) that becomes the next thing to debug.

## 2026-08-04 — Fix f-string SyntaxError on Python 3.11 (L1)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison

**Why.** First live GitHub Actions run (nightly, workflow_dispatch, run ID 30950126934) failed at import with `SyntaxError: f-string expression part cannot include a backslash` on line 1361. Python 3.11 (Actions runner) is stricter than Python 3.14 (workspace) about backslashes inside f-string expressions, so the dry-run passed locally but the deployed run couldn't even parse the file.

**What changed.** Hoisted the `\u00b7` (middle dot) unicode escape out of the f-string join expression into intermediate variables (`_mid_dot`, `_tag_join`). No behavior change — the rendered string is byte-identical.

**What did not change.** No workbook output changes. No other f-strings altered. Audited the full file with an AST walk to confirm no other f-string expressions contain backslashes. (This audit later proved insufficient — see the follow-up entry above.)

**Risk / rollback.** Risk: low, syntactic-only fix. Rollback: revert the commit.

**Verification.** Re-fire the nightly workflow manually via `gh workflow run`. Expect it to progress past the import step and either succeed end-to-end or fail at a later runtime step (which becomes the next thing to debug).

## 2026-08-04 — Snap Shot rebuild — initial deployment (L4)

- **Commit:** initial commit (see git log)
- **File(s):** `snap_shot/rebuild.py`, `snap_shot/requirements.txt`, `.github/workflows/snap-shot-nightly.yml`, `.github/workflows/snap-shot-weekly.yml`, `.github/workflows/snap-shot-poll.yml`, `README.md`
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` v1.0 (Alexis's personal skill library, skill_id `3fbd202a-03c3-4b3b-a379-3e93c29b661b`)

**Why.** Ann's Leasing Snap Shot workbook was a manual weekly rebuild. She wanted an always-fresh version she could edit directly, without losing her hand-authored notes (Deal Activity, Property Flags, Vacant Callouts, Ann's Commentary, Broker Calls, Market Rent overrides, per-unit Notes) on each refresh. This automation rebuilds nightly from AppFolio + ClickUp Broker Reporting data, round-tripping through SharePoint to preserve Ann's edits.

**What changed.** New repo (previously empty). Added:

- `snap_shot/rebuild.py` — 1,781-line standalone Python script; four modes (`nightly`, `weekly`, `poll`, `dry-run`); pulls AppFolio `rent_roll.json`, ClickUp Broker Reporting list 901114227189, downloads and re-uploads the SharePoint workbook at `sites/DBMigration/Shared Documents/General/Brain Snap Shot/Leasing-Snap-Shot.xlsx`; race-skip if file modified within last 15 minutes by a non-automation user.
- `snap_shot/requirements.txt` — `openpyxl`, `requests`, `msal`.
- Three GitHub Actions workflows: nightly (4 AM ET daily), weekly (5 AM ET Mondays, refreshes Broker Active Interest from Broker Beat attachments), poll (every 15 min 7 AM–6 PM ET Mon–Fri, rebuilds if ClickUp checkbox is checked).
- `README.md` documenting the repo and required secrets.

**What did not change.**

- Workbook layout, colors, fonts, KPI banner, and lowest-occupancy-states section are byte-identical to the prototype Ann approved. Dry-run confirmed pixel-identical rendering.
- No changes to any other Prudent Growth automation or ClickUp workflow.
- No AppFolio writes — read-only.
- No ClickUp field changes yet — the "Rebuild Snap Shot" checkbox custom field on list 901114227189 still needs to be added manually (or the poll workflow will exit 0 with a "checkbox field not found" log line each run — safe, not destructive).

**Risk / rollback.**

- Risk: **medium** — new automation with three schedules. Failure modes: SharePoint upload fails (script logs + emails Alexis); AppFolio API 429 (script sleeps + retries once); ClickUp API down (weekly mode's Broker Active Interest refresh degrades to prior week's values).
- Rollback: disable the three workflows (`gh workflow disable snap-shot-nightly.yml` etc.) or delete the `.github/workflows/*.yml` files. Rebuild.py can also be invoked manually with `--mode=dry-run` for local testing without any external writes.

**Verification.**

- Dry-run tested in-session on 2026-08-04: `python3 snap_shot/rebuild.py --mode=dry-run` completes with a workbook at `/tmp/Leasing-Snap-Shot-dryrun.xlsx`; comparing that against the prior prototype via PDF export shows pixel-identical layout, colors, KPI banner, and per-property blocks. Only numeric drift is expected (live AppFolio data between two consecutive pulls).
- First live run: nightly cron fires 2026-08-05 08:00 UTC (4 AM ET). Alexis will spot-check the SharePoint workbook that morning. If Ann's edits from the previous evening are preserved after that first run, the round-trip is verified.
- Monitor first week: check GitHub Actions run history at https://github.com/PGPapattison/Leasing/actions each morning through 2026-08-11.
