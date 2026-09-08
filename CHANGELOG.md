# CHANGELOG

All notable changes to the Leasing automations repo. Newest at the top.

## 2026-09-08 — Snap Shot: add vacancy prospect subtask rollup to col O (v1.9)

- **Commit:** _(this commit)_
- **File(s):** `snap_shot/rebuild.py`, `snap_shot/tests/test_vacancy_prospect_rollup.py` (new)
- **Branch:** `col-o-vacancy-rollup` (NOT merged to main yet — review only)
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` v1.8 → v1.9

**Why.** Follow-through on the 2026-08-25 col P retirement note that flagged col O for a follow-up Thursday change. REMs now edit prospect updates directly in ClickUp Vacancy Pipeline subtasks, but Ann's Snap Shot workbook only showed the parent Vacancy task's Summary field — so every open prospect on a vacant unit was invisible unless she clicked into ClickUp. This surfaces every open subtask directly in col O so she can run the leasing call from the workbook alone.

**What changed.** New function `pull_vacancy_prospect_subtasks(tz)` in `snap_shot/rebuild.py` walks every Vacancy Pipeline task (list `901113575628`), indexes parents by (PropertyId, Unit#), groups subtasks by parent, filters out closed/executed statuses via the new `VACANCY_PROSPECT_HIDE_STATUSES` frozenset (both brief-supplied generic aliases and the actual live status names `lease/expansion executed` and `dead/lost deal/completed` — verified via `cu_get_list_statuses` on 2026-09-08), fetches the newest comment per subtask, and returns `{(pid, unit_norm): [row_dict, …]}`. `run_build()` now calls this pull. `build_workbook()` and `write_property_block()` accept a `vacancy_prospects_by_prop_unit` mapping. For vacant units only, `write_property_block` appends a `Prospects (N):` block to col O. Format: `• <Prospect Name> — <Status Title Case> (M/D: "first 80 chars of latest comment…")`, sorted most-recently-updated first; comment tail omitted entirely when a subtask has no comments. New helpers: `_extract_prospect_name`, `_extract_status_title`, `_format_comment_tail`, `_sort_prospect_subtasks`, `format_prospect_rollup_block`, `cu_get_task_comments`, `_rehydrate_prop_unit_dict`. New CLI flags: `--local-out=<path>` (write .xlsx locally instead of uploading to SharePoint) and `--data-snapshot=<path>` (load all fixtures from JSON and skip every live API call — both require `--mode=dry-run`). New test file `snap_shot/tests/test_vacancy_prospect_rollup.py` with 18 tests; full suite 29 pass.

**What did not change.** Occupied units — they still get their Renewal Pipeline Summary keyed on Tenant ID and no rollup is appended. Renewal Pipeline list (`901113575567`), Docs Workflow list (`901113991446`). Col N (Notes), col P (blank gap), col Q (Last Synced). AppFolio rent-roll pull. Nightly / weekly / afternoon-sweep workflows — no schedule changes, no new workflow file. SharePoint upload path. Ann's edit-preservation flow. Race-condition skip on Ann-in-workbook. No emails or notifications from this run.

**Risk / rollback.** Risk: low. The rollup only appends to col O; if `pull_vacancy_prospect_subtasks` raises the try/except in `run_build` swallows and logs the error, and the build proceeds with an empty rollup (col O still shows the parent's Summary). Rollback: `git revert <sha>` — the change is additive with no schema, workflow, or ClickUp field changes. No mutations to any ClickUp task or field.

**Verification.** `python3 -m py_compile snap_shot/rebuild.py` — succeeds. `pytest snap_shot/tests/` — 29 tests pass (11 pre-existing + 18 new). Local dry-run against a real-data snapshot pulled via the ClickUp connector on 2026-09-08 (17 open subtasks matched to 13 vacant units across 10 properties) produced `/home/user/workspace/snap_shot_v1.9_sample.xlsx` with 8 populated `Prospects (N):` blocks in col O, including two-prospect rollups on Hampton Cove Shops Unit 322 C, Snow Street Village Suite 663, and Somerset Shoppes Unit 7798. **Not merged to main.** Branch `col-o-vacancy-rollup` pushed for review only. No production nightly triggered.

## 2026-08-25 — Retire col P workflow: REMs move back to ClickUp

- **Commit:** _(this commit)_
- **File(s):** `snap_shot/rebuild.py`, `snap_shot/tests/test_lock_rollback.py`, `snap_shot/afternoon_sync_sweep.py` (patched to no-op for col P); deleted: `snap_shot/backfill_last_synced_hash.py`, `snap_shot/one_off_wire_orphans_2026_08_18.py`, `.github/workflows/snap-shot-poll.yml`, `.github/workflows/snap-shot-backfill-last-synced-hash.yml`, `.github/workflows/snap-shot-oneoff-wire-orphans-2026-08-18.yml`
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` v1.7 → v1.8 (bump in follow-up commit by main agent)

**Why.** REMs were drifting between typing prospect updates in workbook col P (which auto-posts to ClickUp) and updating tasks directly in ClickUp. This split the Vacancy Pipeline in ClickUp — new prospects lived only in Excel, so pipeline structure decayed. Pulling REMs back to ClickUp-only for prospect work.

**What changed.** Col P (REM ClickUp Comment) removed from workbook schema (COLS dict entry, header, row-write loop, styling, width). `sync_rem_comments_to_clickup` function and `comment-sync` CLI mode deleted. `CU_LAST_SYNCED_HASH_FIELD` constant and all field-read/write logic removed from `pull_lar_summaries`, rollback path, and callers. `extract_ann_edits` return tuple: 5 → 4 (dropped `last_synced_by_prop_unit`). `pull_lar_summaries` return tuple: 5 → 4 (dropped `last_synced_hash_by_task_id`). `check_rem_comment_floor` and `send_unposted_rem_email` deleted (obsolete once col P is gone). `rollback_locked_workbook` simplified — no longer takes `posted_comment_ids` or `stamped_task_ids`; only sends postponement notice. `snap-shot-poll.yml` workflow deleted. `snap-shot-backfill-last-synced-hash.yml` workflow deleted. `snap_shot/backfill_last_synced_hash.py` deleted. `snap_shot/one_off_wire_orphans_2026_08_18.py` deleted (imported deleted rebuild symbols). `snap-shot-oneoff-wire-orphans-2026-08-18.yml` workflow deleted with it. `afternoon_sync_sweep.py` patched to unpack the new 4-tuple and always report 0 flagged comments (workflow retained but effectively a no-op). `test_lock_rollback.py` updated — B2 scenario removed, B scenario asserts postponement email only.

**What did not change.** Col N (Notes), col O (ClickUp Summary), col Q (Last Synced display timestamp). All property-level Ann edits (Broker Calls, Property Flags, Vacant Callouts, Ann's Commentary, Market Rent, TICAM) still round-trip via `extract_ann_edits`. Renewal Pipeline flow — Renewal col O stays as parent status + latest comment. AppFolio rent-roll pull. Nightly (`snap-shot-nightly.yml`), weekly (`snap-shot-weekly.yml`), and afternoon-sync-sweep workflows. Col O rendering — a Thursday change will upgrade Vacancy col O to show prospect subtasks; NOT in this commit.

**Risk / rollback.** Risk: medium. Ann's Wed morning workbook has an empty col P between O and Q. Column visually present but blank (no header, no width setting, no writes). Documented in email to Ann sent 2026-08-25. Rollback: `git revert <sha>`. The `Last Synced Hash` ClickUp custom field will be deleted by main agent AFTER this commit lands. If a rollback is needed within 24h, recreate the field (loss: all historical hash values, which are meaningless post-migration anyway).

**Verification.** After next nightly rebuild (Wed 4:17 AM ET): workbook col P is empty, col Q populated, no errors referencing `CU_LAST_SYNCED_HASH_FIELD`. Test: `pytest snap_shot/tests/` — all 11 tests pass. Test: `python3 -m py_compile snap_shot/rebuild.py` — succeeds.

## 2026-08-19 — Snap Shot: move REM-comment dedup state off col Q and onto a ClickUp custom field (L3)

- **Commit:** _(this commit)_
- **File(s):** `snap_shot/rebuild.py`, `snap_shot/tests/test_lock_rollback.py`, `snap_shot/backfill_last_synced_hash.py` (already committed `8dc048f`), `.github/workflows/snap-shot-backfill-last-synced-hash.yml` (already committed `8dc048f`); probe files deleted: `snap_shot/probe_delegated_workbook_write.py` + `.github/workflows/snap-shot-probe-workbook-write.yml`
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` v1.0 → v2.0 (pending in follow-up)

**Why.** Since 2026-08-17 the comment-sync path has been trying to stamp col Q of the workbook after posting REM comments to ClickUp. Every practical way to do that under our app-only Microsoft Graph auth has failed:
- `PUT /driveItem/content` (whole-workbook upload) fails with HTTP 423 whenever the workbook is open in Excel (frequent — REMs edit live).
- `POST /workbook/createSession` and range `PATCH` on `/workbook/*` both fail with HTTP 403 WAC token errors under app-only auth. These endpoints require delegated auth, which we don't have.

Net effect: comment-sync could never durably reconcile state — either the workbook was locked, or the endpoint was forbidden — and any partial run risked posting a duplicate REM comment on the next attempt (production incident on run 32280301612 required manually deleting 5 duplicates). The dedup state has to live somewhere the workbook lock can't reach, which means it has to live on ClickUp.

**What changed.**
- New ClickUp short-text custom field `Last Synced Hash` (field id `eddb4ff7-5a48-4919-a5bc-1663045c663b`) on all three LAR lists (Renewal `901113575567`, Vacancy `901113575628`, Documents Workflow `901113991446`). Created 2026-08-19; backfilled 120/120 tasks in run 32290365861 from the current Snap Shot workbook's col Q values.
- `pull_lar_summaries()` now also returns `last_synced_hash_by_task_id: {task_id: sha8_hex_or_empty}`, pulled from that custom field on the same request that already fetches the LAR task list. No extra API round-trips.
- `sync_rem_comments_to_clickup()` no longer reads its diff-detect signal from workbook col Q via `_last_synced_hash()`. It reads from `last_synced_hash_by_task_id` per-target-task, so if a unit maps to 2 LAR tasks and only one has the current sha8 stamped, only the un-stamped one gets the comment (previously the workbook Q cell dedup'd both together). After each successful `cu_post_comment` the function now calls `cu_set_field` to stamp the new hash on that task.
- The function still returns `last_synced_by_prop_unit` for backward compat with `build_workbook`, but values are now plain `"YYYY-MM-DD HH:MM ET"` timestamps (no more `· sha8:xxxxxxxx` suffix). Return signature grew to include `stamped_task_ids` — the list of tasks stamped this run, needed by the rollback path.
- `rollback_locked_workbook` (SharePoint 423 recovery) grew a `stamped_task_ids` parameter. When the SharePoint upload fails after `sync_rem_comments_to_clickup` already stamped, the rollback now BOTH deletes the ClickUp comments AND clears (`cu_set_field(..., "")`) the `Last Synced Hash` on each stamped task, so the next successful run re-posts once instead of skipping.
- `run_comment_sync` (poll mode) no longer downloads the workbook to READ col Q, no longer patches col Q, no longer re-uploads to SharePoint. It only fetches col P via extract_ann_edits for the race-guard + comment text, pulls the LAR summaries with fresh hash stamps, and calls sync. Returns `"synced"` on success (was a file path).
- `run_build` (nightly rebuild) still passes col Q display values into `build_workbook`, but drops the col Q read from `extract_ann_edits` (assigned to `_last_synced_by_prop_unit_unused` and ignored). Col Q is now a display-only visible timestamp, not a state store.
- Deleted `map_last_synced_addresses` (dead — was only used by the removed cell-PATCH path). Deleted `_last_synced_hash()` + `_LAST_SYNCED_HASH_RE` regex (dead — no more sha8-in-cell parsing anywhere). Deleted probe files `snap_shot/probe_delegated_workbook_write.py` + `.github/workflows/snap-shot-probe-workbook-write.yml`.
- `snap_shot/tests/test_lock_rollback.py` adds Scenario B2: verify `rollback_locked_workbook` with `stamped_task_ids` clears each task's Last Synced Hash back to `""`. Existing scenarios A/B/C still pass. `cu_set_field` is now monkey-patched in all rollback-invoking tests.

**What did not change.**
- The workbook display: col Q still exists, still shows a human-readable timestamp for the REM. Header text and column width unchanged. Existing REM filters/views referencing col Q continue to work.
- Nightly rebuild logic: still runs at the same cadence, still uploads the workbook to SharePoint the same way, still calls the same rollback on HTTP 423 (with the added stamp-clear step).
- Afternoon sync-sweep workflow, weekly Broker Beat refresh, LAR summary pulling, TICAM pulling, Ann-edit preservation, race guard, REM-comment safety guard, comment attribution format (`📋 LSS note from {REM name} ({date}):`) — all unchanged.
- ClickUp custom-field IDs (Tenant ID, Property ID, Unit #, all other summary fields) — unchanged.
- AppFolio API auth, SharePoint auth, Microsoft Graph secrets, GitHub secrets — unchanged.
- Poll workflow cron cadence, sync-sweep cadence — unchanged.

**Risk / rollback.**
- Risk: **medium.** State migration; the workbook col Q values will diverge from the ClickUp Last Synced Hash values over time (workbook shows plain timestamps; ClickUp holds the sha8). If the ClickUp field is deleted or renamed, comment-sync would treat every comment as brand-new on the next run — that would resurface the 2026-08-17 duplicate-comment failure mode. Mitigation: field is scoped to LAR lists Ann/REMs don't administer, ClickUp doesn't garbage-collect custom fields, and diff-detect degrades gracefully (`""` → comment re-posted) rather than fatally.
- Rollback: `git revert <this-sha>`. Optional cleanup: delete the `Last Synced Hash` custom field from the 3 LAR lists (ClickUp UI). The 120 stamped values do no harm on their own if we revert.

**Verification.**
1. `python3 snap_shot/tests/test_lock_rollback.py` → `ALL SCENARIOS PASSED` (locally verified 2026-08-19 at ~19:13 UTC).
2. First supervised nightly rebuild run → log line `REM comment sync: N unit(s) synced (M ClickUp comment(s) posted, K sha8 stamps written, 0 stamp failures)`.
3. First poll (`--mode=comment-sync`) run after nightly → `comment-sync: 0 REM comment(s) may be pending sync — skipping AppFolio pull` (because the fresh nightly stamps match all extracted comments).
4. Manually edit one REM comment in col P via the Snap Shot workbook. Wait 15 minutes for the next poll. Verify: (a) the comment posts to the matching LAR task, (b) the Last Synced Hash field on that task shows the new sha8, (c) col Q still shows the prior nightly timestamp (not touched by comment-sync mode), (d) the poll after that skips.

## 2026-08-19 — Snap Shot cell-PATCH: rolled back (Graph workbook/* endpoints not supported under app-only auth) (L5)

- **Commits:** `c2cac5d` (revert of `56515f9`) + `7fa90ac` (revert of `5eaebfa`)
- **File(s):** `snap_shot/rebuild.py`, `snap_shot/one_off_wire_orphans_2026_08_18.py`, `.github/workflows/snap-shot-oneoff-wire-orphans-2026-08-18.yml`
- **Author:** Alexis Pattison

**Why.** Two commits earlier today (`5eaebfa` durable session-based cell-PATCH, then `56515f9` sessionless pivot) claimed to introduce a live-write path for col Q that would work while the Snap Shot workbook stayed open in Excel. Both were rolled back today because they never worked in production:
- `POST /workbook/createSession` returned HTTP 403 `AccessDenied: Could not obtain a WAC access token`. Per Microsoft Graph docs, `workbook/createSession` is **not supported with application permissions** — only delegated (work/school) with `Files.ReadWrite`. Our automation runs client-credentials app-only auth.
- The sessionless fallback (range `PATCH` with no session header) failed the same way — HTTP 403 WAC on all 5 test cells. In practice the entire `/workbook/*` endpoint family on SharePoint `/sites/*` paths requires WAC, which app-only auth cannot obtain. The docs' "session header not required" language does not extend to this scenario.
- Net production side-effect: run 32280301612 posted 5 REM comments to ClickUp before failing at createSession — 5 duplicates that had to be manually deleted.

**What changed.**
- Reverted `56515f9` and `5eaebfa` in full. `patch_cells_via_graph`, `_create_workbook_session`, `_close_workbook_session`, the `--skip-comment-posts` flag, and the `apply-stamps-only` workflow mode are all removed from `main`.
- The one-off wire-orphans script + workflow are back to their pre-cell-PATCH shape: mode `dry-run` or `apply`; on `apply`, writes fields + posts comments + calls `upload_to_sharepoint` to stamp col Q (workbook-locked path with existing rollback semantics).
- 5 duplicate ClickUp comments created by run 32280301612 (`90110262260993`, `90110262260994`, `90110262260999`, `90110262261005`, `90110262261009`) were deleted directly from ClickUp. The original 2026-08-17/18 posts remain in place.

**What did not change.**
- Nightly rebuild, poll workflow, afternoon sync-sweep, `upload_to_sharepoint`, `sync_rem_comments_to_clickup`, race guard, sha8 dedup, Ann/edit preservation — all unchanged.
- ClickUp custom-field IDs, comment attribution format, Q-stamp format — all unchanged.
- No secrets changed.

**Risk / rollback.**
- Risk: **very low.** Restoring a state that was in production for weeks. Only visible behavior change: workbook writes to col Q are once again gated on the workbook not being open in Excel (existing race guard).
- Rollback: `git revert c2cac5d 7fa90ac`. Do not do this until we've migrated to delegated auth — see follow-up.

**Verification.**
- The 5 originally-posted LSS comments still exist on Warehouse Cages (`868kt2tzu`), Wingstop (`868gyguen`), Connectivity Source (`868kcmfen`), Donald Bell (`868kdk09t`), Beau Reinmiller (`868kdjzey`). Verified via `clickup_get_task_comments`: each task now shows exactly one LSS note dated 2026-08-17 or 2026-08-18, no duplicate 2026-08-19 entries.
- Poll workflow at 17:01 UTC saw col P/col Q hashes matching for these 5 rows and skipped them as unchanged — confirms col Q was stamped by a prior successful `upload_to_sharepoint` run and no further duplicates will come from the scheduled poll.
- `git log --oneline` shows `7fa90ac` at HEAD, then `c2cac5d`, then the two reverted commits.

**Follow-up.** L3 (to be filed separately): migrate Snap Shot Graph workbook writes to a delegated OAuth flow so `workbook/createSession` becomes available and live cell-write is possible while the team stays in the workbook. Until that lands, col Q is stamped only via `upload_to_sharepoint` (workbook must not be locked open).

## 2026-08-18 — Snap Shot one-off: coerce Property/Tenant ID to string (L1 bugfix)

- **File(s):** `snap_shot/one_off_wire_orphans_2026_08_18.py`
- **Author:** Alexis Pattison

**Why.** First apply run (32151613912) failed 4 field writes with ClickUp `FIELD_018 "Value is not a valid string"`. The Property ID (`057285dd-…`) and Tenant ID (`2f9249fb-…`) fields are ClickUp "short text" type — they reject numeric payloads. WIRE_UP hard-codes ints. On the same run, SharePoint returned HTTP 423 (workbook was open in Excel), so the lock-safe rollback deleted all 5 just-posted comments — clean state to retry.

**What changed.**
- Wrap both `cu_set_field` calls in `str(…)` so pid=282 / tid=1271 / 2032 / 2025 land as strings.

**What did not change.**
- Nothing else. The dry-run plan, the WIRE_UP contents, the workflow YAML, `rebuild.py`, the tests — all identical.

**Risk / rollback.**
- Risk: low. Same idempotent one-off; sha8(P)==sha8(Q) still causes skip on rerun.
- Rollback: `git revert <sha>`.

**Verification.**
- Re-fire `snap-shot-oneoff-wire-orphans-2026-08-18.yml` with `mode=apply`. All 5 rows should show `✓ posted comment`, `✓ stamped Q<row>`, 3 rows should show `✓ Property ID` or `✓ Tenant ID` writes, and the SharePoint upload should succeed (requires workbook not open in Excel).

## 2026-08-18 — Snap Shot: catch-up 5 orphaned REM comments + fix hyphen-padding normalizer (L2)

- **File(s):** `snap_shot/rebuild.py` (normalizer fix), `snap_shot/tests/test_norm_unit_label.py` (new, 8 tests), `snap_shot/one_off_wire_orphans_2026_08_18.py` (new one-off), `.github/workflows/snap-shot-oneoff-wire-orphans-2026-08-18.yml` (new workflow_dispatch)
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` — no user-visible workflow change; skill markdown does not need bumping (durable behavior is "REM comments still round-trip through col P/Q the same way"; the fix is purely internal to the matcher).

**Why.** Six REM comments from the 2026-08-12 comment-sync run orphaned because no matching ClickUp task existed at post-time (the WARNING lines in the poll log). Alexis's assistant located 5 of the 6 tasks. Investigation on the 5:

1. Cumberland Flex | Warehouse Cages (pid=230, tid=1405) → task `868kt2tzu` — task was created 2026-08-17, after the orphan warning. Fields already correct. Only needed the pending col-P comment posted + col Q stamped.
2. Lakeview Village - 181 B-1 / WingStop (pid=54) → task `868gyguen` (Vacancy Pipeline) — task's Unit # is `181 - B1` with space-padded hyphens; AppFolio's `unit_display` returns `181-B1`. `_norm_unit_label` didn't collapse the hyphen padding, so the `(pid, unit_norm)` lookup missed and the comment orphaned. **Root cause = normalizer bug, not missing task.** Fix collapses `" -"` / `"- "` runs. Also affects any future hand-typed unit label with space-padded hyphens.
3. Savannah Crossing - Connectivity Source, LLC - Unit B2 (pid=211) → task `868kcmfen` — task's Tenant ID field was set to 1479, but the correct tenant per Alexis is 1271. Confirmed 2026-08-18 to overwrite.
4-5. Somerset Shoppes - Donald Bell / Beau Reinmiller (pid=282, tids=2032/2025) → tasks `868kdk09t` / `868kdjzey` — both have blank Property ID + Tenant ID. Need both set + comment posted.
6. Clickety Clack Vape and Gifts (pid=57, unit=1396, tid=1348) — assistant could not locate a matching ClickUp task. Left untouched; next scheduled sync will retry once someone creates the task.

**What changed.**
- `snap_shot/rebuild.py`: `_norm_unit_label` now collapses whitespace-padded hyphens (` - ` / ` -` / `- `) to `-` before returning. Test suite at `snap_shot/tests/test_norm_unit_label.py` locks the behavior across 8 cases (bare number, prefix variants, whitespace padding, stacked hyphens, asymmetric padding, empty/whitespace, no-hyphen-space preserved).
- New `snap_shot/one_off_wire_orphans_2026_08_18.py`: two-part fix per task — (a) write Property ID / Tenant ID via `cu_set_field` where blank or where explicitly forced (Connectivity Source only), (b) post the current col-P comment with the standard "📋 LSS note from {REM} ({date})" preamble via `cu_post_comment`, then stamp col Q with `YYYY-MM-DD HH:MM ET · sha8:<hash>` matching what the regular comment-sync writes. Uses the lock-safe upload path from commit 4237c88 — if SharePoint returns HTTP 423 we DELETE the just-posted comments so the next run posts cleanly. Dry-run by default; needs `--apply` to mutate.
- New workflow `.github/workflows/snap-shot-oneoff-wire-orphans-2026-08-18.yml`: `workflow_dispatch` with `dry-run | apply` input, concurrency group `snap-shot` so it can't race the regular poll.

**What did not change.**
- No changes to the regular poll cadence, no cron changes, no ClickUp custom-field id changes, no new secrets.
- The orphan #6 (Clickety Clack Vape and Gifts) is intentionally untouched — no matching ClickUp task exists yet.
- The Wingstop task's Tenant ID stays blank — Vacancy Pipeline matches by (pid, unit), not tenant, and the unit is vacant.
- The existing lock-safe HTTP 423 rollback path from commit 4237c88 is unchanged and re-used here.

**Risk / rollback.**
- Risk: **low**. The normalizer change is additive (only affects labels with whitespace-padded hyphens — previously those unmatched anything; now they match). 8 unit tests cover the boundary conditions including no-hyphen and space-preserved cases. The one-off is dry-run by default and idempotent — a repeat run after a successful apply sees `sha8(P)==sha8(Q)` on those 5 rows and skips them.
- Rollback: `git revert <sha>` for the normalizer change; delete the one-off script + workflow. If a Tenant ID write went wrong on Connectivity Source, edit the field in ClickUp directly.

**Verification.**
- Local unit tests: `python -m unittest snap_shot.tests.test_norm_unit_label` — all 8 pass; `_norm_unit_label('181 - B1') == _norm_unit_label('181-B1') == '181-b1'`.
- Deploy path: workflow_dispatch → dry-run → apply. Dry-run should log all 5 wire-up rows with the current col-P comment preview and Q-cell address. Apply should log 5 `cu_set_field` writes (some as no-op if the field's already correct), 5 `cu_post_comment` results with comment IDs, 5 col-Q stamps, and a successful SharePoint upload.
- Post-apply: the next comment-sync run should log `98 unchanged skipped` (or similar) and 0 orphans from these 5 units.

## 2026-08-12 — Snap Shot: retire one-off stamp workflow (post-cleanup) (L2)

- **File(s):** `.github/workflows/snap-shot-oneoff-stamp-locked-run.yml` (deleted), `snap_shot/one_off_stamp_locked_run.py` (deleted)
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` — no change.

**Why.** The one-off surgical fix for the 2026-08-12 10:42 AM locked run has been applied successfully (GitHub Actions run 31632120793): 10 Q-column stamps recorded, workbook uploaded. Retiring the script + workflow so no one runs it again by accident. The permanent lock-safe path in `snap_shot/rebuild.py` handles this class of failure going forward.

**What changed.**
- Deleted `snap_shot/one_off_stamp_locked_run.py`.
- Deleted `.github/workflows/snap-shot-oneoff-stamp-locked-run.yml`.

**What did not change.**
- All production behavior identical to the state after commit 4237c88.

**Risk / rollback.**
- Risk: none. Only removes a one-shot script that has already served its purpose.
- Rollback: `git revert <this sha>`.

**Verification.**
- Apply run 31632120793 completed successfully: 10 stamp(s) recorded, workbook uploaded.
- The next comment-sync run should show "98 unchanged skipped, 1 (or so) pending sync" — the pending count reflecting new REM edits since the apply, NOT the 10 residuals from the failed run.

## 2026-08-12 — Snap Shot: rollback ClickUp comments when SharePoint returns HTTP 423 (workbook locked in Excel) (L2)

- **File(s):** `snap_shot/rebuild.py`, `snap_shot/tests/test_lock_rollback.py` (new)
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` — no user-visible workflow change; skill markdown does not need bumping.

**Why.** The 10:42 AM ET poll run on 2026-08-12 (run 31608260834) failed with HTTP 423 `resourceLocked` from SharePoint at the end of a comment-sync cycle, right after successfully posting 10 REM comments to ClickUp. The workbook was open in Excel at the moment we tried to save the patched Q-column (Last Synced) copy, so SharePoint refused the upload. The 10 comments landed on ClickUp tasks, but the paired Last Synced stamps never made it back to the workbook — meaning the next successful comment-sync run would see the same 10 comments as "still pending" (missing sha8 in col Q) and post duplicates.

This is a class of failure: any time a REM has the file open when the poll fires, we can duplicate their notes on the ClickUp side. Fix it once so it self-heals.

**What changed.**
- New `WorkbookLockedError` exception, distinct from `FatalError`. Signals "transient lock, not a real failure" so `main()` doesn't page anyone or exit non-zero.
- `upload_to_sharepoint` now detects HTTP 423, waits 30s, retries once, and if still 423 raises `WorkbookLockedError` (both the simple-PUT branch and the resumable-upload-session branch). Other statuses still raise `FatalError` unchanged.
- `cu_post_comment` now RETURNS the created comment's id (was: returned nothing). Callers ignore it unless they need to roll back.
- New `cu_delete_comment(comment_id)` helper. Best-effort; logs but never raises.
- `sync_rem_comments_to_clickup` now tracks every successfully-posted comment id and returns them as a third tuple element `posted_comment_ids`. Callers updated: `run_comment_sync`, `run_build`.
- New `rollback_locked_workbook(posted_comment_ids, context, error_detail)` helper. Called by `run_comment_sync` and `run_build` when `upload_to_sharepoint` raises `WorkbookLockedError`. Deletes each posted comment via `cu_delete_comment`, logs deleted/failed counts, and sends a friendly "postponed" notice to `ERROR_NOTIFICATION_TO` (distinct subject from a FATAL notice). The workbook is untouched — the SharePoint copy is exactly as it was before the run.
- The rollback is bounded to comments POSTED THIS RUN. Comments from earlier successful runs are safe — they already have their Last Synced stamp round-tripped to SharePoint, so the sha8 dedup won't touch them.
- New self-contained test at `snap_shot/tests/test_lock_rollback.py` covering (a) happy path, (b) rollback deletes every posted comment id, (c) `upload_to_sharepoint` raises `WorkbookLockedError` not `FatalError` on HTTP 423.

**What did not change.**
- No trigger, no field, no ClickUp custom-field id, no cron schedule, no workflow YAML.
- Behavior when the workbook is NOT locked: identical to before this commit — posts happen, stamps land, upload succeeds, `posted_comment_ids` is discarded on the success path.
- The race-condition guard (`is_race_condition`, `RACE_CONDITION_WINDOW_MINUTES`) is unchanged. It still short-circuits BEFORE any ClickUp post if a human touched the workbook in the last 15 min. The lock rollback only fires when the race guard PASSED (so we started posting) but SharePoint refused the upload anyway — typically because a fresh Excel session opened in the ~30 seconds after the race check.
- No secret rotation. No new env vars.
- `send_unposted_rem_email` (for "no matching ClickUp task") is unchanged and complementary — different failure mode.

**Risk / rollback.**
- Risk: **low**. Only fires on HTTP 423; every other code path is identical to pre-change. The rollback deletes ClickUp comments we know we just created (by id) — no wildcard deletions, no risk of removing older content. Wrong-side failure mode: if `cu_delete_comment` itself fails, the notice email lists the comment id so it can be manually cleaned up before the next successful run.
- Rollback: `git revert <sha>`. There's no schema change, no secret change, no external dependency — pure code.

**Verification.**
- Local unit tests all pass:
  - Scenario A (happy path): 3 posted, 0 deleted ✅
  - Scenario B (locked → rollback): 5 posted → all 5 deleted, notice email rendered ✅
  - Scenario C: `upload_to_sharepoint` sees repeated HTTP 423 → raises `WorkbookLockedError` (not `FatalError`) ✅
- `python3 -m py_compile snap_shot/rebuild.py` clean.
- Live verification: next time a REM has the workbook open when the poll fires, we should see "workbook locked — rolling back N comment(s)" in the run log, a "postponed" notice email in Alexis's inbox, and the run exits 0 (not red X). The next successful fire posts those same N comments once and stamps col Q.

## 2026-08-11 — Snap Shot: add afternoon sync-health sweep (5 PM weekdays) (L2)

- **File(s):** `snap_shot/afternoon_sync_sweep.py` (new), `.github/workflows/snap-shot-afternoon-sync-sweep.yml` (new)
- **Author:** Alexis Pattison
- **Skill:** none new; monitor lives alongside `prudent-snap-shot-rebuild` but doesn't touch the rebuild flow.
- **Scheduling:** GitHub Actions cron `0 21 * * 1-5` (5 PM ET Mon–Fri EDT / 4 PM EST). Shares `snap-shot` concurrency group so it cannot overlap rebuilds. Also supports `workflow_dispatch` with a `stale_hours` input.

**Why.** Now that `--mode=comment-sync` (added earlier today) pushes REM notes to ClickUp on the every-15-min poll cadence, we need an end-of-day check that catches any comment that failed to sync during the workday, before it leaks into the next morning. Requested by Alexis so she can troubleshoot the automation flow before REMs notice.

**What changed.**
- New standalone monitor `snap_shot/afternoon_sync_sweep.py`. Downloads the current SharePoint workbook, extracts col-P (REM comment) + col-Q (Last Synced) cells via `extract_ann_edits`, and flags any unit where col P has text AND col Q is either (a) blank, (b) unparseable, or (c) has a timestamp older than STALE_HOURS (default 4).
- Silent when everything is healthy. Sends an email to `ERROR_NOTIFICATION_TO` (apattison@prudentgrowth.com) via Microsoft Graph only when at least one unit is flagged.
- Also sends a distinct failure email if the SharePoint download itself fails, so silence isn't ambiguous.
- Reuses `download_from_sharepoint`, `extract_ann_edits`, `get_graph_access_token`, `http_request`, `now_et`, `ET` timezone constant from `rebuild.py`. No new dependencies. No new secrets.

**What did not change.**
- Zero changes to `rebuild.py`, `sync_rem_comments_to_clickup`, or any comment-sync logic. This is a passive monitor — it only reads.
- No changes to any existing workflow file. The one new workflow file only schedules this monitor; it does not touch nightly/weekly/poll.
- The rebuild's own `send_unposted_rem_email` (which fires from inside the nightly rebuild when a comment can't find a ClickUp task) is unchanged and complementary — it catches a different failure mode (missing ClickUp task) at a different time (post-rebuild).

**Risk / rollback.**
- Risk: **low**. Read-only against SharePoint. Sends at most one email per weekday. Wrong-side failure mode: false positive (emails Alexis when nothing's actually broken) — easy to tune via `--stale-hours`. No workbook writes, no ClickUp writes.
- Rollback: rename `.github/workflows/snap-shot-afternoon-sync-sweep.yml` to `.yml.disabled` to hard-stop, or `git revert 76e7a37` for the workflow + `git revert 9255dbd` for the script.

**Verification.**
- Local unit test on `/tmp/Leasing-Snap-Shot-dryrun.xlsx` with three injected states — fresh stamp (1h old → healthy), stale stamp (6h old → flagged), missing stamp (→ flagged). Result: 1 healthy / 2 flagged, correct reasons.
- Stamp parser test: `"2026-08-11 15:22 ET · sha8:..."` → parses; garbage input → None; empty → None.
- `python3 -m py_compile snap_shot/afternoon_sync_sweep.py` clean.
- Email body rendered locally — sections in order, workbook link intact, troubleshooting checklist embedded.
- Live verification: first scheduled fire tomorrow (Wed 2026-08-12) at 5 PM ET. Expect silence unless something is genuinely stale.

## 2026-08-11 — Snap Shot: add `--mode=comment-sync` for ~15-min REM comment latency (L3)

- **File(s):** `snap_shot/rebuild.py`, `.github/workflows/snap-shot-poll.yml`
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` v1.5 → v1.6 (documents the new mode)

**Why.** REMs (Ann + team) were reporting that comments they typed into the Prudent Snap Shot workbook in the morning weren't appearing in ClickUp until the next-day 4:17 AM nightly rebuild. Root cause: the every-15-min poll workflow only actually rebuilt the workbook when the ClickUp "Rebuild Snap Shot" checkbox was checked — otherwise it was a no-op. That checkbox is rarely used, so col-P notes typically waited 12–18 hours to reach ClickUp. Fix: add a lightweight comment-sync mode that runs on every poll fire, pushing REM notes to ClickUp within ≈15 min of typing.

**What changed.**
- New function `map_last_synced_addresses(workbook_bytes)` in `rebuild.py` — walks the workbook once and returns `{(PropertyId, unit_label): "Q42"}` for every unit row, reusing the exact block-detection heuristic that `extract_ann_edits` uses so key coverage matches 1:1.
- New function `run_comment_sync(dry_run)` in `rebuild.py` — orchestrator: SharePoint download → race-condition guard → extract REM comments + Last Synced stamps → short-circuit if all comments already have matching sha8 (no ClickUp / AppFolio calls in that path) → otherwise pull AppFolio rent-roll + ClickUp LAR summaries → `sync_rem_comments_to_clickup` → patch only the changed Q-cells → upload. Does NOT re-render the workbook, so Ann's other in-flight edits are untouched.
- New CLI mode `--mode=comment-sync` wired into `main()` argparse (adjacent to `poll`).
- `snap-shot-poll.yml` now runs two steps on each fire: (a) the existing `--mode=poll` (honors the ClickUp checkbox) then (b) the new `--mode=comment-sync`. If step (a) does a full rebuild, step (b) sees all comments already synced and no-ops in <5s.
- Header docstring in `rebuild.py` updated to list the new mode.

**What did not change.**
- Every existing safety guard remains in place unchanged:
  - `RACE_CONDITION_WINDOW_MINUTES = 15` — comment-sync aborts if a non-automation edit landed within the window (deferred to next fire).
  - `snap-shot` concurrency group — comment-sync cannot overlap nightly / weekly / poll rebuilds.
  - sha8 dedup in `_comment_hash` / `_last_synced_hash` — unchanged text produces zero duplicate ClickUp comments.
  - REM Comment floor guard in `check_rem_comment_floor` — unchanged; comment-sync doesn't rebuild, so the guard is irrelevant to this path.
- No changes to workbook rendering, AppFolio pull logic, `sync_rem_comments_to_clickup` itself, or any nightly/weekly cadence.
- No new secrets or permissions. Comment-sync uses the same `APATTISON_MS_REFRESH_TOKEN`, `APPFOLIO_API_SECRET`, `CLICKUP_API_TOKEN` the rebuild already uses.
- ClickUp checkbox path preserved via `workflow_dispatch` on the poll workflow AND on each scheduled fire.

**Risk / rollback.**
- Risk: **medium**. Adds one extra script invocation per poll fire (~44 fires/day M–F). Runtime is dominated by AppFolio rent-roll pull (≈10s) + ClickUp LAR pull (≈10s) — but only when there are pending comments; the sha8 short-circuit skips both entirely when nothing is pending. Idempotency verified end-to-end via a local injection test: inject col-P comment → extract → patch Q → second extract → pending=0. ClickUp API budget (100 req/min) is nowhere near saturated.
- Rollback: revert this commit (restores `poll`-only workflow). Or edit `.github/workflows/snap-shot-poll.yml` and delete the second step. No workbook state is stranded — the sha8 dedup means the next nightly rebuild would re-sync anything comment-sync missed.

**Verification.**
- Local unit test on `/tmp/Leasing-Snap-Shot-dryrun.xlsx`: `map_last_synced_addresses` returns 898 Q-cell entries; every key produced by `extract_ann_edits` has a matching address.
- End-to-end simulation on Property 279 Parking Lot (Q28): injected "Test comment from Leah" → extract sees 1 REM comment → stamp written → re-extract confirms Last Synced = new stamp — idempotency check: pending=0 on second run. Test script embedded in the commit context.
- `python3 -m py_compile snap_shot/rebuild.py` clean. `python3 snap_shot/rebuild.py --help` shows `comment-sync` in choices.
- Live verification plan: watch first scheduled fire after merge (next `*/15 11-22 * * 1-5` slot). Confirm the two steps run in order and step (b) finishes in <15s. Also ask Ann to leave a test comment on any unit and confirm it lands in ClickUp within one poll cycle.

## 2026-08-07 — Snap Shot: restore paused schedules + delete _recovery/ folder (L3)

- **File(s):** `.github/workflows/snap-shot-nightly.yml`, `.github/workflows/snap-shot-poll.yml`, `.github/workflows/snap-shot-weekly.yml`, `.github/workflows/snap-shot-delete-recovery-folder.yml` (new, one-off), `snap_shot/delete_recovery_folder.py` (new)
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` unchanged (behavior described in v1.5 is now on its normal cadence)

**Why.** All three Snap Shot schedules have been paused since Wednesday's REM Comment wipe investigation (banner: `[PAUSED 2026-08-07 — REM Comment wipe investigation]`). The two follow-up L2 fixes have since landed and been verified end-to-end: yesterday's unposted-REM visibility (commit 9038a91) and today's (PropertyId, Unit#) unit-label normalization (commit d5106e8, verified against Barker Cypress Vacancy task 868g5mkpu — comment 90110258768064 posted successfully via manual dispatch run 31191836362). The workbook has fully rebuilt twice today with no regressions. Restoring cadence for Monday morning. Also cleaning up the SharePoint `_recovery/` folder created during the incident — the three RECOVERY snapshots (v01/v02/v04) served their diagnostic purpose and are no longer needed.

**What changed.**
- `snap-shot-nightly.yml` — `schedule: - cron: '17 8 * * *'` restored (4:17 AM ET EDT / 3:17 AM EST). Pause banner replaced with a two-line note recording the pause + restore dates.
- `snap-shot-poll.yml` — `schedule: - cron: '*/15 11-22 * * 1-5'` restored (every 15 min, 7 AM–6 PM ET, Mon–Fri). Same pause-note treatment.
- `snap-shot-weekly.yml` — `schedule: - cron: '23 9 * * 1'` restored (5:23 AM ET Mondays, Broker Beat refresh). Same pause-note treatment.
- New `snap_shot/delete_recovery_folder.py` — uses the same delegated-refresh-token auth as `list_recovery_folder.py`, does a single `DELETE /drives/{DRIVE_ID}/items/{FOLDER_ID}` to drop the `_recovery/` folder recursively, then verifies. Idempotent (404-safe).
- New `.github/workflows/snap-shot-delete-recovery-folder.yml` — one-off `workflow_dispatch`-only workflow that runs the delete script. Both the workflow file and the script will be removed in a follow-up commit once the one-off run confirms the folder is gone.

**What did not change.**
- No rebuild-logic changes. Restored crons are the exact expressions from the pre-incident state.
- No changes to `snap-shot-recover-versions.yml` (kept in place; historical record of the incident lives in its previous runs).
- No new secrets or permission scopes — the delete script reuses `APATTISON_MS_REFRESH_TOKEN` with the same `Sites.ReadWrite.All` scope the rebuild already exercises.
- Ann's edits + REM Comment guard behavior in `rebuild.py` — unchanged; the 2026-08-06 archive folder and REM Comment floor guard are still the safeties on any live run.

**Risk / rollback.**
- Risk: **medium**. Re-enables ~48 fires/day on the poll + 1/day nightly + 1/week Monday. All three run behind the `snap-shot` concurrency group so they cannot overlap. The two L2 fixes have been verified; workbook rebuilds twice today (11:05 ET + 11:17 ET) landed cleanly on SharePoint with no comment loss.
- Rollback (schedules): `git revert <sha>` restores the paused blocks. Alternatively, rename the three `.yml` files to `.yml.disabled` to hard-stop.
- Rollback (recovery delete): once the delete workflow runs, the folder is gone from SharePoint's live view but retained in the site recycle bin for 93 days per tenant policy. A restore can be done from the SharePoint site recycle bin UI without touching the code.

**Verification.**
- Schedules: `.github/workflows/snap-shot-*.yml` schedule blocks are uncommented and grep-clean of `[PAUSED]`. First expected fire: Monday 5:23 AM ET (weekly), then 7:00 AM ET poll, then Tuesday 4:17 AM ET nightly.
- Recovery delete: dispatch `Snap Shot — delete _recovery/ folder (one-off)`; run log prints `Found _recovery/`, contents list, `HTTP 204`, then `Verified: _recovery/ no longer exists.` Re-run search on SharePoint returns only `Leasing-Snap-Shot.xlsx` (no `_RECOVERY_v*` files).
- Team notice: heads-up email sent to Ann + REM group Friday afternoon explaining the moccasin orange col P highlight from the L2 unposted-REM visibility change (commit 9038a91) so nobody is confused when they see it Monday.

---

## 2026-08-07 — Snap Shot: normalize unit labels between AppFolio and ClickUp (L2)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` v1.4 → v1.5 (behavior fixed: (PropertyId, Unit#) matching for the Vacancy Pipeline)

**Why.** Immediately after the L2 unposted-REM visibility change landed (commit 9038a91), Alexis pointed out that the one unit flagged as unposted — PropertyId=272, Unit 155, Barker Cypress Marketplace — actually DID have a matching task in the Vacancy Pipeline (ClickUp task `868g5mkpu`). The lookup missed because AppFolio's `unit_display()` returns `'Unit 155'` (with the 'Unit ' prefix) while the Vacancy task's `Unit #` custom field stores just `'155'` (bare number). The `sync_rem_comments_to_clickup()` dict lookup keyed on `(pid, unit_label)` never hit, the comment never posted, and the unit lit up orange in the workbook. Same asymmetry affects the ClickUp Summary passthrough in `write_property_block`. Alexis: "remember all rows that don't have a Tenant ID need to be matched to the Vacancy Pipeline using the Property ID and the Unit # and I found that corresponding task in that list."

**What changed.**
- New `_norm_unit_label(s)` helper (defined right after `unit_display()`). Strips common suite/unit prefixes (`unit`, `suite`, `ste`, `ste.`, leading `#`) and lowercases so `'Unit 155'`, `'155'`, `'#155'`, `'Suite 300A'`, and `'Ste. 300'` all reduce to a canonical form (`'155'`, `'155'`, `'155'`, `'300a'`, `'300'`).
- `pull_lar_summaries()` — builds `summaries_by_prop_unit` and `task_ids_by_prop_unit` on the NORMALIZED unit key (was: raw stripped `str(unit_raw).strip()`). Same variable renamed from `unit` to `unit_norm` for clarity.
- `sync_rem_comments_to_clickup()` — builds `tid_by_prop_unit` on the NORMALIZED unit key from `unit_display()`. Loop over `rem_comments_by_prop_unit.items()` now computes `unit_norm = _norm_unit_label(unit_label)` once per iteration and uses it for both the `tid_by_prop_unit` and `task_ids_by_prop_unit` lookups. `last_synced_by_prop_unit` still uses the raw `unit_label` because that dict round-trips through the SharePoint workbook and must stay workbook-keyed.
- `write_property_block()` — the ClickUp Summary lookup for occupied-tenant-less rows now normalizes the workbook's `unit_label` before hitting `clickup_summaries_by_prop_unit`, matching the ClickUp-side keys built above.

**What did not change.**
- `unit_display()` itself — still returns 'Unit 155' style labels for the workbook.
- Workbook keying (`rem_comments_by_prop_unit`, `last_synced_by_prop_unit`) — still keyed on the raw `unit_val` from the rent-roll cell, so REM edits round-trip through SharePoint the same way.
- TenantId matching path — unchanged; only the (PropertyId, Unit#) path required normalization.
- No changes to task creation, comment posting, or Last Synced stamp behavior beyond the fact that lookups now hit the tasks they were supposed to hit all along.
- Yesterday's L2 unposted-REM visibility (email + orange col P) is untouched — it will just have fewer entries once this fix lands.

**Risk / rollback.**
- Risk: **low**. The normalizer is a pure function, unit-tested for 17 label variants. It only widens what matches — it cannot cause a wrong-task post because both sides are normalized identically, and false positives require two different physical units to normalize to the same canonical form (e.g. `'Unit 155'` and `'155'` at the SAME property, which is nonsensical).
- Rollback: `git revert <sha>`. Restores the prefix-sensitive lookup; Unit 155 will re-appear in the unposted-REM orange list on the next run.

**Verification.**
- Local unit tests: 17/17 pass, covering `'Unit 155'`, `'155'`, `'Suite 300A'`, `'Ste 300A'`, `'Ste. 300'`, `'#155'`, `'# 155'`, `'B101'`, `None`, `''`, `'Unit '`, `'Unit'`, `'united'` (must NOT strip), `'Unit A101'`.
- Local integration test: simulated `task_ids_by_prop_unit[('272', _norm_unit_label('155'))] = ['868g5mkpu']`; lookup with `('272', _norm_unit_label('Unit 155'))` returns `['868g5mkpu']`. Negative case (`'Unit 200'`) correctly misses.
- Production verification: manual dispatch of the Snap Shot rebuild after this commit — the `unposted-REM` email should either not send (empty list) or omit PropertyId=272 Unit 155, and col P for that unit should render gray, not orange, in the next SharePoint upload.

---

## 2026-08-07 — Snap Shot: visibility for REM comments that couldn't post to ClickUp (L2)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` v1.3 → v1.4 (behavior added: unposted-REM email + orange col-P highlight)

**Why.** After yesterday's REM Comment extractor fix (v1.3), the first production rebuild succeeded but silently skipped 1 comment: PropertyId=272 Unit 155 had no matching Renewal / Vacancy / Documents Workflow task, so `sync_rem_comments_to_clickup` logged a warning and moved on. The warning only lives inside the GitHub Actions log — nobody reads that in the normal course of business — so the REM's comment sat in col P with no signal that ClickUp didn't receive it. Alexis: "if a REM comment doesn't post because it can't find a task to post the comment, how do I know?"

**What changed.**
- `sync_rem_comments_to_clickup` return value changed from `dict` to `(dict, list)`. The new list is the "unposted" collection — one entry per unit whose comment couldn't post because no matching task was found. Each entry captures `pid`, `unit`, `tenant_id`, `rem` (REM name from mapping), and `comment` (the REM-authored text). Empty when everything posts cleanly.
- New `send_unposted_rem_email(unposted_units, sharepoint_link)` helper — reuses the same Graph OAuth path as `send_error_email`. Sends a plain-text summary of each unposted unit (property id, unit, tenant id, REM, comment preview capped at 500 chars) to `ERROR_NOTIFICATION_TO` (apattison@prudentgrowth.com). Non-fatal: mail failure logs a warning; rebuild is already successful before this fires.
- `run_build` — Step 7c added, right after `update_control_task_with_refresh_link`. Only fires on live runs (not dry-run) and only when `unposted_rem_units` is non-empty.
- `build_workbook` — new `unposted_rem_units` kwarg. Filters into a per-property `_unposted_units_set` on each mapping entry so the renderer can do a simple set membership check.
- `write_property_block` — new `unposted_units_set` local. When writing col P for a unit whose label is in that set, cell fill is overridden to `#FFE4B5` (soft moccasin orange), applied AFTER the row_fill / column-fill logic so it wins over the lease-expiry row tint AND the default gray. Everything else on that row is untouched.

**What did not change.**
- The sync logic itself — same warning is still logged, same "leave col P alone, do not stamp Last Synced" behavior. This change is pure visibility; it doesn't alter what posts or what doesn't post.
- No change to the workbook write path for units that DO have matching tasks (col P fill still gray, Last Synced stamps still land in col Q).
- No new dependencies, secrets, or environment variables.
- ClickUp is not touched by this change.
- Neither guard from yesterday's L3 was modified (REM Comment floor guard, archive folder, force_rebuild flag all still behave identically).

**Risk / rollback.**
- Risk: **low**. All changes are additive. The email is best-effort (wrapped in try/except); the orange fill is a cosmetic override on a single cell per unposted unit.
- Rollback: `git revert <sha>` restores the tuple → dict return signature and drops the email + orange highlight. No data cleanup needed — the workbook renders fine either way.

**Verification.**
- Local synthetic test (`/tmp/highlight_test.xlsx`): 3-unit property with 1 unit in unposted set → confirmed col P renders as `#FFE4B5` for that unit only, other two units stay gray.
- Local sync return-shape test: 2 comments (1 has a matching task, 1 doesn't) → confirmed `(stamps, unposted)` tuple where `unposted` contains only the unmatched unit with all 5 fields populated correctly.
- Production verification: manual dispatch with `force_rebuild=true` — Unit 155 (PropertyId=272) still has no matching task, so the run should (a) send an email to apattison@prudentgrowth.com containing that unit, and (b) render col P for Unit 155 as orange in the uploaded workbook.

## 2026-08-07 — Snap Shot: fix REM Comment wipe + 4 safety guardrails (L3)

- **File(s):** `snap_shot/rebuild.py`, `.github/workflows/snap-shot-nightly.yml`, `.github/workflows/snap-shot-poll.yml`, `.github/workflows/snap-shot-weekly.yml`, `snap_shot/RCA_2026-08-07.md`
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` v1.2 → v1.3 (behavior added: REM Comment safety guard, SharePoint _archive/ folder, per-block extraction diagnostics)

**Why.** Leah added 4 REM Comments in the Snap Shot the evening of 2026-08-06. The next scheduled rebuild wiped them all before they synced to ClickUp. Full RCA in `snap_shot/RCA_2026-08-07.md`. Root cause: `NOTE_LABEL_TO_KEY` mapped 6 labels but the renderer writes 7 note rows ("Broker / Contact" was missing from the map). The 6-vs-7 mismatch caused the extractor's label-consumption loop to overshoot the RENT ROLL header by 1 row on every block, so every scheduled rebuild since Commit 2 shipped has silently extracted 0 REM Comments and wiped whatever REMs had typed.

**What changed.**
- `rebuild.py`
  - `NOTE_LABEL_TO_KEY` — added `"broker / contact": "broker_contact"`. Extracted value is intentionally ignored by `write_property_block` (that row is auto-populated from the ClickUp Broker Directory on every rebuild); the map entry exists so the label-consumption loop consumes all 7 rows and stops in the right place. **This alone is the fix.**
  - `extract_ann_edits` — RENT ROLL scan now starts from `notes_start` (top of the note block) with a wider 25-row window instead of from `r` (post-loop cursor). Defensive against future label-count drift.
  - New `check_rem_comment_floor()` — refuses to proceed with rebuild if extracted REM Comment count drops to 0 (hard floor) or by >50% (soft floor) versus the previous run. Bypassed by `force_rebuild=true`. State persisted to `snap_shot/data/last_rem_count.json`.
  - `run_build` — calls `check_rem_comment_floor` right after extraction, before AppFolio pull, so a tripped guard fails fast without wasting API calls.
  - New `archive_current_snap_shot()` — uploads the current live Snap Shot to `WORKING DOCS/_archive/Leasing-Snap-Shot__pre_{utc-timestamp}.xlsx` before every overwrite. Retention: 14 days, auto-pruned. Non-fatal: archive failures log a warning but do not block the upload.
  - `run_build` — calls `archive_current_snap_shot(current_wb_bytes)` right before `upload_to_sharepoint`. Reuses the bytes already downloaded, no second Graph call.
  - Per-block extraction diagnostic — if extraction returns 0 REM comments across >0 blocks, dumps per-block note-key detail so future debugging doesn't require a re-run.
  - `--force-rebuild` CLI flag (also reads env `SNAP_SHOT_FORCE_REBUILD`) plumbed through `run_build` and `poll_and_maybe_rebuild`. Bypasses the new REM Comment safety guard.
- Workflows
  - `snap-shot-nightly.yml` — `force_rebuild` input description updated (now covers both guards); new `SNAP_SHOT_FORCE_REBUILD` env var propagated to the script.
  - `snap-shot-poll.yml`, `snap-shot-weekly.yml` — added the same `force_rebuild` workflow_dispatch input and env var wiring so all three workflows have consistent manual-override behavior. Weekly also gained the race-window override that nightly already had.

**What did not change.**
- Cron schedules — all three (`snap-shot-nightly`, `snap-shot-poll`, `snap-shot-weekly`) remain PAUSED from the prior commit. Restoring them is a separate commit once we've dispatched manual test runs and confirmed the fixes work end-to-end against production data.
- `write_property_block` (renderer) — unchanged. Still writes the same 7 note rows in the same order and still ignores `broker_contact` when reading `_notes` back.
- REM → ClickUp sync logic (`sync_rem_comments_to_clickup`) — unchanged. The wipe wasn't caused by the sync; it was caused by extraction returning zero to feed back into the render.
- TICAM row layout — unchanged. Coincidentally shipped in the same window as heavy Leah REM-comment usage but was NOT the cause.
- Race-condition window logic — unchanged; already had a `force_rebuild` bypass.
- No ClickUp field IDs, no auth, no data source URLs, no email routing.

**Risk / rollback.**
- Risk: medium. The fix is small and syntactically minimal, but it changes what the extractor "sees" on every scheduled run. Mitigations: (1) unit-tested locally with a synthetic 2-property post-TICAM workbook — 3 REM comments extracted, 0 warnings; (2) unit-tested the safety guard through 8 scenarios (hard floor, soft floor, first-run, bypass, upgrade path) — all pass; (3) crons stay paused until a manual dispatch verifies against production data; (4) `_archive/` folder gives us a 30-second SharePoint recovery if anything else surfaces.
- Rollback: `git revert 2887b68..HEAD` reverts fixes; workbook fallback to bundled JSON still works (well-exercised code path). Or leave crons paused indefinitely if guards keep tripping.

**Verification.**
- Local: `python3 -m py_compile snap_shot/rebuild.py` — PASS.
- Local synthetic test: 2-property post-TICAM workbook with 3 REM comments distributed across blocks — all 3 extracted correctly, log shows "3 REM ClickUp Comments, 3 Last-Synced stamps", zero "no RENT ROLL section found" warnings. Contrast with last live nightly (run 31173169805) that showed "0 REM ClickUp Comments" and 72 identical RENT ROLL warnings.
- Local safety-guard test: 8 scenarios all pass.
- Production: manual `gh workflow run snap-shot-nightly.yml` dispatch after commit. Success criteria: log shows `>=4 REM ClickUp Comments` extracted (Leah's are back in the live workbook now), REM count guard logs `OK to proceed`, `_archive/` folder receives its first snapshot copy, no `no RENT ROLL section found` warnings, workbook uploaded successfully.
- Restore schedules only after production dispatch passes.

## 2026-08-06 — Snap Shot: bump ClickUp Summary font 8pt → 10pt (L1)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison
- **Skill:** no bump (behavior described in v1.2 still current)

**Why.** Alexis asked for the ClickUp Summary column (col O) to be more readable. 8pt italic slate was hard to scan next to the 9pt navy REM Comment column, especially on Renewal/Vacancy summaries that pack a lot of information into the bullet list.

**What changed.**
- `rebuild.py` — ClickUp Summary cell font raised from `size=8` to `size=10` (italic + SLATE color unchanged). Row-height priming for long summaries adjusted from `>100 chars/line, 14pt line-height` to `>85 chars/line, 16pt line-height` so multi-line summaries still render fully wrapped without truncation at the new font size. Base row height for a wrapped summary raised from 17 to 19.

**What did not change.**
- Column O width (still 80).
- Font family (Calibri), italic styling, SLATE color, top-aligned + wrap.
- REM ClickUp Comment column (col P) still 9pt navy — the intentional visual contrast between muted read-only summary and full-brightness editable comment is preserved.
- No other font sizes touched (headers, banner, TICAM row, Last Synced all unchanged).
- No ClickUp fields, no cron schedule, no auth, no data source, no skill behavior change.

**Risk / rollback.**
- Risk: low. Cell-formatting-only change. Worst case: some very long summaries now render at height 409 (the max cap) and Excel adds a scroll clip inside the cell — visually unpleasant but not data loss.
- Rollback: revert commit.

**Verification.**
- Local syntax check passed (`python3 -m py_compile snap_shot/rebuild.py`).
- Next nightly rebuild will apply the new font. On any property block with a Renewal or Vacancy row, col O text should be visibly larger and match the readability of the surrounding data cells while remaining italic + muted.

## 2026-08-06 — Snap Shot: add 2026 TICAM (NNN) row per property (L2)

- **File(s):** `snap_shot/rebuild.py`, `snap_shot/data/property_ticam_map.json` (new)
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` v1.1 → v1.2

**Why.** Alexis asked for the annual NNN estimates (PGP CAM inc. admin fee, Tax, Insurance, Water, Assoc. Fee — per SF) to appear on each property block so brokers see the current pass-through economics without leaving the Snap Shot. Data source is the ClickUp "TICAM Rates" list (`901112111796`) filtered to Year=2026. Only the five per-SF fields she named are shown; blanks are omitted so we never render "$0.00".

**What changed.**
- Added `pull_ticam_rates_2026(snap_shot_property_names)` in `rebuild.py`. Discovers custom-field IDs by name via `GET /list/{TICAM_LIST_ID}/field` so a ClickUp field rename doesn't silently break the fetch. Filters to Year=="2026", drops rows where all five rate fields are None/0, resolves each row's property via `Property (A/O)` dropdown first with a fallback to the task's `name` field (most existing TICAM tasks don't have the dropdown set). Handles ClickUp's two `status` payload shapes (dict on REST v2 list-tasks; bare string on some paginated shapes). Non-fatal: any exception returns `{}` and the workbook still builds with "2026 TICAM: not published" everywhere.
- Added `snap_shot/data/property_ticam_map.json` — Snap Shot name → TICAM dropdown name overrides for the 13 properties whose names don't already match verbatim, including a two-entry array for South Memorial (which has both "South Memorial" and "Tulsa Memorial" dropdown options pointing at the same asset).
- Added `write_ticam_row(ws, row, ticam_data)` renderer. Writes one merged A:Q row with light-gray fill and navy 9-pt text just below the address/REM row. Format: `2026 TICAM · Confirmed: PGP CAM $X.XX  ·  Tax $X.XX  ·  Ins $X.XX  ·  Water $X.XX  ·  Assoc $X.XX    →  {clickup_url}`. When no 2026 row exists (or every rate is blank), renders `2026 TICAM: not published` in italic slate. The ClickUp URL is a live hyperlink.
- Called `pull_ticam_rates_2026` from `run_build` and threaded `ticam_by_property` through `build_workbook` → mapping loop → `write_property_block` → `write_ticam_row`.
- Updated `extract_ann_edits` to probe `block_start + 4` first (new layout: banner, address, TICAM, spacer, DEAL ACTIVITY) then fall back to `block_start + 3` (legacy pre-TICAM layout). Without this, the first rebuild after deploy would fail to extract Ann's edits from the current SharePoint copy.

**What did not change.**
- Ann's edit preservation (Ann Notes, Deal Activity 2, Broker Reported Rate, Market Rent, ROFR flags, LOI/lease dropdowns), REM ClickUp Comment sync, race-condition guard, DST behavior, nightly/weekly cron schedule — all untouched.
- Only 2026 rates are shown; 2025 and 2027 rows are ignored.
- The TICAM list itself is not written to — read-only.
- The rebuild continues to succeed even if the TICAM list is inaccessible; the row falls back to "not published" everywhere.

**Risk / rollback.**
- Risk: low. Non-fatal fetch (returns `{}` on any error); every write is inside `write_property_block` and only adds one visible row per property.
- Rollback: revert this commit. The `extract_ann_edits` offset probe (+4 then +3) is designed to tolerate reading old SharePoint copies uploaded before this deploy, so the rollback is safe even mid-cycle.

**Verification.**
- Local smoke test with 7 mock tasks + 6 Snap Shot names covers: dropdown-set match (Ashcroft), task-name fallback (Amberwood — the live payload has no dropdown value), 2025 row filtered out, all-zero row omitted, South Memorial dupe-dropdown reconciled to the confirmed row, Stratford Plaza mapping-file reconciles missing paren. All six assertions pass.
- `python3 -m py_compile snap_shot/rebuild.py` — clean.
- After merge, trigger `Snap Shot — nightly rebuild` with `force_rebuild=true` and open the resulting workbook in SharePoint. Expect one TICAM row directly below each property banner. Amberwood Plaza should show `PGP CAM $3.40  ·  Tax $1.03  ·  Ins $0.43` and link to task `868k4u5ab`. Properties without a 2026 task (e.g. Barker Cypress Marketplace at time of writing) should show "2026 TICAM: not published".

## 2026-08-06 — Snap Shot: move nightly + weekly crons off the :00 slot (L2)

- **File(s):** `.github/workflows/snap-shot-nightly.yml`, `.github/workflows/snap-shot-weekly.yml`
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` v1.1 (times updated in-place, no version bump)

**Why.** The nightly workflow's scheduled `0 8 * * *` fire never actually ran — the file has been on main since Aug 4 20:18 UTC, and both expected scheduled fires (Aug 5 and Aug 6 at 08:00 UTC) were dropped by GitHub Actions. Only 3 scheduled events fired across the whole repo in 36 hours, none of which was the nightly. GitHub's documented behavior: on-the-hour crons are heavily contended globally and scheduled events at those slots are the most likely to be throttled or dropped, with no retry. The workbook has been getting rebuilt only when Alexis manually triggers `force_rebuild=true`.

**What changed.**
- `snap-shot-nightly.yml`: cron `'0 8 * * *'` → `'17 8 * * *'` (4:17 AM EDT / 3:17 AM EST).
- `snap-shot-weekly.yml`: cron `'0 9 * * 1'` → `'23 9 * * 1'` (5:23 AM EDT / 4:23 AM EST).
- Added a comment in each YAML file explaining why we intentionally avoid :00.

**What did not change.**
- The on-demand poll workflow (`*/15 11-22 * * 1-5`) is untouched — :00 is only 1 of 4 fire slots per hour there, so throttling matters less.
- No Python or workbook-layout changes.
- Ann's edit-preservation, REM ClickUp Comment sync, race-condition guard, DST behavior — all unchanged.
- Manual `workflow_dispatch` still available on both workflows with the `force_rebuild` override on nightly.

**Risk / rollback.**
- Risk: low. Only shifts fire time by 17 / 23 minutes.
- Rollback: revert this commit; but be aware the previous cron literally never fired, so rolling back reintroduces the bug.

**Verification.**
- Tomorrow (2026-08-07) at 08:17 UTC / 4:17 AM ET, expect a scheduled run to appear in `gh run list --workflow "Snap Shot — nightly rebuild" --repo PGPapattison/Leasing` with `event=schedule`. If it does not fire within 15 minutes of the scheduled time, investigate further (workflow-file freshness, repo activity, or GitHub Actions status).
- Next Monday (2026-08-10) at 09:23 UTC / 5:23 AM ET, expect the weekly rebuild to fire.

---

## 2026-08-06 — Snap Shot: REM ClickUp Comment column + write-back to ClickUp tasks (L3 — Commit 2 of 2)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` v1.0 → v1.1 (pending, same-day update)

**Why.** Commit 1 gave REMs a read-only ClickUp Summary column in the Snap Shot. Alexis then asked for the reverse channel — a place in the workbook where REMs can type a comment that automatically gets posted back to the matching ClickUp task(s), attributed with the REM's name. This closes the loop: REMs read the AI summary in col O and reply directly in col P, no context-switch to ClickUp.

**What changed.**
- Added **column P “REM ClickUp Comment”** (REM-editable) and **column Q “Last Synced”** (read-only receipt) to the Rent Roll section of every property block. `LAST_COL` widened from O → Q. `COL_WIDTHS` set: O=80 (was 44 — Alexis: rows were too tall at the narrower width), P=40, Q=24. REM name banner still merges `L{row}:{LAST_COL}{row}` and now spans L–Q. TOTALS row leaves both new columns blank.
- `pull_lar_summaries()` now returns FOUR indexes instead of two: adds `task_ids_by_tenant_id = {tenant_id: [task_id, ...]}` and `task_ids_by_prop_unit = {(pid, unit): [task_id, ...]}`. Task-id lists are accumulated across all three LAR lists (Renewal + Vacancy + Docs Workflow) so a REM comment for a tenant on multiple lists posts to every matching task (per Alexis: “Post to all matching tasks”).
- New helper **`sync_rem_comments_to_clickup()`** — iterates every extracted REM comment, computes an sha8 hash of whitespace-normalized text, and posts to matching tasks ONLY when the hash differs from the sha8 embedded in the paired Last Synced cell. Comment format: `📋 LSS note from {rem_name} ({YYYY-MM-DD}):\n\n{comment}`. On success, refreshes Last Synced to `"YYYY-MM-DD HH:MM ET · sha8:xxxxxxxx"`. Empty comments and units with no matching task are skipped (no stamp, so a later run can retry).
- `extract_ann_edits()` now returns FIVE values (was three): adds `rem_comments_by_prop_unit = {(pid, unit_label): comment}` and `last_synced_by_prop_unit = {(pid, unit_label): stamp}`. Both are round-tripped through SharePoint so REM edits survive nightly rebuilds and diff-detect is stable across runs.
- `build_workbook()` signature adds `rem_comments_by_prop_unit=None` and `last_synced_by_prop_unit=None`; per-property subsets are filtered to `mapping_entry["_rem_comments_by_unit"]` and `mapping_entry["_last_synced_by_unit"]` (keyed by unit_label alone). `write_property_block()` reads those to populate cols P and Q per unit; primes row height when col P has a long comment.
- `run_build()` wires the pipeline in the correct order: extract → rent-roll pull → LAR pull (with task-ids) → sync REM comments → build → upload. Sync is non-fatal.
- Added `import hashlib`.

**What did not change.**
- Column N (“Notes”) is still Ann’s per-unit note field and still round-trips through `unit_notes.json` — nothing about her workflow changes. Only the new column P posts to ClickUp.
- Ann-authored deal-activity notes, Property Flags, Broker Calls, Ann's Commentary, Market Rent overrides, Broker Active Interest (weekly-only refresh) all still preserved via `extract_ann_edits`.
- REM names in `property_mapping_all73.json` are unchanged (still baked in from Commit 1 — 30 Parker, 24 Leah, 18 Kerri).
- ClickUp Summary (col O) is still read-only from ClickUp → Excel. Nothing round-trips from O back to ClickUp.
- The race-condition guard (skip when Ann edited within 15 minutes) is unchanged.
- Dry-run mode is preserved: `sync_rem_comments_to_clickup` receives `dry_run=True` and logs targets without posting.

**Risk / rollback.**
- Risk: medium. New write path (this is our first automation that posts REM-authored content back into ClickUp task comments). Diff-detect is hash-based, so unless a REM manually strips the sha8 from col Q, the same comment cannot double-post.
- Rollback: `git revert <this-sha>`. Column P/Q content already in SharePoint will simply be ignored by the reverted code (extract_ann_edits will return three values again). No cleanup needed on ClickUp — comments already posted stay as historical record.

**Verification.**
- Local smoke test: `python3 /tmp/local_test2.py` builds a two-unit fixture, confirms cols O/P/Q populate, hash helpers round-trip, REM banner spans L:Q.
- Production: trigger `gh workflow run "Snap Shot — nightly rebuild" --repo PGPapattison/Leasing --ref main -f force_rebuild=true` and verify logs show `LAR pull totals: task-ids — N tenants → tasks, M (prop,unit) → tasks` and `REM comment sync: X unit(s) synced (Y ClickUp comment(s) posted), Z unchanged skipped, ...`.
- First real REM comment: after this ships, watch the next rebuild for the first sync event and spot-check the posted comment on the target ClickUp task.

---

## 2026-08-06 — Snap Shot: add ClickUp Summary column + REM name banner (L3 — Commit 1 of 2)

- **File(s):** `snap_shot/rebuild.py`, `snap_shot/data/property_mapping_all73.json`
- **Author:** Alexis Pattison
- **Skill:** `prudent-snap-shot-rebuild` (pending v1.1 update in a later commit — skill still describes prior schema until Commit 2 lands)

**Why.** Alexis: "It would be ideal if we could add a ClickUp Summary column to the Rent Roll that could pull the summary from ClickUp and paste it there each day and also have a column that the REM's could make a note on and then that note be transfered into ClickUp each night when the update runs. In order for it to show the proper name of the person posting the comment, wed probably also need to add the name of the Real Estate Manager to each property in the LSS." Then follow-up: "Could the REM name go right justified on the same row the Property ID and Address is on? Also, the REM Names wont change and in the odd case they did, we can manually change it on that row so im not sure you need to pull that (#2 above) each night. Could we do Commit 1 now and then Commit 2 later today?"

This is Commit 1 of 2: the read-only side (ClickUp → Excel) plus the REM name banner. Commit 2 (later today) adds the write-side REM Note column that round-trips comments back to ClickUp with REM name attribution.

**What changed.**
- **New Rent Roll column O ("ClickUp Summary", width 44)** on every property block. Read-only from ClickUp. Populated nightly from the three LAR lists (Renewal Pipeline `901113575567`, Vacancy Pipeline `901113575628`, Documents Workflow `901113991446`).
- **New helper `pull_lar_summaries()`** in `rebuild.py`. Pulls every task from all three LAR lists, reads each task's Summary custom field (Renewal + Vacancy share field ID `93833d96-…`; Documents Workflow uses `190e6156-…`), and returns two indexes:
  - `by_tenant_id`: `{TenantId → summary_text}` used for occupied units. First-match-wins in list order Renewal > Vacancy > Documents Workflow.
  - `by_prop_unit`: `{(PropertyId, Unit#) → summary_text}` used for vacant units (no TenantId). Vacancy Pipeline only.
- **Match logic in `write_property_block` per-unit loop:** try Tenant ID first, fall back to (Property ID, Unit #). No match → empty cell.
- **New REM name banner** on the address row of every property block. Same row as "Property ID: NNN · Address", right-justified in a new L–O merged region: `"Real Estate Manager: {name}"`. Falls back to italic em-dash if REM is unset. Same BLUE_MID fill as the address so the split is visually invisible.
- **REM names baked into `property_mapping_all73.json`.** 72/72 rows now carry `rem_name`. Populated by joining the Property master list (`901112873404`, field "🏠 Real Estate Manager 2" `e0d44d11-…`) on Property ID (`057285dd-…`). Distribution: Parker Owen 30, Leah Sykes 24, Kerri Blevins 18. Per Alexis, this is a one-time bake — REM assignments don't change nightly; on the rare change they'll be edited manually in the mapping file.
- **`build_workbook()` signature extended** with two new optional kwargs (`clickup_summaries_by_tenant`, `clickup_summaries_by_prop_unit`), threaded down to `write_property_block` via `mapping_entry["_clickup_summaries_by_tenant/_by_prop_unit"]` (same pattern already used for `_market_rent_overrides`, `_unit_notes`, etc.).
- **`run_build()`** calls `pull_lar_summaries()` before `build_workbook()` in every mode (nightly, weekly, dry-run). Non-fatal on failure: individual list fetch errors log warnings and skip; a total failure logs a warning and continues with empty summaries rather than crashing the whole build.
- **Cell styling for the new column:** 8pt italic slate on very-light-gray fill (`F5F6F8`), wrap-text on. Visually distinguishes it as "informational, not for editing."
- **Row height priming:** if the summary is >60 chars, pre-set the row height to match wrapped-line count (up to Excel's 409pt limit). Only raises height, never lowers — protects any existing height priming from the Notes column.
- **Totals row** in `write_property_block` now blanks column O along with the existing Notes blank.
- **`LAST_COL`** changed from `"N"` to `"O"`. Every full-width merge (`B{row}:{LAST_COL}{row}`) automatically widens, including the top LEASING SNAP SHOT title bar, KPI tiles, and per-property-name banner.
- **KPI tile K column** now spans K–O (5 columns wide vs 4 previously) since `tile_ranges` uses `LAST_COL`. Accepted as-is — the fourth tile is slightly wider than the first three. Symmetric rework is a future polish item.

**What did not change.**
- **Ann's edit extraction (`extract_ann_edits`)** is untouched. It reads columns N and earlier for property overrides / unit notes / market rent overrides. Column O is never read on the extraction path, so ClickUp Summary values in the SharePoint file are always overwritten from the live LAR pull on every rebuild — not preserved as "user edits."
- **All existing Ann-authored fields** (Deal Activity, Property Flags, Vacant Callouts, Ann's Commentary, Broker Calls, Market Rent overrides, per-unit Notes) still round-trip through SharePoint and are preserved on every rebuild.
- **Broker Beat / Broker Active Interest** weekly refresh unchanged.
- **Notes column (N)** unchanged — still REM-editable, still wrap-text, still round-tripped.
- **Scheduling / cron / GitHub Actions workflows** unchanged.
- **Change-detection SharePoint skip window** (15-min editor bypass) unchanged.
- **ClickUp write paths** — NO writes to ClickUp in Commit 1. `pull_lar_summaries` is strictly read-only. Comment posting comes in Commit 2.

**Risk / rollback.**
- Risk: **medium**. Adds a new nightly ClickUp read across 3 lists (~500+ tasks). If ClickUp is down or rate-limits, individual list pulls log a warning and skip — the workbook still builds, just with empty ClickUp Summary cells. Widening `LAST_COL` from N to O affects every full-width merge in the workbook; the KPI tiles now show a slightly wider last tile. Print layout still fits landscape letter (`fitToWidth=1`).
- Rollback: revert this commit. The mapping file's `rem_name` field will remain but be unused — no cleanup needed.

**Verification.**
1. Trigger a nightly run with `force_rebuild=true`. In the logs, look for `LAR pull totals:` — should show non-zero counts for both indexes.
2. Open the resulting workbook. On any property block:
   - Row 2 of the block should show the address on the left and "Real Estate Manager: {name}" right-justified.
   - Rent Roll should have a new far-right column labeled "ClickUp Summary" (col O).
   - Occupied units with active LAR tasks should show their AI summary text; unmatched units are blank.
3. Confirm Ann's edits (Deal Activity, Notes) are preserved on the next round-trip — pick one property, add a fake note, run the rebuild, confirm the note persists.
4. Confirm `extract_ann_edits` still reports the same counts it did before the change (no false-positive column-O reads).

**Follow-up (Commit 2, later today).**
- Add REM-editable "REM Note" column (col P) that round-trips into ClickUp as a task comment with `📋 LSS note from {REM name} ({date}): {text}` attribution.
- Hidden state sheet for "Last Synced Note" per unit, to diff-detect changes so we only post when the note actually changed.
- Update `prudent-snap-shot-rebuild` skill to v1.1 with the new schema, at that time.
- Team email announcing the visible new column to REMs (Leah, Kerri, Parker, Alexis) — defer until Commit 2 ships so the whole feature is announced at once.

## 2026-08-05 — List description: overwrite instead of splice (L1)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison

**Why.** Alexis: "when you pin the last update in the description of the lease, can you delete the last post and just copy over so we dont end up having a long string on the list description?" — the splice-and-preserve behavior was risking a growing list description over time (each run kept the below-divider content and layered the block on top). She wants the list description to always be just the current refresh link — nothing else.

**What changed.**
- `update_broker_reporting_list_description()` now overwrites the entire list description with just the refresh block (`📄 **Latest Snap Shot workbook:** [Open in Excel Online](...) / _Last refreshed ..._`). No divider, no preserved history.
- Task description behavior unchanged — it still splices into the existing description because the control task has real workflow docs below the auto-managed block.

**What did not change.**
- Control task description: still uses splice logic (preserves workflow docs).
- Refresh block format itself.
- Link source (Graph webUrl → Excel Online).

**Risk / rollback.**
- Risk: low. If someone manually adds content to the list description, the next rebuild wipes it. That's the requested behavior. If she ever wants list-level docs, they should live on the pinned control task instead.
- Rollback: revert this commit.

**Verification.**
- Fire nightly with `force_rebuild=true`. Confirm list description shows ONLY the refresh block, no prior content or dividers.

## 2026-08-05 — Right-size note rows (drop 4-line floor + wrap safety margin) (L1)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison

**Why.** Alexis flagged that Deal Activity & Notes rows on Barker Cypress and Castle Shops looked too tall relative to their content — wasted vertical space. Inspection confirmed rows with 1–2 lines of content were rendering at 63–78pt when 33–48pt would suffice.

**Root causes:**
1. `label_min_lines = {Deal Activity: 4, Ann's Commentary: 4, BAI: 2}` was forcing populated rows to be at least 4 (or 2) lines tall regardless of actual content length.
2. `safety_margin=1` in `_row_height_for` was adding 1 extra line to every populated row on top of the +15pt headroom — double-buffering.

**What changed.**
- Removed the `label_min_lines` floors entirely. Populated rows now size purely to their wrapped-line count. Empty rows still collapse to 18pt via the `has_real_content` check.
- Pass `safety_margin=0` to `_row_height_for` for note rows. The +15pt headroom already covers ~1 line of live-typing edge cases.
- Label wrap floor still respected (a long label like "Broker Active Interest from Weekly Reporting" still gets the ~3 lines it needs to display).

**Height impact (Barker Cypress + Castle Shops before → after):**
- Deal Activity `Ultra Nails\nPAC Payments`: 78pt → 48pt (–38%)
- Property Flags `**Homeless\n**Roof`: 63pt → 48pt
- Deal Activity `Different Touch — 4/30/2027 — Expected to Renew...`: 78pt → 48pt
- Broker Calls `Bi-Weekly Calls: 7/23/2026`: 48pt → 33pt
- Empty Deal Activity/Ann's Commentary: 18pt (unchanged — still collapse)
- Populated BAI ~120 chars with 3 items: 153pt → 93pt

**What did not change.**
- Empty rows still collapse to 18pt.
- Long populated rows still size proportional to content.
- +15pt live-typing headroom still applied to any populated row.
- Per-unit Notes rows still Excel-auto-fit natively.

**Risk / rollback.**
- Risk: low. Height comes from actual content length + headroom. If a row's content happens to sit exactly on a wrap boundary, +15pt buffer catches it.
- Rollback: revert this commit.

**Verification.**
- Fire nightly with `force_rebuild=true`. Confirm Barker Cypress and Castle Shops Deal Activity rows are visibly proportional to their content.

## 2026-08-05 — Clean ClickUp description format, use Excel Online URL, trim row headroom (L2)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison

**Why.** Three follow-ups after seeing the description live in ClickUp:
1. HTML-comment markers (`<!-- SNAP_SHOT_LAST_REFRESH_BEGIN -->`) were rendering as literal text — ClickUp doesn't strip HTML comments.
2. The link was a direct file URL that triggered desktop Excel; Alexis wants it to open in Excel Online in the browser.
3. 40% row-height padding on note rows was too aggressive — rows looked visibly puffy.

**What changed.**
- Removed HTML-comment markers entirely. New format detects the auto-managed block by its content signature (`📄 **Latest Snap Shot workbook:**` header + `_Last refreshed ..._` italic footer).
- Cleanup regex sweeps legacy HTML-comment markers AND the bold-italic "auto-managed by Snap Shot rebuild" markers from any previous version on the first run.
- New rendered format:
  ```
  📄 **Latest Snap Shot workbook:** [Open in Excel Online](…)

  _Last refreshed 2026-08-05 12:15 PM ET · nightly rebuild_
  ```
- Link source: use Graph API's `webUrl` from the upload response instead of the constructed `Shared Documents/...` URL. Graph returns the `_layouts/15/Doc.aspx?sourcedoc={GUID}&action=default` URL which opens in Excel Online.
- Row headroom: dropped from `row_h * 1.4` (min +30pt) to a fixed `+15pt` for note rows with content. Enough for a small mid-day addition without visible puffiness.

**What did not change.**
- Location of the link (list description at top + control task description).
- Per-unit Notes auto-fit — those cells are unmerged and Excel handles auto-fit natively.
- Rest of description content — the workflow docs below the auto-managed section stay intact.

**Risk / rollback.**
- Risk: low. Splice tested against 4 scenarios (empty description, legacy HTML-comment format, current format, plain-text original). All round-trip cleanly.
- Rollback: revert this commit.

**Verification.**
- Fire nightly with `force_rebuild=true`. Confirm the ClickUp description shows clean formatted markdown (no visible marker text) and the link opens in Excel Online.

## 2026-08-05 — Add live-typing headroom to Deal Activity / Ann's Commentary / BAI rows (L1)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison

**Why.** After collapsing empty note rows, Alexis asked whether populated rows would auto-grow if more text is typed between rebuilds. Answer: no — those cells are merged (col C through last col) and Excel's auto-fit doesn't work on merged cells. Solution: over-provision the height whenever content is present so mid-day additions still fit without clipping. Next nightly recalculates exactly.

**What changed.**
- After the estimator computes `row_h` for a note row with real content, multiply by 1.4 (with a floor of +30pt / ~2 extra lines, capped at Excel's 409pt max). Applied to Deal Activity, Ann's Commentary, Broker Active Interest, Property Flags, Broker Calls.
- Empty rows unaffected — they still collapse to 18pt.

**What did not change.**
- Estimator logic itself (chars_per_line, line_height, semicolon-line counting).
- Per-unit Notes rows — those are unmerged and Excel auto-fits them live.
- Any other formatting.

**Risk / rollback.**
- Risk: low. Rows are visibly taller when populated. Cap at 409pt prevents any absurd growth.
- Rollback: revert this commit.

**Verification.**
- Fire nightly with `force_rebuild=true`. Confirm populated Deal Activity / Ann's Commentary rows have visible whitespace below the text.

## 2026-08-05 — Collapse empty Deal Activity/Ann's Commentary rows + write link to LIST description too (L2)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison

**Why.** Alexis flagged two follow-ups after the initial ClickUp-link + row-height changes:
1. Empty Deal Activity and Ann's Commentary rows were still ~68pt tall (min_lines=4), which looked wasteful when no content was in them.
2. The SharePoint link was landing on the pinned CONTROL TASK's description, not on the LIST description at the top of the Broker Reporting list (where the whole team sees it first).

**What changed.**
- Row-height calc: when the value cell is empty/whitespace (or just the BAI placeholder text), skip the `label_min_lines` floor for that row. Empty rows now collapse to just what the label needs (~17pt). Rows with real content still use `min_lines=4` for Deal Activity / Ann's Commentary so wrap fidelity is preserved when content is present.
- Added `cu_get_list(list_id)` and `cu_update_list_description(list_id, markdown_content)` helpers wrapping `GET/PUT /list/{id}`.
- Added `update_broker_reporting_list_description(sharepoint_url, mode)` that rewrites a marked-off "Latest Snap Shot" section at the top of the Broker Reporting list description (list `901114227189`), same marker-block pattern as the task-description updater.
- Wired into `run_build` as Step 7 (LIST update) with the existing task-description update kept as Step 7b for redundancy. Both non-fatal.

**What did not change.**
- Rows with real Deal Activity or Ann's Commentary content still wrap the same as before.
- Rest of the list description / task description — workflow docs stay verbatim below the auto-managed section.
- SharePoint folder path, filename, upload logic.

**Risk / rollback.**
- Risk: low. Both updates are non-fatal on failure. Row-height change only affects empty note rows.
- Rollback: revert this commit.

**Verification.**
- Fire nightly with `force_rebuild=true`. Confirm empty Deal Activity/Ann's Commentary rows are much shorter, and both the list description AND control task description have a "Latest Snap Shot" section at the top.

## 2026-08-05 — Stamp ClickUp control task with SharePoint link on each rebuild (L2)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison

**Why.** Alexis requested that after every successful rebuild, the SharePoint link to the fresh workbook be pasted into the description of the pinned ClickUp control task (`868km8qph`, "SNAP SHOT REBUILD — check to refresh"). This gives the team a one-click link to the latest copy without navigating SharePoint.

**What changed.**
- Added `WORKBOOK_WEB_URL` constant built from `SHAREPOINT_FOLDER_PATH` + `WORKBOOK_FILENAME` (URL-encoded per RFC 3986).
- Added `cu_update_task_description(task_id, markdown_description)` helper wrapping `PUT /task/{id}` with `markdown_content` payload.
- Added `update_control_task_with_refresh_link(sharepoint_url, mode)` helper that:
  - Finds the control task via `find_control_task()`.
  - Fetches the current markdown description.
  - Rewrites (or prepends on first run) a marked-off "Latest Snap Shot" section between `<!-- SNAP_SHOT_LAST_REFRESH_BEGIN -->` and `<!-- SNAP_SHOT_LAST_REFRESH_END -->` markers. Preserves the rest of the description (workflow docs, troubleshooting, etc.) verbatim.
  - Includes: SharePoint link, refresh timestamp, and rebuild mode (nightly/weekly/poll).
- Wired into `run_build` as "Step 7" after the SharePoint upload. Wrapped in try/except so a ClickUp description update failure never fails the rebuild itself (workbook is already live at that point).

**What did not change.**
- SharePoint folder path, filename, upload logic.
- The rest of the control task description — workflow docs, troubleshooting steps, etc.
- Any other note-row / column / formatting behavior.

**Risk / rollback.**
- Risk: low. Description update is non-fatal; failures are logged as warnings and the rebuild still succeeds. On first run against a description that doesn't have the markers, the new section is prepended (not replacing anything).
- Rollback: revert this commit. The existing description will regain any content between markers on the next rebuild if the markers still exist, or you can manually edit them out.

**Verification.**
- Fire the nightly. Confirm the control task description now has a "Latest Snap Shot" section at the top with a clickable link, and the rest of the docs are intact below.

## 2026-08-05 — Unfreeze top rows + let per-unit Notes rows auto-grow live (L2)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison

**Why.** Two workflow improvements Alexis requested after using the workbook:
1. The frozen panes at `A12` locked the top 11 rows visible while scrolling, which limited how much of a property block Ann and the team could see at once. There's no meaningful benefit to the freeze because each property block is already visually bounded by its own header.
2. When Ann or the team types into a per-unit Notes cell, the row height was NOT growing. It was fixed at 17pt for empty notes and capped at 120pt for prefilled notes. Anything longer got clipped until the user manually resized. She wants the row to auto-grow as she types.

**What changed.**
- **Freeze panes disabled.** `ws.freeze_panes = "A12"` commented out. Scrolling now shows the full workbook without a locked top area.
- **Per-unit Notes rows now auto-grow.** Root cause: openpyxl was writing an explicit `row_dimensions[row].height` for every unit row (17pt empty, 17-120pt for prefilled notes). Excel treats any explicit height as a hard lock and disables auto-fit — so when the user typed a long note, the row would clip instead of growing.
  - New behavior: only set an explicit row height when the incoming prefilled note is already >30 chars (so the initial view isn't clipped). Empty notes and short notes get NO explicit height, so Excel auto-fits on open AND continues to auto-fit as the user types.
  - Removed the 120pt cap. Row height can now grow up to Excel's 409pt max.
  - Notes cell already had `wrap_text=True` + `vertical="top"`, which is what triggers Excel's auto-fit — no cell-level changes needed.

**What did not change.**
- Property-block note rows (Broker Calls, Deal Activity, Property Flags, Vacant Callouts, Ann's Commentary, Broker Active Interest) still use the pre-computed estimator because those cells are MERGED (C..LAST_COL) and Excel auto-fit is unreliable on merged cells.
- All other formatting: colors, borders, column widths, font, logo, etc.
- Ann-authored edit preservation still works: `extract_ann_edits` still pulls the current SharePoint file's per-unit notes and rewrites them into the new build.

**Risk / rollback.**
- Risk: low. If Excel auto-fit behaves unexpectedly on any user's setup, we can restore the explicit heights by reverting the `if unit_note and len(unit_note) > 30:` block.
- Rollback: revert this commit.

**Verification.**
- Open the workbook after the next nightly. Confirm (a) the top rows are no longer frozen, (b) type a long note into any per-unit Notes cell and confirm the row grows to fit as you type.
- REM / VA impact: none. This is a workflow improvement for editors, transparent to consumers.

## 2026-08-05 — Make every note row expand to fit content, never clip (L1)

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison

**Why.** First pass at fixing Broker Active Interest clipping (chars_per_line 145->85, max_height 110->280) fixed the reported case but Alexis flagged the general problem: if Property Flags, Vacant Callouts, Deal Activity, or any other note grows long, it should also auto-expand to show all text. Since the row-height estimator is shared across ALL note rows, the safer fix is to make the estimator lean generous — users would rather see a slightly tall row than a clipped one.

**What changed.**
- `chars_per_line` 85 -> 72 (deliberately narrower than measured width, so borderline wraps have headroom).
- `line_height` 14pt -> 15pt (matches 9pt Calibri single-line height with breathing room, avoids descender-clipping seen on a couple of rows).
- Added `safety_margin=1` extra line for any row with visible content: if Excel actually wraps to N+1 lines instead of N, no clipping.
- `max_height` 280pt -> 409pt (Excel's absolute per-row max). Only used by unusually long notes.
- Added extensive inline comments so future authors understand the estimator's rationale.

**What did not change.**
- The estimator is still applied to every note row via the same `for label, val in note_labels:` loop, so this fix universally covers Broker/Contact, Broker Calls, Property Flags, Vacant Callouts, Deal Activity, Broker Active Interest, and Ann's Commentary.
- `wrap_text=True` and `vertical=top` still set on every value cell.
- No column widths, layouts, or other formatting.

**Risk / rollback.**
- Risk: rows will average taller than before. Empty rows still get the 18pt floor; single-line rows get roughly the same height. Only rows with real content grow.
- Rollback: revert this commit + the earlier 145->85 commit to restore the original 145 / 14 / 110 estimator.

**Verification.**
- Fire nightly. Open the workbook and spot-check property blocks that were clipping earlier (Broker Active Interest on any active leasing property; Property Flags on a property with a long flag list). All text should now be visible.

## 2026-08-05 — Fix Broker Active Interest row height clipping content (L1) [SUPERSEDED]

- **File(s):** `snap_shot/rebuild.py`
- **Author:** Alexis Pattison

**Why.** Alexis reported that on every property block the 'Broker Active Interest from Weekly Reporting' row was cutting off text mid-line despite `wrap_text=True` being set on the cell. Root cause: the row-height estimator `_row_height_for` assumed 145 characters per visible line and capped max height at 110pt. Measured against actual column width (merged C..LAST_COL at 9pt Calibri), the true chars-per-line is closer to 85, and Broker Active Interest content can easily exceed the 110pt cap.

**What changed.**
- `_row_height_for` now uses `chars_per_line=85` (was 145) and `max_height=280pt` (was 110pt).
- Estimator now respects explicit `\n` line breaks in source text (previously ignored, so paragraphs with real newlines under-counted lines).
- Added `"Broker Active Interest from Weekly Reporting": 2` to `label_min_lines` so the row is always at least 2 lines tall.

**What did not change.**
- `wrap_text=True` was already set on the value cell — no alignment or wrap changes needed.
- Other note rows (Deal Activity, Ann's Commentary, etc.) also benefit from the tighter estimate but keep their existing `min_lines` floors.
- No changes to workbook layout, columns, or any other formatting.

**Risk / rollback.** [SUPERSEDED by the next commit — the 85/280 values were still too tight for edge cases; see above for the final values.]

**Verification.** See above.

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
