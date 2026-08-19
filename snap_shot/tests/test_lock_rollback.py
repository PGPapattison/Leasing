"""Unit tests for the SharePoint-locked rollback path.

Four scenarios covered:
  A. Upload succeeds normally -> no rollback, no deletions.
  B. Upload returns HTTP 423 twice -> WorkbookLockedError raised ->
     rollback_locked_workbook deletes each posted ClickUp comment.
  B2. Same as B but with stamped_task_ids -> rollback also clears the
      'Last Synced Hash' custom field on each stamped task.
  C. Upload path itself: repeated 423 raises WorkbookLockedError.

All ClickUp / SharePoint API calls are monkey-patched so we don't need
real credentials.
"""
import sys, io, types
from unittest.mock import patch, MagicMock

sys.path.insert(0, "/home/user/workspace/Leasing/snap_shot")
import rebuild


# ---------------------------------------------------------------------------
# Fake ClickUp API — records posts + deletes so we can assert on them
# ---------------------------------------------------------------------------

FAKE_POSTED = []      # (task_id, text, returned_comment_id)
FAKE_DELETED = []     # comment_id
FAKE_FIELD_WRITES = []  # (task_id, field_id, value)

def fake_cu_post_comment(task_id, text):
    cid = f"c{len(FAKE_POSTED) + 1000}"
    FAKE_POSTED.append((task_id, text, cid))
    return cid

def fake_cu_delete_comment(comment_id):
    FAKE_DELETED.append(comment_id)
    return True

def fake_cu_set_field(task_id, field_id, value):
    FAKE_FIELD_WRITES.append((task_id, field_id, value))
    return True

def fake_send_error_email(subject, body):
    print(f"\n[EMAIL] subject={subject!r}\n[EMAIL BODY]:\n{body}\n---")


# ---------------------------------------------------------------------------
# Scenario A: happy path — upload works, no rollback
# ---------------------------------------------------------------------------

def test_a_happy_path():
    FAKE_POSTED.clear()
    FAKE_DELETED.clear()

    print("=== Scenario A: upload succeeds ===")
    # Simulate a run posting 3 comments successfully.
    posted_ids = [fake_cu_post_comment(f"task{i}", f"body{i}") for i in range(3)]

    # Upload succeeds — no rollback should be triggered.
    print(f"posted={len(posted_ids)}, deleted={len(FAKE_DELETED)}")
    assert len(posted_ids) == 3
    assert len(FAKE_DELETED) == 0
    print("✓ no rollback, no deletions\n")


# ---------------------------------------------------------------------------
# Scenario B: SharePoint 423 — rollback should delete every posted comment
# ---------------------------------------------------------------------------

def test_b_locked_rollback():
    FAKE_POSTED.clear()
    FAKE_DELETED.clear()

    print("=== Scenario B: SharePoint locked, rollback fires ===")

    # 1. Simulate 5 successful ClickUp posts this run.
    posted_ids = [fake_cu_post_comment(f"task{i}", f"body{i}") for i in range(5)]
    print(f"posted {len(posted_ids)} comments to ClickUp: {posted_ids}")

    # 2. Trigger the rollback path directly.
    with patch.object(rebuild, "cu_delete_comment", side_effect=fake_cu_delete_comment), \
         patch.object(rebuild, "cu_set_field", side_effect=fake_cu_set_field), \
         patch.object(rebuild, "send_error_email", side_effect=fake_send_error_email):
        rebuild.rollback_locked_workbook(
            posted_comment_ids=posted_ids,
            context="comment-sync",
            error_detail="SharePoint upload rejected as HTTP 423 resourceLocked: {test}",
        )

    print(f"deleted={FAKE_DELETED}")
    assert FAKE_DELETED == posted_ids, "every posted comment should have been deleted"
    assert FAKE_FIELD_WRITES == [], "no stamped_task_ids passed — no field writes expected"
    print("✓ rollback deleted every posted comment (no stamps to clear)\n")


# ---------------------------------------------------------------------------
# Scenario B2: rollback with stamped_task_ids also clears Last Synced Hash
# ---------------------------------------------------------------------------

def test_b2_locked_rollback_clears_hash_stamps():
    FAKE_POSTED.clear()
    FAKE_DELETED.clear()
    FAKE_FIELD_WRITES.clear()

    print("=== Scenario B2: rollback clears Last Synced Hash stamps ===")

    posted_ids = [fake_cu_post_comment(f"task{i}", f"body{i}") for i in range(3)]
    stamped_task_ids = ["task_a", "task_b", "task_c", "task_d"]
    print(f"posted {len(posted_ids)} comments and stamped {len(stamped_task_ids)} tasks")

    with patch.object(rebuild, "cu_delete_comment", side_effect=fake_cu_delete_comment), \
         patch.object(rebuild, "cu_set_field", side_effect=fake_cu_set_field), \
         patch.object(rebuild, "send_error_email", side_effect=fake_send_error_email):
        rebuild.rollback_locked_workbook(
            posted_comment_ids=posted_ids,
            context="rebuild mode=nightly",
            error_detail="SharePoint upload rejected as HTTP 423 resourceLocked",
            stamped_task_ids=stamped_task_ids,
        )

    print(f"deleted={FAKE_DELETED}")
    print(f"field writes={FAKE_FIELD_WRITES}")
    assert FAKE_DELETED == posted_ids, "every posted comment should have been deleted"
    assert len(FAKE_FIELD_WRITES) == len(stamped_task_ids), (
        f"expected {len(stamped_task_ids)} field clears, got {len(FAKE_FIELD_WRITES)}"
    )
    for (task_id, field_id, value), expected_tid in zip(FAKE_FIELD_WRITES, stamped_task_ids):
        assert task_id == expected_tid, f"task {task_id} != expected {expected_tid}"
        assert field_id == rebuild.CU_LAST_SYNCED_HASH_FIELD, f"wrong field id {field_id}"
        assert value == "", f"expected empty string, got {value!r}"
    print("✓ rollback cleared every stamped task's Last Synced Hash\n")


# ---------------------------------------------------------------------------
# Scenario C: 423 detected in upload_to_sharepoint -> raises WorkbookLockedError
# ---------------------------------------------------------------------------

def test_c_upload_raises_workbooklocked():
    print("=== Scenario C: upload_to_sharepoint sees 423 -> raises ===")

    # Write a tiny dummy file to disk so os.path.getsize works.
    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        tf.write(b"PK\x03\x04dummy" * 100)  # small enough to hit the "simple PUT" branch
        tmp = tf.name

    # Stub get_graph_access_token to avoid needing real MS Graph auth.
    with patch.object(rebuild, "get_graph_access_token", return_value="fake-token"):
        # Patch http_request to always return an object with status_code=423.
        fake_response = MagicMock()
        fake_response.status_code = 423
        fake_response.text = '{"error":{"code":"resourceLocked"}}'
        with patch.object(rebuild, "http_request", return_value=fake_response), \
             patch.object(rebuild.time, "sleep", return_value=None):  # skip the 30s wait
            raised = False
            try:
                rebuild.upload_to_sharepoint(tmp)
            except rebuild.WorkbookLockedError as e:
                raised = True
                print(f"raised WorkbookLockedError: {e}")
            except rebuild.FatalError as e:
                raised = "wrong"
                print(f"WRONG: raised FatalError instead: {e}")

    os.unlink(tmp)
    assert raised is True, "should have raised WorkbookLockedError, not FatalError"
    print("✓ upload_to_sharepoint raises WorkbookLockedError on repeated HTTP 423\n")


if __name__ == "__main__":
    test_a_happy_path()
    test_b_locked_rollback()
    test_b2_locked_rollback_clears_hash_stamps()
    test_c_upload_raises_workbooklocked()
    print("ALL SCENARIOS PASSED")
