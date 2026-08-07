#!/usr/bin/env python3
"""Open the v02 recovery file and print the 4 REM Comment locations
(property banner, unit, tenant, comment text) so Leah knows exactly which
rows to update in the live workbook."""
import glob, sys
from openpyxl import load_workbook

def find_property_banner(ws, comment_row):
    """Walk upward from comment_row to find the nearest merged banner row.
    Property banner cells are typically in column B, bold, and span multiple cells."""
    for r in range(comment_row, max(0, comment_row - 60), -1):
        v = ws.cell(row=r, column=2).value  # col B
        if v and isinstance(v, str) and v.strip():
            s = v.strip()
            # Skip data-row values by checking if it looks like a property name
            # (property banners typically don't have $ signs, numbers only, etc.)
            # More reliable: look for row where col C is empty (banner) vs data row
            c_val = ws.cell(row=r, column=3).value
            d_val = ws.cell(row=r, column=4).value
            # A banner row usually has B populated but C and D empty (or C has address)
            if not d_val:  # No TenantID -> likely banner or address row
                # Check row above too — if it's also non-empty and D is empty, it's the property title
                # For simplicity, return first B-populated row where D is empty
                if not c_val or (isinstance(c_val, str) and any(x in c_val.lower() for x in ["dr", "st", "rd", "ave", "blvd", "hwy", "ln", "ct", "way", "pkwy"])):
                    # This might be an address row, walk up one more
                    for r2 in range(r - 1, max(0, r - 5), -1):
                        v2 = ws.cell(row=r2, column=2).value
                        d2 = ws.cell(row=r2, column=4).value
                        if v2 and isinstance(v2, str) and v2.strip() and not d2:
                            return v2.strip(), r2
                return s, r
    return "(unknown)", 0

def main():
    files = sorted(glob.glob("recovery_out/snap_v02_*.xlsx"))
    if not files:
        print("[fatal] no v02 file")
        sys.exit(1)
    path = files[0]
    print(f"Opening {path}\n")
    wb = load_workbook(path, data_only=True)
    ws = wb["Snap Shot"]
    print(f"Sheet: Snap Shot ({ws.max_row} rows)\n")
    print("=" * 70)
    count = 0
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=16).value  # col P
        if v and isinstance(v, str) and v.strip() and v.strip().upper() != "REM CLICKUP COMMENT":
            count += 1
            unit = ws.cell(row=r, column=2).value or ""
            tenant = ws.cell(row=r, column=3).value or ""
            tid = ws.cell(row=r, column=4).value or ""
            prop, banner_r = find_property_banner(ws, r)
            print(f"\nREM Comment #{count} (row {r})")
            print(f"  Property : {prop}  (banner row {banner_r})")
            print(f"  Unit     : {unit}")
            print(f"  Tenant   : {tenant}")
            print(f"  TenantID : {tid}")
            print(f"  Comment  : {v.strip()!r}")
    print("\n" + "=" * 70)
    print(f"Total: {count} REM Comments")

if __name__ == "__main__":
    main()
