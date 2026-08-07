"""Dump full property block layout so we can see where DEAL ACTIVITY and RENT ROLL sit."""
import glob
from openpyxl import load_workbook

path = sorted(glob.glob("recovery_out/snap_v02_*.xlsx"))[0]
wb = load_workbook(path, data_only=True)
ws = wb["Snap Shot"]

# Dump the first 60 rows of the sheet - covers the first property block
print(f"Sheet: {ws.max_row} rows")
print("="*90)
for r in range(1, 65):
    b = ws.cell(row=r, column=2).value
    c = ws.cell(row=r, column=3).value
    b_str = str(b)[:60] if b else ""
    c_str = str(c)[:30] if c else ""
    print(f"  row {r:4d} B={b_str!r:65s}  C={c_str!r}")
