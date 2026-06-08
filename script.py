"""
Phishing Training Tracker — Automation Script
----------------------------------------------
Usage:
    python phishing_tracker_update.py <source_file.xlsx> <tracking_file.xlsx>

    1st argument = ABI Learning source report  (read only — never modified)
    2nd argument = Phishing Tracking file      (updated in place)

Output:
    - Tracking file is updated in place

Optimised for large datasets (1M+ rows) using pandas for all data work
and openpyxl only for final formatting.
"""

import sys, os, re, copy
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ── Config ────────────────────────────────────────────────────────────────────
NEW_COLUMNS   = ["Training Start Date", "Training End Date", "Transcript Status"]
EMP_ID_COL    = "Global Employee ID"
SOURCE_HDR    = 20          # header row index (0-based) in source file = row 21

# ── Styles ────────────────────────────────────────────────────────────────────
ORANGE_FILL  = PatternFill("solid", fgColor="FFC000")
YELLOW_FILL  = PatternFill("solid", fgColor="FFFF00")
NO_FILL      = PatternFill(fill_type=None)
HEADER_FONT  = Font(bold=True, name="Arial", size=10)
DATA_FONT    = Font(name="Arial", size=10)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_ALIGN   = Alignment(horizontal="left",   vertical="center", wrap_text=False)
THIN         = Side(style="thin")
BORDER       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Columns that get yellow header styling in output — cosmetic only, no logic uses this
YELLOW_COLS  = {"Zone","Country","Global Employee ID","Local Employee ID",
                "Employee Name","Employee Status","Worker Type","Employee Group",
                "Management Level","First Hire Date","Last Hire Date","Position Name",
                "Job Family Group","ABI Entity 2","Macro Entity Level 2 (Zone)",
                "text before","Employee Email","Band 4+",
                "Manager Employee ID Level 01","Manager Name Level 01","BSC"}

# ── Step 1 — Build lookup from source ────────────────────────────────────────
def build_lookup(source_path: str) -> pd.DataFrame:
    """Read source, return DataFrame keyed by Employee ID with 3 output columns."""
    print(f"  Reading: {os.path.basename(source_path)}")
    df = pd.read_excel(source_path, header=SOURCE_HDR,
                       dtype={"Employee ID": str, "Employee Status": str,
                              "Transcript Status": str})
    df["Employee ID"] = df["Employee ID"].astype(str).str.strip()
    df = df[df["Employee ID"].notna() & ~df["Employee ID"].isin(["nan","None",""])]

    # Rank: keep best status per employee
    rank_map  = {"Completed": 0, "In Progress": 1, "Not Started": 2}
    df["_rank"] = df["Transcript Status"].map(rank_map).fillna(99)
    df = df.sort_values("_rank").drop_duplicates(subset="Employee ID", keep="first")

    # Format dates
    df["_start"] = pd.to_datetime(df["Training Start Date"],  errors="coerce")
    df["_end"]   = pd.to_datetime(df["Transcript Completed Date"], errors="coerce")

    def calc_end(row):
        terminated = str(row.get("Employee Status","")).strip().lower() == "terminated"
        has_start  = pd.notna(row["_start"])
        if terminated and has_start:
            return "Terminated"
        return row["_end"].strftime("%d-%b-%Y") if pd.notna(row["_end"]) else ""

    lookup = pd.DataFrame({
        "Employee ID":          df["Employee ID"].values,
        "Training Start Date":  df["_start"].dt.strftime("%d-%b-%Y").fillna("").values,
        "Training End Date":    df.apply(calc_end, axis=1).values,
        "Transcript Status":    df["Transcript Status"].fillna("").values,
    }).set_index("Employee ID")

    print(f"  → {len(lookup):,} unique employee records extracted")
    return lookup

# ── Step 2 — Merge into tracking ─────────────────────────────────────────────
def merge_tracking(tracking_path: str, lookup: pd.DataFrame) -> pd.DataFrame:
    """Load tracking file (All Data or first sheet), drop stale new cols, merge lookup."""
    print(f"  Reading: {os.path.basename(tracking_path)}")

    # Try 'All Data' sheet first, fall back to first sheet
    xl   = pd.ExcelFile(tracking_path)
    sheet = "All Data" if "All Data" in xl.sheet_names else xl.sheet_names[0]
    df   = pd.read_excel(tracking_path, sheet_name=sheet,
                         dtype={EMP_ID_COL: str})

    # Drop any previously added new columns to avoid duplicates
    drop = [c for c in df.columns if c in NEW_COLUMNS or re.match(r".+\.\d+$", c)]
    df   = df.drop(columns=drop, errors="ignore")

    df[EMP_ID_COL] = df[EMP_ID_COL].astype(str).str.strip().str.split(".").str[0]

    # Merge
    df = df.merge(lookup[NEW_COLUMNS], left_on=EMP_ID_COL,
                  right_index=True, how="left")
    for col in NEW_COLUMNS:
        df[col] = df[col].fillna("Not Found")

    matched   = (df["Transcript Status"] != "Not Found").sum()
    unmatched = (df["Transcript Status"] == "Not Found").sum()
    print(f"  → Matched: {matched:,}  |  Not found: {unmatched:,}")
    return df

# ── Main ──────────────────────────────────────────────────────────────────────
def update_tracking(source_path: str, tracking_path: str):
    folder = os.path.dirname(os.path.abspath(tracking_path))

    print("\n[1/4] Building lookup from source file ...")
    lookup = build_lookup(source_path)

    print("\n[2/4] Merging into tracking file ...")
    df = merge_tracking(tracking_path, lookup)

    print("\n[3/4] Saving updated tracking file ...")
    # Overwrite the tracking file
    with pd.ExcelWriter(tracking_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="All Data", index=False)
        ws = writer.sheets["All Data"]
        # Apply header styles
        for col_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font, cell.border, cell.alignment = HEADER_FONT, BORDER, CENTER_ALIGN
            if col_name in set(NEW_COLUMNS):
                cell.fill = ORANGE_FILL
            elif col_name in YELLOW_COLS:
                cell.fill = YELLOW_FILL
            else:
                cell.fill = NO_FILL
        ws.freeze_panes    = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}1"
    print(f"  → Saved: {os.path.basename(tracking_path)}")

    # ── Zone files with year tabs ─────────────────────────────────────────────
    ZONE_COL = "Macro Entity Level 2 (Zone)"
    if ZONE_COL not in df.columns:
        print(f"\n⚠  '{ZONE_COL}' column not found — skipping zone files")
    else:
        print("\n[4/4] Writing zone files ...")
        folder      = os.path.dirname(os.path.abspath(tracking_path))
        new_col_set = set(NEW_COLUMNS)

        df["_year"] = pd.to_datetime(
            df["Training Start Date"], format="mixed", errors="coerce"
        ).dt.year.fillna(0).astype(int).astype(str).replace("0", "No Date")

        all_cols = [c for c in df.columns if c != "_year"]

        for zone in sorted(df[ZONE_COL].dropna().unique()):
            zone_df  = df[df[ZONE_COL] == zone]
            safe     = re.sub(r"[^\w\s]", "", zone).strip()
            safe     = re.sub(r"\s+", "_", safe).title()
            out_path = os.path.join(folder, f"{safe}.xlsx")

            wb = Workbook()
            wb.remove(wb.active)

            years = sorted(
                zone_df["_year"].unique(),
                key=lambda y: (y == "No Date", y)
            )

            for year in years:
                year_df = zone_df[zone_df["_year"] == year][all_cols]
                ws      = wb.create_sheet(title=str(year)[:31])

                for ci, col_name in enumerate(all_cols, start=1):
                    cell = ws.cell(row=1, column=ci, value=col_name)
                    cell.font, cell.border, cell.alignment = HEADER_FONT, BORDER, CENTER_ALIGN
                    if col_name in new_col_set:
                        cell.fill = ORANGE_FILL
                    elif col_name in YELLOW_COLS:
                        cell.fill = YELLOW_FILL
                    else:
                        cell.fill = NO_FILL

                for ri, (_, row) in enumerate(year_df.iterrows(), start=2):
                    for ci, col_name in enumerate(all_cols, start=1):
                        val = row[col_name]
                        if not isinstance(val, str) and pd.isna(val):
                            val = ""
                        cell = ws.cell(row=ri, column=ci, value=val)
                        cell.font, cell.fill, cell.border = DATA_FONT, NO_FILL, BORDER
                        cell.alignment = CENTER_ALIGN if col_name in new_col_set else LEFT_ALIGN

                for ci, col_name in enumerate(all_cols, start=1):
                    max_len = max(
                        len(str(col_name)),
                        year_df[col_name].astype(str).str.len().max() if len(year_df) else 0
                    )
                    ws.column_dimensions[get_column_letter(ci)].width = min(int(max_len) + 4, 40)

                ws.freeze_panes    = "A2"
                ws.auto_filter.ref = f"A1:{get_column_letter(len(all_cols))}1"

            wb.save(out_path)
            print(f"  → {safe}.xlsx  ({len(zone_df)} rows | tabs: {', '.join(years)})")

        df.drop(columns=["_year"], inplace=True)

    print(f"\n✅  Done! → {folder}\n")

# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("\nUsage:")
        print("  python phishing_tracker_update.py <source_file.xlsx> <tracking_file.xlsx>")
        print("\n  1st argument = ABI Learning source report  (read only)")
        print("  2nd argument = Phishing Tracking file      (updated in place)\n")
        sys.exit(1)

    source_path   = sys.argv[1]
    tracking_path = sys.argv[2]

    if not os.path.exists(source_path):
        print(f"\n❌  Source file not found: {source_path}\n")
        sys.exit(1)
    if not os.path.exists(tracking_path):
        print(f"\n❌  Tracking file not found: {tracking_path}\n")
        sys.exit(1)

    try:
        update_tracking(source_path, tracking_path)
    except Exception as e:
        import traceback
        print(f"\n❌  Error: {e}")
        traceback.print_exc()
        sys.exit(1)
