# -*- coding: utf-8 -*-
"""
Created on Tue Apr 21 13:42:44 2026

@author: LCHIBUIKE
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Apr 21 12:47:11 2026

@author: LCHIBUIKE
"""

import pandas as pd
import glob

# ----------------------------------
# Configuration
# ----------------------------------
OUTPUT_FILE = "promoter_summary.xlsx"

EXCLUDED_STATUS = [
    "TRANSIT",
    "TRÀNSIT",
    "EXHAUSTED",
    "BLOQUEADA",
    "ELIMINATED"
]

all_data = []
files_without_promoter_column = []

# ----------------------------------
# Find Excel files in current folder
# ----------------------------------
files = glob.glob("*sample*.xls*")

if not files:
    raise FileNotFoundError(
        "No Excel files containing 'sample' found in the current directory."
    )

print(f"🔍 Found {len(files)} file(s) to process.\n")

# ----------------------------------
# Process each file
# ----------------------------------
for i, file in enumerate(files, start=1):
    print(f"➡️  [{i}/{len(files)}] Processing file: {file}")

    df = pd.read_excel(file, engine="openpyxl", header=1)

    # --- Detect promoter column ---
    promoter_col = next(
        (col for col in df.columns if "promoter" in col.lower()),
        None
    )

    if promoter_col is None:
        print("   ⚠️  No promoter column found → using 'Unknown'")
        files_without_promoter_column.append(file)

    # --- Normalize Sample status ---
    df["Sample status"] = (
        df["Sample status"]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    # --- Remove excluded statuses ---
    initial_rows = len(df)
    df = df[~df["Sample status"].isin(EXCLUDED_STATUS)]
    removed_rows = initial_rows - len(df)

    print(f"   ↳ Removed {removed_rows} row(s)")

    # --- Build output dataframe ---
    if promoter_col:
        df_out = df[["Collection", promoter_col]].rename(
            columns={promoter_col: "Promoter"}
        )

        # ✅ Keep rows with missing promoter → set to "Unknown"
        df_out["Promoter"] = (
            df_out["Promoter"]
            .astype(str)
            .str.strip()
            .replace({"": "Unknown", "nan": "Unknown"})
            .fillna("Unknown")
        )
    else:
        # Promoter column missing entirely
        df_out = df[["Collection"]].copy()
        df_out["Promoter"] = "Unknown"

    all_data.append(df_out)

# ----------------------------------
# Combine all files
# ----------------------------------
combined_df = pd.concat(all_data, ignore_index=True)

# ----------------------------------
# Aggregate result
# ----------------------------------
final_table = (
    combined_df
    .groupby(["Collection", "Promoter"])
    .size()
    .reset_index(name="Number of samples")
    .sort_values(
        by=["Collection", "Number of samples"],
        ascending=[True, False]
    )
)

# ----------------------------------
# Save output
# ----------------------------------
final_table.to_excel(OUTPUT_FILE, index=False)

# ----------------------------------
# Summary
# ----------------------------------
print("\n✅ Processing complete.")
print(f"💾 Output saved as: {OUTPUT_FILE}")
print(f"📊 Total samples counted: {final_table['Number of samples'].sum()}")

if files_without_promoter_column:
    print("\n⚠️ Files without promoter column (counted as 'Unknown'):")
    for f in files_without_promoter_column:
        print(f"   - {f}")
