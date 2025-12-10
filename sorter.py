import glob
import os
import pandas as pd

# ============= CONFIGURATION =============
INPUT_FOLDER = "output"          # folder containing your result CSVs
OUTPUT_FILE = "your_csv_pattern"
EMAIL_COL = "sender_email"        # adjust if your email column name differs
CSV_GLOB_PATTERN = "your_csv_pattern"
# ========================================

def main():
    paths = glob.glob(os.path.join(INPUT_FOLDER, CSV_GLOB_PATTERN))
    if not paths:
        raise FileNotFoundError(f"No matching CSV files found in {INPUT_FOLDER!r}")

    # Oldest → newest, then reverse so newest first
    paths = sorted(paths)
    paths.reverse()

    print("Processing (newest first):")
    for p in paths:
        print("  ", os.path.basename(p))

    seen_emails = set()
    collected_rows = []

    for path in paths:
        # READ WITH SEMICOLON
        df = pd.read_csv(
            path,
            sep=";",             # matches your existing files
            engine="python",
            on_bad_lines="warn"  # or "skip" if too messy
        )

        if EMAIL_COL not in df.columns:
            raise KeyError(
                f"Column {EMAIL_COL!r} not found in {os.path.basename(path)}. "
                f"Columns: {df.columns.tolist()}"
            )

        df["__source_file"] = os.path.basename(path)  # optional trace

        for _, row in df.iterrows():
            email = row[EMAIL_COL]
            if email not in seen_emails:
                seen_emails.add(email)
                collected_rows.append(row)

    # Build final DataFrame
    final_df = pd.DataFrame(collected_rows)

    if final_df.empty:
        print("No rows collected. Nothing to write.")
        return

    # ---- CUSTOM SORTING BY COLUMNS D AND E ----
    # Columns D and E = 4th and 5th columns (0-based index 3 and 4)
    columns = list(final_df.columns)

    if len(columns) < 5:
        raise ValueError(
            f"Expected at least 5 columns to apply D/E logic, but got {len(columns)}: {columns}"
        )

    col_d = columns[3]  # column D
    col_e = columns[4]  # column E

    # Treat NaN or empty string as "no value"
    has_d = final_df[col_d].notna() & (final_df[col_d].astype(str).str.strip() != "")
    has_e = final_df[col_e].notna() & (final_df[col_e].astype(str).str.strip() != "")

    # Group codes:
    # 0 = no D, no E
    # 1 = D only
    # 2 = E only
    # 3 = D and E
    group_code = (
        (~has_d & ~has_e) * 0 +
        (has_d & ~has_e) * 1 +
        (~has_d & has_e) * 2 +
        (has_d & has_e) * 3
    )

    final_df["__group_order"] = group_code

    # Sort first by this logic, then by email for stability
    final_df = final_df.sort_values(
        by=["__group_order", EMAIL_COL],
        ascending=[True, True]
    ).reset_index(drop=True)

    # Drop helper column
    final_df = final_df.drop(columns=["__group_order"])

    # WRITE WITH SEMICOLON SO COLUMNS DISPLAY CORRECTLY
    final_df.to_csv(OUTPUT_FILE, index=False, sep=";")

    print()
    print(f"Unique emails retained: {len(final_df)}")
    print(f"Final CSV written to: {OUTPUT_FILE}")
    print(f"Sorted by presence of values in columns D ({col_d}) and E ({col_e}).")

if __name__ == "__main__":
    main()
