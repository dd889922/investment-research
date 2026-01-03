import pandas as pd
import sys
from pathlib import Path

p = Path(__file__).resolve().parent.parent / "data" / "eia" / "figure6_data.xlsx"
print("Excel path:", p, "exists=", p.exists())
try:
    sheets = pd.read_excel(p, sheet_name=None, engine="openpyxl")
    print("Sheets:", list(sheets.keys()))
except Exception as e:
    print("Error reading sheets:", e)
try:
    raw = pd.read_excel(p, sheet_name="Generation", engine="openpyxl", header=None)
    print("Raw shape:", raw.shape)
    print(raw.head(30).to_string())
    raw2 = raw.dropna(axis=1, how="all")
    print("After dropna cols shape:", raw2.shape)
    # find header row
    known = {"year", "coal", "natural gas", "nuclear", "hydro", "non-hydro renewables", "petroleum"}
    header_idx = None
    for i in range(min(50, len(raw2))):
        vals = [str(v).strip().lower() for v in list(raw2.iloc[i].dropna())]
        if any("year" in v for v in vals):
            match_count = sum(1 for v in vals if any(k in v for k in known))
            if match_count >= 3:
                header_idx = i
                break
    print("Detected header_idx:", header_idx)
    if header_idx is not None:
        header_vals = [str(v) for v in list(raw2.iloc[header_idx])]
        df = raw2.iloc[header_idx+1:].copy()
        df.columns = header_vals
        print("Constructed df shape:", df.shape)
        print("Columns sample:", list(df.columns)[:20])
        print(df.head(10).to_string())
    else:
        print("No header row detected in first 50 rows.")
except Exception as e:
    print("Error raw read:", e)
