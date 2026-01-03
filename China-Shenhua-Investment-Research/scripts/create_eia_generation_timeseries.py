import os
import re
import requests
import pandas as pd
import matplotlib.pyplot as plt

EIA_FIG6_URL = "https://www.eia.gov/international/content/analysis/countries_long/China/content/analysis/countries_long/China/excel/figure6_data.xlsx"
DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "eia")
ALT_DATA_DIR = "/Users/david.zhou/investment-research/China-Shenhua-Investment-Research/data/eia"
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "..", "01-industry", "images", "china_electricity_generation_timeseries_2014_2023.png")

KNOWN_SERIES = {
    "coal": ["coal"],
    "natural_gas": ["natural gas", "gas"],
    "nuclear": ["nuclear"],
    "hydro": ["hydro", "hydropower", "hydroelectric"],
    "non_hydro_renewables": ["non-hydro renewables", "other renewables", "renewables (non-hydro)", "solar", "wind"],
    "petroleum": ["petroleum", "petroleum-fired", "petroleum and other liquids"],
}


def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def download_excel(url: str, dest_path: str) -> None:
    if os.path.exists(dest_path):
        try:
            with open(dest_path, "rb") as f:
                head = f.read(2)
            if head.startswith(b"PK"):
                return
            else:
                os.remove(dest_path)
        except Exception:
            try:
                os.remove(dest_path)
            except Exception:
                pass
    ensure_dir(os.path.dirname(dest_path))
    headers = {
        "User-Agent": "Mozilla/5.0 (GitHub Copilot)",
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
        "Referer": "https://www.eia.gov/international/content/analysis/countries_long/China/",
    }
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    content_type = resp.headers.get("content-type", "")
    content = resp.content
    # basic validation for xlsx (zip-based): should start with PK
    if ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" not in content_type) and not content.startswith(b"PK"):
        raise ValueError(f"Unexpected content type when downloading Excel: {content_type} status={resp.status_code}")
    with open(dest_path, "wb") as f:
        f.write(content)


def normalize_col(name: str) -> str:
    return re.sub(r"[^a-z]", "", name.strip().lower())


def select_series_columns(df: pd.DataFrame) -> dict:
    if df.shape[1] == 0:
        raise ValueError("No columns detected in DataFrame. Unable to map series.")
    cols = {normalize_col(str(c)): c for c in df.columns}
    series_map = {}
    # find year column by name or by values
    year_col = None
    for c in df.columns:
        n = normalize_col(str(c))
        if n in ("year", "years"):
            year_col = c
            break
    if year_col is None:
        # try by values: numeric in 2000-2035
        for c in df.columns:
            try:
                vals = pd.to_numeric(df[c], errors="coerce")
                if vals.notna().sum() > 0:
                    subset = vals.dropna()
                    if (subset.between(2000, 2035).mean() > 0.6):
                        year_col = c
                        break
            except Exception:
                pass
    # build series mapping
    for key, candidates in KNOWN_SERIES.items():
        for cand in candidates:
            cand_norm = normalize_col(cand)
            if cand_norm in cols:
                series_map[key] = cols[cand_norm]
                break
    if year_col is None:
        # fallback: first column if available
        try:
            year_col = df.columns[0]
        except Exception:
            raise ValueError("Failed to detect a year column and no columns available.")
    return {"year": year_col, **series_map}


def load_figure6_dataframe(xlsx_path: str) -> pd.DataFrame:
    # Try multiple sheets and header rows to locate the data table
    sheets = pd.read_excel(xlsx_path, sheet_name=None, engine="openpyxl")
    for name, sheet in sheets.items():
        for header_row in range(0, 6):
            try:
                df = pd.read_excel(xlsx_path, sheet_name=name, engine="openpyxl", header=header_row)
                # drop all-empty columns
                df = df.dropna(axis=1, how="all")
                if df.shape[1] == 0:
                    continue
                mapping = select_series_columns(df)
                # require at least 3 series present and year found
                present = [k for k in ["coal", "natural_gas", "nuclear", "hydro", "non_hydro_renewables", "petroleum"] if k in mapping]
                if len(present) >= 3 and mapping.get("year") in df.columns:
                    return df
            except Exception:
                continue
    # Fallback: headerless read, try to detect header row by scanning first 15 rows
    for name in sheets.keys():
        try:
            raw = pd.read_excel(xlsx_path, sheet_name=name, engine="openpyxl", header=None)
        except Exception:
            continue
        for header_row in range(0, 15):
            try:
                header_vals = raw.iloc[header_row].astype(str).tolist()
                df2 = raw.iloc[header_row + 1 :].copy()
                df2.columns = header_vals
                df2 = df2.dropna(axis=1, how="all").dropna(how="all")
                if df2.shape[1] == 0:
                    continue
                mapping = select_series_columns(df2)
                present = [k for k in ["coal", "natural_gas", "nuclear", "hydro", "non_hydro_renewables", "petroleum"] if k in mapping]
                if len(present) >= 3 and mapping.get("year") in df2.columns:
                    return df2
            except Exception:
                continue
    # final fallback: read first sheet with default header
    return pd.read_excel(xlsx_path, sheet_name=0, engine="openpyxl")


def build_dataframe_from_generation_sheet(xlsx_path: str) -> pd.DataFrame:
    try:
        raw = pd.read_excel(xlsx_path, sheet_name="Generation", engine="openpyxl", header=None)
    except Exception:
        return pd.DataFrame()
    raw = raw.dropna(axis=1, how="all")
    # detect header row with many year values (2000-2035)
    header_idx = None
    year_cols_idx = []
    years = []
    for i in range(min(30, len(raw))):
        row = raw.iloc[i]
        num_years = []
        idxs = []
        for j, v in enumerate(list(row)):
            try:
                val = float(v)
                if 2000 <= val <= 2035:
                    num_years.append(int(val))
                    idxs.append(j)
            except Exception:
                continue
        if len(num_years) >= 5:
            header_idx = i
            years = num_years
            year_cols_idx = idxs
            break
    if header_idx is None or not years:
        return pd.DataFrame()
    # parse subsequent rows as series; first column is series label
    data_rows = raw.iloc[header_idx + 1 :]
    series_map = {}
    for _, r in data_rows.iterrows():
        label = str(r.iloc[0]).strip().lower()
        if not label or label in ("nan",):
            continue
        # collect values from year columns
        values = []
        for j in year_cols_idx:
            try:
                values.append(float(r.iloc[j]))
            except Exception:
                values.append(float("nan"))
        series_map[label] = values
    if not series_map:
        return pd.DataFrame()
    # build tidy dataframe with 'year' + known series columns
    df_out = pd.DataFrame({"year": years})
    def get_series(lbls):
        for lbl in lbls:
            key = str(lbl).strip().lower()
            if key in series_map:
                return series_map[key]
        return None
    # add core series
    core_series = {
        "coal": ["coal"],
        "natural gas": ["natural gas", "gas"],
        "nuclear": ["nuclear"],
        "hydroelectric": ["hydroelectric", "hydro"],
        "petroleum": ["petroleum"],
        "solar": ["solar"],
        "wind": ["wind"],
        "biomass and waste": ["biomass and waste"],
        "geothermal": ["geothermal"],
    }
    for col_name, candidates in core_series.items():
        vals = get_series(candidates)
        if vals is not None:
            df_out[col_name] = vals
    # compute non-hydro renewables if components present
    nh_components = ["solar", "wind", "biomass and waste", "geothermal"]
    if all(c in df_out.columns for c in nh_components):
        df_out["Non-hydro Renewables"] = sum(df_out[c] for c in nh_components)
    return df_out


def plot_timeseries(df: pd.DataFrame, mapping: dict) -> None:
    plt.figure(figsize=(10, 6))
    year = df[mapping["year"]]
    # professional palette
    colors = {
        "coal": "#2C3E50",
        "natural_gas": "#C0392B",
        "nuclear": "#F1C40F",
        "hydro": "#2980B9",
        "non_hydro_renewables": "#27AE60",
        "petroleum": "#7F8C8D",
    }
    labels = {
        "coal": "Coal",
        "natural_gas": "Natural Gas",
        "nuclear": "Nuclear",
        "hydro": "Hydropower",
        "non_hydro_renewables": "Non-hydro Renewables",
        "petroleum": "Petroleum & Other Liquids",
    }
    plotted_any = False
    for key in ["coal", "natural_gas", "nuclear", "hydro", "non_hydro_renewables", "petroleum"]:
        if key in mapping:
            plt.plot(year, df[mapping[key]], label=labels[key], color=colors[key], linewidth=2)
            plotted_any = True
    plt.title("China Electricity Generation by Source (2014–2023)", pad=14)
    plt.xlabel("Year")
    plt.ylabel("Generation (TWh)")
    plt.grid(True, alpha=0.3, linestyle="--")
    if plotted_any:
        plt.legend(loc="upper left", frameon=False)
    plt.tight_layout()
    # source footer
    plt.figtext(0.5, 0.01,
                "Data source: U.S. EIA – figure6_data.xlsx (retrieved)\n"
                f"{EIA_FIG6_URL}",
                ha="center", fontsize=9, color="#555555")
    ensure_dir(os.path.dirname(OUTPUT_PATH))
    plt.savefig(OUTPUT_PATH, dpi=200)
    plt.close()


def main():
    ensure_dir(DATA_DIR)
    xlsx_path = os.path.join(DATA_DIR, "figure6_data.xlsx")
    # fall back to absolute path if needed
    alt_xlsx_path = os.path.join(ALT_DATA_DIR, "figure6_data.xlsx")
    try:
        download_excel(EIA_FIG6_URL, xlsx_path)
    except Exception as e:
        print("Failed to download EIA Excel:", e)
        print("If network blocks the file, please manually download from:")
        print(EIA_FIG6_URL)
        print(f"and place it at: {xlsx_path}")
    if not os.path.exists(xlsx_path) and os.path.exists(alt_xlsx_path):
        xlsx_path = alt_xlsx_path
        print("Using alternative data path:", xlsx_path)
    # read robustly
    print("Reading:", xlsx_path, "exists=", os.path.exists(xlsx_path))
    df = load_figure6_dataframe(xlsx_path)
    mapping = {}
    try:
        mapping = select_series_columns(df)
    except Exception:
        pass
    present = [k for k in ["coal", "natural_gas", "nuclear", "hydro", "non_hydro_renewables", "petroleum"] if k in mapping]
    if df.shape[1] == 0 or len(present) < 3:
        print("Standard loader failed; attempting structured parse from 'Generation' sheet...")
        df_alt = build_dataframe_from_generation_sheet(xlsx_path)
        if df_alt.shape[1] == 0:
            raise ValueError("Unable to parse EIA figure6 Excel into a usable table.")
        df = df_alt
        mapping = select_series_columns(df)
    plot_timeseries(df, mapping)
    print(f"Saved timeseries chart to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
