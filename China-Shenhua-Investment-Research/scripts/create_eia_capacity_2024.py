import os
import re
import requests
import pandas as pd
import matplotlib.pyplot as plt

EIA_FIG7_URL = "https://www.eia.gov/international/content/analysis/countries_long/China/content/analysis/countries_long/China/excel/figure7_data.xlsx"
DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "eia")
ALT_DATA_DIR = "/Users/david.zhou/investment-research/China-Shenhua-Investment-Research/data/eia"
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "..", "01-industry", "images", "china_installed_generation_capacity_2024.png")

KNOWN_CAPACITY = {
    "coal": ["coal"],
    "natural_gas": ["natural gas", "gas"],
    "oil": ["oil", "petroleum"],
    "nuclear": ["nuclear"],
    "hydro": ["hydro", "hydropower", "hydroelectric"],
    "solar": ["solar"],
    "wind": ["wind"],
    "storage": ["storage", "pumped-storage", "battery"],
    "other": ["other"],
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
    if ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" not in content_type) and not content.startswith(b"PK"):
        raise ValueError(f"Unexpected content type when downloading Excel: {content_type} status={resp.status_code}")
    with open(dest_path, "wb") as f:
        f.write(content)


def normalize_col(name: str) -> str:
    return re.sub(r"[^a-z]", "", name.strip().lower())


def extract_2024_capacity(df: pd.DataFrame) -> pd.Series:
    # locate a suitable row (prefer 2024, else last non-empty)
    year_col = None
    for c in df.columns:
        if normalize_col(str(c)) in ("year", "years"):
            year_col = c
            break
    df_work = df.dropna(how="all")
    if year_col is not None and year_col in df_work.columns:
        try:
            df_year = df_work[df_work[year_col] == 2024]
            if df_year.empty:
                df_year = df_work.tail(1)
        except Exception:
            df_year = df_work.tail(1)
    else:
        df_year = df_work.tail(1)
    row = df_year.iloc[0]
    # build capacity map; if columns are multiindex or odd headers, normalize
    values = {}
    for key, cands in KNOWN_CAPACITY.items():
        for cand in cands:
            cand_norm = normalize_col(cand)
            for col in df_work.columns:
                try:
                    col_norm = normalize_col(str(col))
                except Exception:
                    col_norm = normalize_col(str(col))
                if col_norm == cand_norm:
                    try:
                        values[key] = float(row[col])
                    except Exception:
                        pass
                    break
            if key in values:
                break
    return pd.Series(values)


def load_figure7_dataframe(xlsx_path: str) -> pd.DataFrame:
    sheets = pd.read_excel(xlsx_path, sheet_name=None, engine="openpyxl")
    for name in sheets.keys():
        for header_row in range(0, 6):
            try:
                df = pd.read_excel(xlsx_path, sheet_name=name, engine="openpyxl", header=header_row)
                df = df.dropna(axis=1, how="all").dropna(how="all")
                if df.shape[1] == 0:
                    continue
                s = extract_2024_capacity(df)
                if len(s) >= 4:  # at least a few sources found
                    return df
            except Exception:
                continue
    return pd.read_excel(xlsx_path, sheet_name=0, engine="openpyxl")


def plot_capacity_bars(series: pd.Series) -> None:
    # order: coal, natural_gas, nuclear, hydro, wind, solar, storage, oil, other
    order = ["coal", "natural_gas", "nuclear", "hydro", "wind", "solar", "storage", "oil", "other"]
    labels_map = {
        "coal": "Coal",
        "natural_gas": "Natural Gas",
        "nuclear": "Nuclear",
        "hydro": "Hydropower",
        "wind": "Wind",
        "solar": "Solar",
        "storage": "Storage",
        "oil": "Oil/Petroleum",
        "other": "Other",
    }
    colors = {
        "coal": "#2C3E50",
        "natural_gas": "#C0392B",
        "nuclear": "#F1C40F",
        "hydro": "#2980B9",
        "wind": "#16A085",
        "solar": "#F39C12",
        "storage": "#8E44AD",
        "oil": "#7F8C8D",
        "other": "#95A5A6",
    }
    keys = [k for k in order if k in series.index]
    vals = [series[k] for k in keys]
    labels = [labels_map[k] for k in keys]
    cols = [colors[k] for k in keys]

    plt.figure(figsize=(10, 6))
    plt.bar(labels, vals, color=cols)
    plt.title("China Installed Generation Capacity by Source (2024)", pad=14)
    plt.ylabel("Capacity (GW)")
    plt.xticks(rotation=30, ha="right")
    for i, v in enumerate(vals):
        plt.text(i, v, f"{v:.0f}", ha="center", va="bottom", fontsize=9)
    plt.grid(axis="y", alpha=0.3, linestyle="--")
    plt.tight_layout()
    plt.figtext(0.5, 0.01,
                "Data source: U.S. EIA – figure7_data.xlsx (retrieved)\n"
                f"{EIA_FIG7_URL}",
                ha="center", fontsize=9, color="#555555")
    ensure_dir(os.path.dirname(OUTPUT_PATH))
    plt.savefig(OUTPUT_PATH, dpi=200)
    plt.close()


def main():
    ensure_dir(DATA_DIR)
    xlsx_path = os.path.join(DATA_DIR, "figure7_data.xlsx")
    alt_xlsx_path = os.path.join(ALT_DATA_DIR, "figure7_data.xlsx")
    try:
        download_excel(EIA_FIG7_URL, xlsx_path)
    except Exception as e:
        print("Failed to download EIA Excel:", e)
        print("If network blocks the file, please manually download from:")
        print(EIA_FIG7_URL)
        print(f"and place it at: {xlsx_path}")
    if not os.path.exists(xlsx_path) and os.path.exists(alt_xlsx_path):
        xlsx_path = alt_xlsx_path
        print("Using alternative data path:", xlsx_path)
    print("Reading:", xlsx_path, "exists=", os.path.exists(xlsx_path))
    df = load_figure7_dataframe(xlsx_path)
    s = extract_2024_capacity(df)
    plot_capacity_bars(s)
    print(f"Saved capacity chart to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
