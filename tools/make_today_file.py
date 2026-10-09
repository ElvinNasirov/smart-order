"""
Generate data_demo/today_sales.xlsx: one day of sales for the day after the last
date in output/sales.db. Quantities are taken from the same weekday one week
earlier, scaled by a random factor in [0.85, 1.15].
Run from the repo root: python tools/make_today_file.py
"""
import sqlite3, numpy as np, pandas as pd
from pathlib import Path

rng = np.random.default_rng(42)

con = sqlite3.connect("output/sales.db")
df = pd.read_sql(
    "SELECT date, store, product_name, category, quantity, unit_price FROM sales", con
)
df["date"] = pd.to_datetime(df["date"])

last = df["date"].max()
tomorrow = last + pd.Timedelta(days=1)
ref_date = tomorrow - pd.Timedelta(days=7)

base = df[df["date"] == ref_date][
    ["store", "product_name", "category", "quantity", "unit_price"]
].copy()

if base.empty:
    print(f"No data for {ref_date.date()}, falling back to {last.date()}")
    base = df[df["date"] == last][
        ["store", "product_name", "category", "quantity", "unit_price"]
    ].copy()

factors = rng.uniform(0.85, 1.15, len(base))
base["quantity"] = (base["quantity"] * factors).round().astype(int).clip(lower=0)
base.insert(0, "date", tomorrow.strftime("%Y-%m-%d"))

out_dir = Path("data_demo")
out_dir.mkdir(exist_ok=True)
out_path = out_dir / "today_sales.xlsx"
base.to_excel(out_path, index=False)
print(f"Written {len(base)} rows -> {out_path}  (date={tomorrow.date()})")
