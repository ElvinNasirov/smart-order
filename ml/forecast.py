import sqlite3, numpy as np, pandas as pd, lightgbm as lgb
from pathlib import Path

OUT, REP = Path("output"), Path("reports"); REP.mkdir(exist_ok=True)

HOL_NAMES_FIXED = {
    "01-01": "New Year",        "01-02": "New Year",
    "03-08": "Women's Day",
    "03-20": "Novruz",          "03-21": "Novruz",  "03-22": "Novruz",
    "03-23": "Novruz",          "03-24": "Novruz",
    "05-09": "Victory Day",     "05-28": "Independence Day",
    "06-15": "National Salvation Day",
    "11-08": "Victory Day",     "11-09": "Flag Day",
    "12-31": "Solidarity Day",
}
RAMAZAN = ["2019-06-05","2020-05-24","2021-05-13","2022-05-02","2023-04-21",
           "2024-04-10","2025-03-30","2026-03-20"]
QURBAN  = ["2019-08-12","2020-07-31","2021-07-20","2022-07-09","2023-06-28",
           "2024-06-16","2025-06-06","2026-05-27"]

hol_named = {}
hol = []
for y in range(2015, 2028):
    for md, hname in HOL_NAMES_FIXED.items():
        dt = pd.Timestamp(f"{y}-{md}")
        hol_named[dt] = hname
        hol.append(f"{y}-{md}")
for d in RAMAZAN:
    t = pd.Timestamp(d)
    hol_named[t] = "Ramadan Bayram"
    hol_named[t + pd.Timedelta(days=1)] = "Ramadan Bayram"
    hol.extend([d, str((t + pd.Timedelta(days=1)).date())])
for d in QURBAN:
    t = pd.Timestamp(d)
    hol_named[t] = "Gurban Bayram"
    hol_named[t + pd.Timedelta(days=1)] = "Gurban Bayram"
    hol.extend([d, str((t + pd.Timedelta(days=1)).date())])
HOL = pd.DatetimeIndex(sorted(set(pd.to_datetime(hol))))

def upcoming_holiday_label(date):
    for delta in range(4):
        candidate = date + pd.Timedelta(days=delta)
        if candidate in hol_named:
            name = hol_named[candidate]
            if delta == 0:
                return f"{name} today"
            elif delta == 1:
                return f"{name} tomorrow"
            else:
                return f"{name} in {delta} days"
    return ""

df = pd.read_sql("select date, product_name, quantity from sales",
                 sqlite3.connect(OUT / "sales.db"))
df["date"] = pd.to_datetime(df["date"])
df = df.groupby(["product_name", "date"], as_index=False)["quantity"].sum()
last = df["date"].max(); tomorrow = last + pd.Timedelta(days=1)
fut = pd.DataFrame({"product_name": df["product_name"].unique(), "date": tomorrow})
d = pd.concat([df, fut], ignore_index=True).sort_values(["product_name", "date"])

g = d.groupby("product_name")["quantity"]
for l in [1, 7, 14, 28, 364]:
    d[f"lag_{l}"] = g.shift(l)
for w in [7, 28]:
    d[f"avg_{w}"] = g.transform(lambda s: s.shift(1).rolling(w, min_periods=1).mean())
d["dow"] = d.date.dt.dayofweek; d["month"] = d.date.dt.month
doy = d.date.dt.dayofyear
d["day_sin"] = np.sin(2 * np.pi * doy / 365.25); d["day_cos"] = np.cos(2 * np.pi * doy / 365.25)
d["is_holiday"] = d.date.isin(HOL).astype(int)
nxt = HOL.values[np.minimum(np.searchsorted(HOL.values, d.date.values), len(HOL) - 1)]
d["days_to_holiday"] = np.clip((nxt - d.date.values) / np.timedelta64(1, "D"), 0, 30)
d["product"] = d["product_name"].astype("category")
F = ["lag_1","lag_7","lag_14","lag_28","lag_364","avg_7","avg_28","dow","month",
     "day_sin","day_cos","is_holiday","days_to_holiday","product"]

hist, fut = d[d.date <= last], d[d.date == tomorrow].copy()
days = hist.date.nunique()
cut = last - pd.Timedelta(days=min(90, max(days // 5, 1)))
tr, te = hist[hist.date <= cut], hist[hist.date > cut]
wape = lambda y, p: np.abs(y - p).sum() / max(y.sum(), 1e-9)
m = lgb.LGBMRegressor(n_estimators=400, learning_rate=0.05, verbose=-1)
m.fit(tr[F], tr.quantity)
res = pd.DataFrame({"model": ["Yesterday", "Last week", "28-day avg", "LightGBM"],
    "WAPE": [wape(te.quantity, te.lag_1.fillna(0)), wape(te.quantity, te.lag_7.fillna(0)),
             wape(te.quantity, te.avg_28.fillna(0)), wape(te.quantity, m.predict(te[F]).clip(0))]})
res["WAPE"] = (res["WAPE"] * 100).round(1)
res.to_csv(REP / "model_comparison.csv", index=False); print(res)

m.fit(hist[F], hist.quantity)
fut["forecast"] = m.predict(fut[F]).clip(0).round(1)
fut["order_qty"] = np.ceil(fut["forecast"] * 1.1)
fut["upcoming_holiday"] = fut["date"].apply(upcoming_holiday_label)
out = fut[["product_name", "date", "forecast", "order_qty", "upcoming_holiday"]].sort_values("order_qty", ascending=False)
out["date"] = out["date"].dt.strftime("%Y-%m-%d")
out.to_csv(OUT / "forecast_tomorrow.csv", index=False)
print(f"Forecast for {tomorrow.date()} ready: {len(out)} products. Trained on {days} days.")
