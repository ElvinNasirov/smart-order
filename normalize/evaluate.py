import sqlite3, pandas as pd
gt = pd.read_csv("data/ground_truth.csv")
db = pd.read_sql("select * from sales", sqlite3.connect("output/sales.db"))
for d in (gt, db):
    d["date"] = pd.to_datetime(d["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    d["product_name"] = d["product_name"].astype(str).str.strip().str.capitalize()
m = gt.merge(db, on=["date", "product_name"], how="left", suffixes=("_true", "_pred"))
print(f"Строк эталона найдено в БД: {m.quantity_pred.notna().mean():.1%}")
print(f"Количество совпало точно:   {(m.quantity_true == m.quantity_pred).mean():.1%}")
print(f"Сумма qty эталон/БД: {gt.quantity.sum():.0f} / {db.quantity.sum():.0f}")
err = pd.read_csv("data/errors_injected.csv")
iss = pd.read_excel("output/issues_report.xlsx")
print(f"Ошибок внесено: {len(err)}, найдено правилами: {(iss.action != 'Замечание ИИ').sum()}")