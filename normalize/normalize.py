import json, re, sqlite3, requests, pandas as pd
from functools import lru_cache
from pathlib import Path

MODEL = "qwen2.5:7b-instruct"
FIELDS = ["date", "store", "product_name", "category", "quantity", "unit_price"]
OUT = Path("output"); OUT.mkdir(exist_ok=True)
RU = {"янв":"01","фев":"02","мар":"03","апр":"04","ма":"05","июн":"06","июл":"07",
      "авг":"08","сен":"09","окт":"10","ноя":"11","дек":"12"}

PROMPT = """Это первые 20 строк листа Excel с продажами магазина (номер строки: значения):
{preview}
Верни JSON:
{{"header_row": номер строки с заголовками,
 "layout": "long" если строка = товар за день, "wide" если колонки это дни месяца,
 "columns": {{"date": имя колонки или null, "store": колонка магазина/филиала или null, "product_name": ..., "category": ... или null,
   "quantity": ... или null, "unit_price": ... или null}},
 "day_columns": [имена колонок-дней, только для wide],
 "problems": [замеченные проблемы в данных, по-русски]}}
Имена колонок пиши точно как в строке заголовков."""

def ask_llm(prompt):
    r = requests.post("http://localhost:11434/api/generate", timeout=600, json={
        "model": MODEL, "prompt": prompt, "format": "json", "stream": False,
        "options": {"temperature": 0}})
    return json.loads(r.json()["response"])

def name(c):
    return str(int(c)) if isinstance(c, float) and c.is_integer() else str(c).strip()

@lru_cache(maxsize=None)
def parse_date(x):
    if isinstance(x, pd.Timestamp) or hasattr(x, "year"): return pd.Timestamp(x)
    s = str(x).lower().strip()
    if re.match(r"^\d{4}-\d{1,2}-\d{1,2}", s):
        return pd.to_datetime(s[:10], format="%Y-%m-%d", errors="coerce")
    for k, v in RU.items():
        if k in s: s = re.sub(k + r"\w*\.?", v, s); break
    return pd.to_datetime(s, dayfirst=True, errors="coerce")

@lru_cache(maxsize=None)
def to_num(x):
    s = re.sub(r"[^\d,.\-]", "", str(x)).replace(",", ".")
    return pd.to_numeric(s, errors="coerce")

def load(path):
    if path.suffix.lower() == ".csv":
        raw = pd.read_csv(path, header=None, dtype=str)
    else:
        raw = pd.read_excel(path, sheet_name=0, header=None)
    preview = "\n".join(f"{i}: {list(r)}" for i, r in raw.head(20).iterrows())
    m = ask_llm(PROMPT.format(preview=preview))
    print(path.name, "->", json.dumps(m, ensure_ascii=False))
    h = int(m["header_row"])
    df = raw.iloc[h + 1:].copy()
    df.columns = [name(c) for c in raw.iloc[h]]
    cols = {k: v for k, v in m["columns"].items() if v and v in df.columns}
    if m.get("layout") == "wide":
        days = [d for d in map(str, m.get("day_columns", [])) if d in df.columns]
        df = df.melt(id_vars=[cols["product_name"]], value_vars=days,
                     var_name="day", value_name="quantity")
        y, mo = re.search(r"(\d{4})\D(\d{1,2})", path.stem).groups()
        df["date"] = pd.to_datetime(y + "-" + mo + "-" + df["day"].str.extract(r"(\d+)")[0],
                                    errors="coerce")
        df = df.rename(columns={cols["product_name"]: "product_name"})
        df = df[df["date"].notna() | df["quantity"].notna()]
    else:
        df = df.rename(columns={v: k for k, v in cols.items()})
    for f in FIELDS:
        if f not in df: df[f] = None
    df = df[FIELDS].copy()
    df["source_file"] = path.name
    df["excel_row"] = df.index + 2
    return df, m.get("problems", [])

def check(df):
    df["date"] = df["date"].map(parse_date)
    df["quantity"] = df["quantity"].map(to_num)
    df["unit_price"] = df["unit_price"].map(to_num)
    df["product_name"] = df["product_name"].astype(str).str.strip().str.capitalize()
    df = df[df["product_name"].notna() & (df["product_name"] != "Nan")]
    rules = [
        (df["date"].isna(), "Неверная или пустая дата", "Уточнить у менеджера, строку не грузим"),
        (df["quantity"].isna(), "Количество не число или пусто", "Уточнить, строку не грузим"),
        (df["quantity"] < 0, "Отрицательное количество", "Скорее всего возврат, проверить"),
        (df.duplicated(["date", "store", "product_name", "quantity"]), "Дубликат строки", "Удалить дубль"),
    ]
    issues, bad = [], pd.Series(False, index=df.index)
    for mask, problem, action in rules:
        for _, r in df[mask].iterrows():
            issues.append({"file": r.source_file, "row": r.excel_row,
                           "product": r.product_name, "problem": problem, "action": action})
        bad |= mask
    return df[~bad], issues

frames, issues = [], []
files = [p for p in sorted(Path("data").iterdir())
         if p.suffix.lower() in (".xlsx", ".xls", ".csv")
         and p.name not in ("ground_truth.csv", "errors_injected.csv")]
for path in files:
    df, llm_problems = load(path)
    good, iss = check(df)
    frames.append(good); issues += iss
    issues += [{"file": path.name, "row": "", "product": "", "problem": p,
                "action": "Замечание ИИ"} for p in llm_problems]

sales = pd.concat(frames, ignore_index=True).drop(columns="excel_row")
sales["date"] = sales["date"].dt.strftime("%Y-%m-%d")
with sqlite3.connect(OUT / "sales.db") as con:
    sales.to_sql("sales", con, if_exists="replace", index=False)
sales.head(50000).to_excel(OUT / "normalized.xlsx", index=False)  # полная база в sales.db
pd.DataFrame(issues).to_excel(OUT / "issues_report.xlsx", index=False)
print(f"OK: {len(sales)} строк в БД, {len(issues)} проблем в отчёте")
