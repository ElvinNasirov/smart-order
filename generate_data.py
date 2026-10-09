"""
Sintetik satış hesabatları generatoru (hakaton üçün).

Yaradılan fayllar:
  sales_2021_08.xlsx   Avqust 2021  - AZ sütunlar, DD.MM.YYYY, uzun format
  sales_2023_12.xlsx   Dekabr 2023  - RU sütunlar, "5 Dec 2023", 3 başlıq sətri, qiymətlər "1,20 AZN"
  sales_2025_01.xlsx   Yanvar 2025  - geniş format (mal x gün), qiymətlər ayrı vərəqdə
  ground_truth.csv     təmiz data (date, product_name, category, quantity, unit_price, source_file)
  errors_injected.csv  qəsdən əlavə olunan səhvlər (file, sheet, row, column, error_type, ...)
  products.csv         məhsul kataloqu (AZ adı, RU adı, kateqoriya, vahid)

Qeyd: quantity = 0 olan günlər ground_truth-a daxil edilmir.
Market adı uydurmadır.
"""
import calendar
from datetime import date

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

SEED = 42
rng = np.random.default_rng(SEED)

STORE_RU = "ООО «Ясамал Маркет», г. Баку"

# (AZ adı, RU adı, kateqoriya, 2021 qiyməti AZN, orta gündəlik satış, vahid, mövsüm etiketi)
PRODUCTS = [
    ("Süd 1L", "Молоко 1л", "Süd məhsulları", 1.20, 40, "ədəd", None),
    ("Kefir 1L", "Кефир 1л", "Süd məhsulları", 1.30, 20, "ədəd", None),
    ("Qatıq 0.5kq", "Катык 0,5кг", "Süd məhsulları", 1.00, 18, "ədəd", "summer"),
    ("Ağ pendir", "Сыр белый", "Süd məhsulları", 7.50, 6, "kq", "festive"),
    ("Kərə yağı 200q", "Масло сливочное 200г", "Süd məhsulları", 2.80, 12, "ədəd", "festive"),
    ("Yumurta 10 əd", "Яйца 10 шт", "Süd məhsulları", 1.90, 15, "ədəd", "festive"),
    ("Ağ çörək", "Хлеб белый", "Çörək", 0.50, 80, "ədəd", None),
    ("Lavaş", "Лаваш", "Çörək", 0.40, 50, "ədəd", None),
    ("Dondurma plombir", "Мороженое пломбир", "Dondurma", 0.80, 25, "ədəd", "ice"),
    ("Dondurma eskimo", "Мороженое эскимо", "Dondurma", 0.70, 20, "ədəd", "ice"),
    ("Qara çay 100q", "Чай черный 100г", "Çay və qəhvə", 2.20, 15, "ədəd", "winter"),
    ("Yaşıl çay 100q", "Чай зеленый 100г", "Çay və qəhvə", 2.50, 6, "ədəd", None),
    ("Qəhvə 100q", "Кофе растворимый 100г", "Çay və qəhvə", 3.50, 8, "ədəd", "winter"),
    ("Şəkər 1kq", "Сахар 1кг", "Bakaleya", 1.60, 15, "ədəd", "festive"),
    ("Düyü 1kq", "Рис 1кг", "Bakaleya", 2.40, 10, "ədəd", "festive"),
    ("Makaron 400q", "Макароны 400г", "Bakaleya", 1.10, 14, "ədəd", None),
    ("Günəbaxan yağı 1L", "Масло подсолнечное 1л", "Bakaleya", 2.90, 9, "ədəd", "festive"),
    ("Mal əti", "Говядина", "Ət", 12.00, 8, "kq", "meat"),
    ("Toyuq", "Курица", "Ət", 5.50, 12, "kq", "meat"),
    ("Qoyun əti", "Баранина", "Ət", 13.00, 5, "kq", "meat"),
    ("Şokolad", "Шоколад плиточный", "Şirniyyat", 1.50, 18, "ədəd", "sweet"),
    ("Peçenye 300q", "Печенье 300г", "Şirniyyat", 1.80, 14, "ədəd", "sweet"),
    ("Konfet", "Конфеты", "Şirniyyat", 9.00, 4, "kq", "sweet"),
    ("Tort", "Торт", "Şirniyyat", 12.00, 2, "ədəd", "cake"),
    ("Su 1.5L", "Вода 1,5л", "İçkilər", 0.50, 45, "ədəd", "summer"),
    ("Limonad 1L", "Лимонад 1л", "İçkilər", 1.10, 20, "ədəd", "summer"),
    ("Alma", "Яблоки", "Meyvə-tərəvəz", 1.80, 15, "kq", None),
    ("Pomidor", "Помидоры", "Meyvə-tərəvəz", 2.00, 18, "kq", "summer"),
    ("Mandarin", "Мандарины", "Meyvə-tərəvəz", 2.50, 6, "kq", "mandarin"),
]
AZ_TO_RU = {p[0]: p[1] for p in PRODUCTS}
MONTHS_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

GROUND_TRUTH = []
ERRORS = []


# ---------------------------------------------------------------- simulyasiya
def season_factor(tag, d):
    m, day = d.month, d.day
    # Yeni il təsiri: 20 dekabrdan artır, 31-də pik, yanvarın ilk həftəsi azalır
    if m == 12 and day >= 20:
        ny = 1 + (day - 19) / 12 * 2.0
    elif m == 1 and day <= 2:
        ny = 2.0
    elif m == 1 and day <= 7:
        ny = 1.3
    else:
        ny = 1.0

    f = {
        "ice": {8: 3.0, 12: 0.3, 1: 0.2},
        "summer": {8: 1.7, 12: 0.8, 1: 0.7},
        "winter": {8: 0.8, 12: 1.2, 1: 1.3},
        "mandarin": {8: 0.1, 12: 3.0, 1: 2.0},
    }.get(tag, {}).get(m, 1.0)

    if tag in ("sweet", "mandarin"):
        f *= ny
    elif tag == "cake":
        f *= ny ** 1.5
    elif tag == "meat":
        f *= 1 + (ny - 1) * 0.8
    elif tag == "festive":
        f *= 1 + (ny - 1) * 0.5

    if d.weekday() >= 5:  # həftəsonu
        f *= 1.2
    if m == 1 and day == 1:  # 1 yanvar mağaza sakit olur
        f *= 0.6
    return f


def sample_qty(base, unit, factor):
    lam = base * factor
    if unit == "kq":
        return round(float(rng.gamma(4.0, lam / 4.0)), 1) if lam > 0 else 0.0
    return int(rng.poisson(lam))


def price_for(p2021, year):
    k = {2021: 1.00, 2023: 1.25, 2025: 1.40}[year]
    return round(round(p2021 * k / 0.05) * 0.05, 2)


def simulate(year, month):
    n = calendar.monthrange(year, month)[1]
    out = []
    for day in range(1, n + 1):
        d = date(year, month, day)
        for az, ru, cat, p, base, unit, tag in PRODUCTS:
            q = sample_qty(base, unit, season_factor(tag, d))
            out.append(dict(date=d, az=az, ru=ru, category=cat, unit=unit,
                            quantity=q, unit_price=price_for(p, year)))
    return out


def add_ground_truth(sim, source_file):
    for r in sim:
        if r["quantity"] > 0:
            GROUND_TRUTH.append(dict(date=r["date"].isoformat(), product_name=r["az"],
                                     category=r["category"], quantity=float(r["quantity"]),
                                     unit_price=r["unit_price"], source_file=source_file))


# ---------------------------------------------------------------- səhvlər
def log(file, sheet, row, column, etype, orig, new):
    ERRORS.append(dict(file=file, sheet=sheet, row=row, column=column, error_type=etype,
                       original_value=orig, injected_value=new))


def col_label(header, idx):
    return f"{get_column_letter(idx + 1)} ({header[idx]})"


def inject_long(rows, file, sheet, first_row, header, c, typo_map, text_fn, bad_date):
    """Uzun format üçün 8 səhv. c = {'date','name','qty','price'} sütun indeksləri."""
    # 1) təkrarlanan sətir
    k = int(rng.integers(0, len(rows) - 1))
    rows.insert(k + 1, list(rows[k]))
    log(file, sheet, first_row + k + 1, "(bütün sətir)", "duplicate_row",
        f"{first_row + k}-ci sətrin surəti", "")
    used = {k, k + 1}

    def pick(cond=lambda r: True):
        while True:
            i = int(rng.integers(0, len(rows)))
            if i not in used and cond(rows[i]):
                used.add(i)
                return i

    def is_pos(r):
        return isinstance(r[c["qty"]], (int, float)) and r[c["qty"]] > 0

    # 2-3) boş xanalar
    for col in (c["qty"], c["price"]):
        i = pick()
        log(file, sheet, first_row + i, col_label(header, col), "empty_cell", rows[i][col], "")
        rows[i][col] = None

    # 4-5) adda yazı səhvi
    for good, bad in list(typo_map.items())[:2]:
        i = pick(lambda r, g=good: r[c["name"]] == g)
        log(file, sheet, first_row + i, col_label(header, c["name"]), "typo", good, bad)
        rows[i][c["name"]] = bad

    # 6) mənfi miqdar
    i = pick(is_pos)
    q = rows[i][c["qty"]]
    log(file, sheet, first_row + i, col_label(header, c["qty"]), "negative_quantity", q, -q)
    rows[i][c["qty"]] = -q

    # 7) yanlış tarix
    i = pick()
    log(file, sheet, first_row + i, col_label(header, c["date"]), "invalid_date",
        rows[i][c["date"]], bad_date)
    rows[i][c["date"]] = bad_date

    # 8) rəqəm əvəzinə mətn
    i = pick(is_pos)
    q = rows[i][c["qty"]]
    log(file, sheet, first_row + i, col_label(header, c["qty"]), "text_instead_of_number",
        q, text_fn(q))
    rows[i][c["qty"]] = text_fn(q)


# ---------------------------------------------------------------- Excel
def write_table(ws, header, rows, start_row=1):
    for j, h in enumerate(header, 1):
        ws.cell(start_row, j, h).font = Font(bold=True)
    for i, r in enumerate(rows, start_row + 1):
        for j, v in enumerate(r, 1):
            if v is not None:
                ws.cell(i, j, v)
    for j in range(1, len(header) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 22 if j <= 2 else 12


def az_num(v):
    return f"{v}".replace(".", ",")


# ---------------------------------------------------------------- 1) Avqust 2021
def file_2021_08():
    fname, sheet = "sales_2021_08.xlsx", "Satış"
    sim = simulate(2021, 8)
    add_ground_truth(sim, fname)

    header = ["Tarix", "Məhsul", "Kateqoriya", "Miqdar", "Qiymət"]
    rows = [[r["date"].strftime("%d.%m.%Y"), r["az"], r["category"], r["quantity"], r["unit_price"]]
            for r in sim if r["quantity"] > 0]

    inject_long(rows, fname, sheet, first_row=2, header=header,
                c=dict(date=0, name=1, qty=3, price=4),
                typo_map={"Süd 1L": "Sud 1L", "Ağ çörək": "Ag corek"},
                text_fn=lambda q: f"{az_num(q)} ədəd" if isinstance(q, int) else f"{az_num(q)} kq",
                bad_date="32.08.2021")

    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    write_table(ws, header, rows)
    wb.save(fname)


# ---------------------------------------------------------------- 2) Dekabr 2023
def file_2023_12():
    fname, sheet = "sales_2023_12.xlsx", "Продажи"
    sim = simulate(2023, 12)
    add_ground_truth(sim, fname)

    def azn(v):
        return f"{v:.2f}".replace(".", ",") + " AZN"

    header = ["Дата", "Наименование", "Кол-во", "Цена за ед.", "Сумма"]
    clean = [r for r in sim if r["quantity"] > 0]
    rows = [[f"{r['date'].day} {MONTHS_EN[r['date'].month - 1]} {r['date'].year}",
             r["ru"], r["quantity"], azn(r["unit_price"]), azn(r["quantity"] * r["unit_price"])]
            for r in clean]
    total_sum = sum(r["quantity"] * r["unit_price"] for r in clean)

    inject_long(rows, fname, sheet, first_row=5, header=header,
                c=dict(date=0, name=1, qty=2, price=3),
                typo_map={"Молоко 1л": "Moloko 1л", "Хлеб белый": "Хлеб белыи"},
                text_fn=lambda q: f"{q} шт" if isinstance(q, int) else f"{az_num(q)} кг",
                bad_date="31 Nov 2023")

    # 1C üslubunda sonda "Итого" sətri
    rows.append(["Итого", None, None, None, azn(total_sum)])
    log(fname, sheet, 5 + len(rows) - 1, "(bütün sətir)", "total_row", "", "Итого")

    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws["A1"] = STORE_RU
    ws["A1"].font = Font(bold=True, size=13)
    ws.merge_cells("A1:E1")
    ws["A2"] = "Отчет о продажах за декабрь 2023 г."
    ws.merge_cells("A2:E2")
    ws["A3"] = "Валюта: AZN"
    write_table(ws, header, rows, start_row=4)
    wb.save(fname)


# ---------------------------------------------------------------- 3) Yanvar 2025 (geniş)
def file_2025_01():
    fname, s_sales, s_price = "sales_2025_01.xlsx", "Satış", "Qiymətlər"
    sim = simulate(2025, 1)
    add_ground_truth(sim, fname)

    days = 31
    grid = {}
    for r in sim:
        grid.setdefault(r["az"], {})[r["date"].day] = r["quantity"]

    header = ["Məhsul"] + list(range(1, days + 1)) + ["Cəmi"]
    rows = []
    for az, *_ in PRODUCTS:
        vals = [grid[az][d] for d in range(1, days + 1)]
        rows.append([az] + vals + [round(sum(vals), 1)])

    # 1) təkrarlanan məhsul sətri
    k = int(rng.integers(0, len(rows) - 1))
    rows.insert(k + 1, list(rows[k]))
    log(fname, s_sales, 2 + k + 1, "(bütün sətir)", "duplicate_row", f"{2 + k}-ci sətrin surəti", "")
    used_rows = {k, k + 1}

    # 2) yanlış tarix: mövcud olmayan "32" günü sütunu
    header.insert(days + 1, 32)
    for r in rows:
        r.insert(days + 1, int(rng.integers(1, 20)))
    log(fname, s_sales, 1, col_label(header, days + 1), "invalid_date", "", "32 yanvar sütunu")

    used_cells = set()

    def pick_cell(cond=lambda v: True):
        while True:
            i = int(rng.integers(0, len(rows)))
            d = int(rng.integers(1, days + 1))
            if i not in used_rows and (i, d) not in used_cells and cond(rows[i][d]):
                used_cells.add((i, d))
                return i, d

    def is_pos(v):
        return isinstance(v, (int, float)) and v > 0

    # 3) adda yazı səhvi
    i = next(j for j, r in enumerate(rows) if r[0] == "Dondurma plombir" and j not in used_rows)
    used_rows.add(i)
    log(fname, s_sales, 2 + i, col_label(header, 0), "typo", "Dondurma plombir", "Dondurma plombr")
    rows[i][0] = "Dondurma plombr"

    # 4-5) boş xanalar
    for _ in range(2):
        i, d = pick_cell()
        log(fname, s_sales, 2 + i, col_label(header, d), "empty_cell", rows[i][d], "")
        rows[i][d] = None

    # 6) mənfi miqdar
    i, d = pick_cell(is_pos)
    log(fname, s_sales, 2 + i, col_label(header, d), "negative_quantity", rows[i][d], -rows[i][d])
    rows[i][d] = -rows[i][d]

    # 7-8) rəqəm əvəzinə mətn
    for _ in range(2):
        i, d = pick_cell(is_pos)
        v = rows[i][d]
        txt = f"{v} əd" if isinstance(v, int) else f"{az_num(v)} kq"
        log(fname, s_sales, 2 + i, col_label(header, d), "text_instead_of_number", v, txt)
        rows[i][d] = txt

    # Qiymətlər vərəqi
    p_header = ["Məhsul", "Qiymət (AZN)", "Vahid"]
    p_rows = [[az, price_for(p, 2025), unit] for az, ru, cat, p, base, unit, tag in PRODUCTS]
    j = int(rng.integers(0, len(p_rows)))
    txt = f"{p_rows[j][1]:.2f}".replace(".", ",") + " AZN"
    log(fname, s_price, 2 + j, col_label(p_header, 1), "text_instead_of_number", p_rows[j][1], txt)
    p_rows[j][1] = txt

    wb = Workbook()
    ws = wb.active
    ws.title = s_sales
    write_table(ws, header, rows)
    for col in range(2, len(header) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 6
    ws.column_dimensions["A"].width = 22
    ws2 = wb.create_sheet(s_price)
    write_table(ws2, p_header, p_rows)
    wb.save(fname)


# ---------------------------------------------------------------- main
if __name__ == "__main__":
    file_2021_08()
    file_2023_12()
    file_2025_01()

    gt = pd.DataFrame(GROUND_TRUTH).sort_values(["source_file", "date", "product_name"])
    gt.to_csv("ground_truth.csv", index=False)
    pd.DataFrame(ERRORS).to_csv("errors_injected.csv", index=False)
    pd.DataFrame([dict(product_name=p[0], product_name_ru=p[1], category=p[2], unit=p[5])
                  for p in PRODUCTS]).to_csv("products.csv", index=False)

    print(f"ground_truth.csv: {len(gt)} sətir")
    print(pd.DataFrame(ERRORS).groupby("file").size().rename("səhv sayı"))
