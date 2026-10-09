import subprocess, sys
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Smart Order", layout="wide")
st.title("Smart Order: заказ на завтра")
DATA, OUT = Path("data"), Path("output")

def run(script):
    r = subprocess.run([sys.executable, script], capture_output=True, text=True)
    st.code(r.stdout[-3000:] + r.stderr[-3000:])
    return r.returncode == 0

tab1, tab2, tab3 = st.tabs(["1. Загрузка данных", "2. Заказ на завтра", "3. Качество"])

with tab1:
    files = st.file_uploader("Excel-отчёты магазина", type=["xlsx", "xls"],
                             accept_multiple_files=True)
    if st.button("Проверить и свести в базу"):
        for f in files or []:
            (DATA / f.name).write_bytes(f.getbuffer())
        with st.spinner("Локальная модель анализирует файлы..."):
            ok = run("normalize/normalize.py")
        if ok:
            st.success("База готова")
            st.subheader("Найденные несостыковки")
            st.dataframe(pd.read_excel(OUT / "issues_report.xlsx"))
            st.subheader("Единая таблица")
            st.dataframe(pd.read_excel(OUT / "normalized.xlsx"))

with tab2:
    if st.button("Построить прогноз"):
        with st.spinner("Модель считает..."):
            run("ml/forecast.py")
    f = OUT / "forecast_tomorrow.csv"
    if f.exists():
        fc = pd.read_csv(f)
        st.dataframe(fc)
        st.download_button("Скачать заказ", fc.to_csv(index=False), "order.csv")
    else:
        st.info("Прогноз появится, когда Роя добавит ml/forecast.py")

with tab3:
    if Path("normalize/evaluate.py").exists() and st.button("Оценить нормализацию"):
        run("normalize/evaluate.py")
    m = Path("reports/model_comparison.csv")
    if m.exists():
        st.dataframe(pd.read_csv(m))
    img = Path("reports/figures/novruz_semeni.png")
    if img.exists():
        st.image(str(img))