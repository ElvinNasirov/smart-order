import io, os, subprocess, sys
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Smart Order", layout="wide")
st.title("Smart Order: Tomorrow's Order")
DATA, OUT = Path("data"), Path("output")

def run(script):
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    r = subprocess.run([sys.executable, script], capture_output=True, text=True,
                       encoding="utf-8", errors="replace", env=env)
    st.code(r.stdout[-3000:] + r.stderr[-3000:])
    return r.returncode == 0

tab1, tab2, tab3 = st.tabs(["1. Load Data", "2. Tomorrow's Order", "3. Quality"])

with tab1:
    files = st.file_uploader("Store Excel reports", type=["xlsx", "xls"],
                             accept_multiple_files=True)
    if st.button("Validate & Import"):
        for f in files or []:
            (DATA / f.name).write_bytes(f.getbuffer())
        with st.spinner("Local model analyzing files..."):
            ok = run("normalize/normalize.py")
        if ok:
            st.success("Database ready")
            st.subheader("Data Issues Found")
            st.dataframe(pd.read_excel(OUT / "issues_report.xlsx"))
            st.subheader("Unified Table")
            st.dataframe(pd.read_excel(OUT / "normalized.xlsx"))

with tab2:
    if st.button("Build Forecast"):
        with st.spinner("Model is calculating..."):
            run("ml/forecast.py")
    f = OUT / "forecast_tomorrow.csv"
    if f.exists():
        fc = pd.read_csv(f)
        if "date" in fc.columns and len(fc):
            st.subheader(f"Order for {fc['date'].iloc[0]}")
        st.dataframe(fc)
        buf = io.BytesIO()
        fc.to_excel(buf, index=False, engine="openpyxl")
        st.download_button("Download Order", buf.getvalue(),
                           "order_tomorrow.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Click 'Build Forecast'")

with tab3:
    if Path("normalize/evaluate.py").exists() and st.button("Evaluate Normalization"):
        run("normalize/evaluate.py")
    m = Path("reports/model_comparison.csv")
    if m.exists():
        st.dataframe(pd.read_csv(m))
    img = Path("reports/figures/novruz_semeni.png")
    if img.exists():
        st.image(str(img))
