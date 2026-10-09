import io, os, subprocess, sys
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Smart Order", layout="wide")
st.title("Smart Order: Tomorrow's Order")
DATA, OUT, ORDERS = Path("data"), Path("output"), Path("orders")

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
    today_file = st.file_uploader("Upload today's sales report", type=["xlsx", "xls"],
                                  key="today_upload")
    if today_file is not None:
        upload_key = f"{today_file.name}_{today_file.size}"
        if st.session_state.get("processed_upload") != upload_key:
            st.session_state["processed_upload"] = upload_key
            with st.status("Processing today's report...", expanded=True) as s:
                s.write("1/3 Saving file...")
                (DATA / today_file.name).write_bytes(today_file.getbuffer())

                s.write("2/3 Local AI is checking and adding the report to the database (may take 1-2 min)...")
                ok = run("normalize/normalize.py")
                if not ok:
                    s.update(label="Failed at step 2", state="error")
                    st.stop()

                s.write("3/3 Forecasting tomorrow's order...")
                ok = run("ml/forecast.py")
                if not ok:
                    s.update(label="Failed at step 3", state="error")
                    st.stop()

                _fc = OUT / "forecast_tomorrow.csv"
                _date = pd.read_csv(_fc)["date"].iloc[0] if _fc.exists() else ""
                s.update(label=f"Done: order for {_date} is ready", state="complete")

    if st.button("Build Forecast"):
        with st.spinner("Model is calculating..."):
            run("ml/forecast.py")

    f = OUT / "forecast_tomorrow.csv"
    if f.exists():
        fc = pd.read_csv(f)
        forecast_date = fc["date"].iloc[0] if "date" in fc.columns and len(fc) else ""
        st.success(f"Order for {forecast_date} is ready")
        st.dataframe(fc)

        buf = io.BytesIO()
        fc.to_excel(buf, index=False, engine="openpyxl")
        st.download_button("Download Order", buf.getvalue(),
                           "order_tomorrow.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if st.button("Add to order history workbook"):
            ORDERS.mkdir(exist_ok=True)
            hist_path = ORDERS / "order_history.xlsx"
            sheet = (forecast_date or "forecast")[:31]
            if hist_path.exists():
                with pd.ExcelWriter(hist_path, engine="openpyxl", mode="a",
                                    if_sheet_exists="replace") as w:
                    fc.to_excel(w, sheet_name=sheet, index=False)
            else:
                with pd.ExcelWriter(hist_path, engine="openpyxl") as w:
                    fc.to_excel(w, sheet_name=sheet, index=False)
            st.success(f"Saved to {hist_path.resolve()}")
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
