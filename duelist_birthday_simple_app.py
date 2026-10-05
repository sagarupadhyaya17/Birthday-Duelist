"""Simple Duelist birthday app.  Run:  streamlit run duelist_birthday_simple.py"""
import calendar
import io
import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Tomorrow's Birthdays", page_icon="🎂", layout="wide")
st.title("🎂 Tomorrow's Birthdays")

WANTED = {  # our name -> accepted header names (lower-case, letters/digits only)
    "Client Code": ["ClientCode", "clientcode"],
    "Name": ["Name","name", "clientname"],
    "Main Code": ["MainCode","maincode"],
    "Mobile": ["Mobile","mobileno", "mobile", "mobilenumber"],
    "Date of Birth": ["DateOfBirth","dateofbirth", "dob"],
}
ENGINES = {"xlsb": "pyxlsb", "xlsx": "openpyxl", "xlsm": "openpyxl", "xls": "xlrd"}


def key(text):
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


@st.cache_data(show_spinner="Reading file…")
def load(data: bytes, filename: str) -> pd.DataFrame:
    ext = filename.rsplit(".", 1)[-1].lower()
    read = (lambda **kw: pd.read_csv(io.BytesIO(data), **kw)) if ext == "csv" else (
        lambda **kw: pd.read_excel(io.BytesIO(data), engine=ENGINES[ext], **kw))

    headers = {key(h): h for h in read(nrows=0).columns}
    found = {}
    for ours, names in WANTED.items():
        match = next((headers[n] for n in names if n in headers), None)
        if match is None:
            st.error(f"Column for '{ours}' not found in the file.")
            st.stop()
        found[ours] = match

    df = read(usecols=list(found.values()), dtype=str)[list(found.values())]
    df.columns = list(found)

    for col in df.columns:  # tidy text, keep leading zeros
        df[col] = df[col].str.strip().str.replace(r"\.0+$", "", regex=True)

    dob = df["Date of Birth"]
    serial = pd.to_numeric(dob, errors="coerce").between(1, 80000)  # Excel serial dates
    from_serial = pd.to_datetime(pd.to_numeric(dob.where(serial)), unit="D", origin="1899-12-30")
    from_text = pd.to_datetime(dob.where(~serial), dayfirst=True, format="mixed", errors="coerce")
    df["Date of Birth"] = from_serial.astype("datetime64[ns]").fillna(from_text.astype("datetime64[ns]"))
    return df


def to_excel(df: pd.DataFrame) -> bytes:
    out = df.copy()
    out["Date of Birth"] = out["Date of Birth"].dt.strftime("%d-%b-%Y")  # all text -> codes stay text
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        out.to_excel(xw, index=False, sheet_name="Tomorrow Birthdays")
        ws = xw.sheets["Tomorrow Birthdays"]
        for i, col in enumerate(out.columns):
            ws.set_column(i, i, max(14, out[col].astype(str).str.len().max() + 2 if len(out) else 14))
    return buf.getvalue()


file = st.file_uploader("Upload Duelist file", type=["xlsb", "xlsx", "xlsm", "xls", "csv"])
if not file:
    st.info("Upload the Duelist file to begin.")
    st.stop()

df = load(file.getvalue(), file.name)
tomorrow = datetime.now(ZoneInfo("Asia/Kathmandu")).date() + timedelta(days=1)

m, d = df["Date of Birth"].dt.month, df["Date of Birth"].dt.day
match = (m == tomorrow.month) & (d == tomorrow.day)
if tomorrow.month == 2 and tomorrow.day == 28 and not calendar.isleap(tomorrow.year):
    match |= (m == 2) & (d == 29)  # leap-day birthdays
result = df[match].sort_values("Name")

c1, c2, c3 = st.columns(3)
c1.metric("Records in file", f"{len(df):,}")
c2.metric("Valid birth dates", f"{df['Date of Birth'].notna().sum():,}")
c3.metric(f"Birthdays on {tomorrow:%d %b %Y}", len(result))

if result.empty:
    st.success("No client has a birthday tomorrow.")
else:
    st.dataframe(result, hide_index=True, width="stretch",
                 column_config={"Date of Birth": st.column_config.DateColumn(format="DD MMM YYYY")})
    st.download_button("⬇️ Download Excel", to_excel(result),
                       f"Tomorrow_Birthdays_{tomorrow:%d-%b-%Y}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")