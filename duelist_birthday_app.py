"""
Duelist Birthday Desk
---------------------
Upload a Duelist file (.xlsb / .xlsx / .xlsm / .xls / .csv), extract
ClientCode, Name, MainCode, Mobile and Date of Birth, and see who has a
birthday today, tomorrow or in the coming days. Export any list to Excel.

Run:  streamlit run duelist_birthday_app.py
"""
from __future__ import annotations

import calendar
import io
import re
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
import xlsxwriter

TZ = ZoneInfo("Asia/Kathmandu")

# Standard field -> accepted header names (compared after lower-casing and
# stripping everything except letters and digits).
ALIASES = {
    "ClientCode": ["ClientCode","clientcode", "clientid"],
    "Name": ["Name","name", "clientname", "customername"],
    "MainCode": ["MainCode","maincode"],
    "Mobile": ["MobileNo","mobileno", "mobile", "mobilenumber", "mobileno1", "phoneno"],
    "Mobile2": ["Phone","phone"],
    "DOB": ["DateOfBirth","dateofbirth", "dob"],
}
LABELS = {
    "ClientCode": "Client code",
    "Name": "Name",
    "MainCode": "Main code",
    "Mobile": "Mobile",
    "Mobile2": "Second mobile (optional fallback)",
    "DOB": "Date of birth",
}
REQUIRED = ["ClientCode", "Name", "MainCode", "Mobile", "DOB"]
NONE = "— none —"
DEFAULT_SMS = (
    "Dear {name}, wishing you a very Happy Birthday! "
    "Thank you for being a valued client."
)


# --------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------
def _norm(s) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def _engine(filename: str):
    ext = filename.rsplit(".", 1)[-1].lower()
    return {"xlsb": "pyxlsb", "xlsx": "openpyxl", "xlsm": "openpyxl", "xls": "xlrd"}.get(ext)


def _is_csv(filename: str) -> bool:
    return filename.lower().endswith((".csv", ".txt"))


@st.cache_data(show_spinner=False)
def list_sheets(data: bytes, filename: str) -> list[str]:
    if _is_csv(filename):
        return ["(csv)"]
    with pd.ExcelFile(io.BytesIO(data), engine=_engine(filename)) as xf:
        return list(xf.sheet_names)


@st.cache_data(show_spinner=False)
def read_headers(data: bytes, filename: str, sheet: str) -> list[str]:
    if _is_csv(filename):
        df = pd.read_csv(io.BytesIO(data), nrows=0, dtype=str, encoding_errors="replace")
    else:
        df = pd.read_excel(io.BytesIO(data), sheet_name=sheet, nrows=0, engine=_engine(filename))
    return [str(c) for c in df.columns]


@st.cache_data(show_spinner="Reading file…")
def read_columns(data: bytes, filename: str, sheet: str, cols: tuple) -> pd.DataFrame:
    """Read ONLY the requested columns, everything as text (keeps leading zeros)."""
    cols = list(cols)
    if _is_csv(filename):
        return pd.read_csv(io.BytesIO(data), usecols=cols, dtype=str, encoding_errors="replace")
    return pd.read_excel(
        io.BytesIO(data), sheet_name=sheet, usecols=cols, dtype=str, engine=_engine(filename)
    )


def detect_columns(headers: list[str]) -> dict[str, str | None]:
    by_norm = {}
    for h in headers:
        by_norm.setdefault(_norm(h), h)
    return {
        field: next((by_norm[a] for a in names if a in by_norm), None)
        for field, names in ALIASES.items()
    }


# --------------------------------------------------------------------------
# Cleaning
# --------------------------------------------------------------------------
_EMPTY = {"", "nan", "none", "null", "nat", "<na>"}


def clean_text(s: pd.Series, strip_decimal: bool = False) -> pd.Series:
    s = s.astype("string").str.strip()
    if strip_decimal:
        s = s.str.replace(r"^(\d+)\.0+$", r"\1", regex=True)
    return s.mask(s.str.lower().isin(_EMPTY))


def parse_dob(s: pd.Series) -> pd.Series:
    """Handles dd/mm/yyyy text, ISO text, Excel serial numbers and datetimes."""
    s = clean_text(s)
    num = pd.to_numeric(s, errors="coerce")
    serial = num.between(1, 80000).fillna(False).astype(bool)
    from_serial = pd.to_datetime(num.where(serial), unit="D", origin="1899-12-30", errors="coerce")
    text = s.where(~serial)
    p1 = pd.to_datetime(text, format="%d/%m/%Y", errors="coerce")
    left = text.where(p1.isna())
    p2 = pd.to_datetime(left, dayfirst=True, format="mixed", errors="coerce")
    out = from_serial.astype("datetime64[ns]")
    out = out.combine_first(p1.astype("datetime64[ns]")).combine_first(p2.astype("datetime64[ns]"))
    return out.dt.normalize()


def build_dataset(raw: pd.DataFrame, mapping: dict, fallback_mobile: bool, today: date):
    df = pd.DataFrame(index=raw.index)
    df["ClientCode"] = clean_text(raw[mapping["ClientCode"]], strip_decimal=True)
    df["Name"] = clean_text(raw[mapping["Name"]])
    df["MainCode"] = clean_text(raw[mapping["MainCode"]], strip_decimal=True)
    mobile = clean_text(raw[mapping["Mobile"]], strip_decimal=True)
    if fallback_mobile and mapping.get("Mobile2"):
        mobile = mobile.fillna(clean_text(raw[mapping["Mobile2"]], strip_decimal=True))
    df["Mobile"] = mobile
    df["DOB_raw"] = clean_text(raw[mapping["DOB"]])
    dob = parse_dob(raw[mapping["DOB"]])
    plausible = (dob.dt.year >= 1900) & (dob <= pd.Timestamp(today))
    df["DOB"] = dob.where(plausible)
    df["DOB_issue"] = df["DOB_raw"].notna() & df["DOB"].isna()
    return df


def with_birthdays(df: pd.DataFrame, ref: date) -> pd.DataFrame:
    v = df[df["DOB"].notna()].copy()
    if v.empty:
        for c in ("Next Birthday", "Days To Go", "Turning Age"):
            v[c] = pd.Series(dtype="float64")
        return v
    m, d = v["DOB"].dt.month, v["DOB"].dt.day

    def occurrence(year: int) -> pd.Series:
        dd = d.mask((m == 2) & (d == 29) & (not calendar.isleap(year)), 28)  # leap-day babies
        return pd.to_datetime(pd.DataFrame({"year": year, "month": m, "day": dd})).astype("datetime64[ns]")

    r = pd.Timestamp(ref)
    this_year = occurrence(ref.year)
    nxt = this_year.where(this_year >= r, occurrence(ref.year + 1))
    v["Next Birthday"] = nxt
    v["Days To Go"] = (nxt - r).dt.days
    v["Turning Age"] = nxt.dt.year - v["DOB"].dt.year
    return v


def unique_clients(df: pd.DataFrame) -> pd.DataFrame:
    return df[~df["ClientCode"].duplicated() | df["ClientCode"].isna()]


def collapse_clients(df: pd.DataFrame) -> pd.DataFrame:
    """One row per client: several loans -> MainCodes joined in one cell."""
    if df.empty:
        df = df.copy()
        df["Loans"] = pd.Series(dtype="int64")
        return df
    key = df["ClientCode"].where(df["ClientCode"].notna(), "~" + df.index.astype(str))
    g = df.groupby(key.rename("_key"), sort=False)
    out = g.first()
    out["MainCode"] = g["MainCode"].agg(lambda x: ", ".join(dict.fromkeys(x.dropna())))
    out["Loans"] = g.size()
    return out.reset_index(drop=True)


# --------------------------------------------------------------------------
# Presentation helpers
# --------------------------------------------------------------------------
VIEW_COLS = [
    ("ClientCode", "Client Code"),
    ("Name", "Name"),
    ("MainCode", "Main Code"),
    ("Mobile", "Mobile"),
    ("DOB", "Date of Birth"),
    ("Turning Age", "Turning Age"),
    ("Next Birthday", "Birthday On"),
    ("Days To Go", "Days To Go"),
]


def to_view(df: pd.DataFrame, template: str | None = None, show_loans: bool = False) -> pd.DataFrame:
    cols = list(VIEW_COLS)
    if show_loans and "Loans" in df:
        cols.append(("Loans", "Loans"))
    out = df[[c for c, _ in cols]].rename(columns=dict(cols)).reset_index(drop=True)
    if template:
        names = df["Name"].fillna("Valued Client").astype(str).str.title().reset_index(drop=True)
        out["SMS Message"] = [template.replace("{name}", n) for n in names]
    return out


def build_excel(view: pd.DataFrame, title: str, ref: date, source: str) -> bytes:
    """Excel file with Client Code / Main Code / Mobile stored as real text."""
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})
    sheet_name = re.sub(r"[\[\]\:\*\?\/\\]", "-", title)[:31]
    ws = wb.add_worksheet(sheet_name)

    head = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#5B21B6",
                          "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True})
    txt = wb.add_format({"num_format": "@", "border": 1, "border_color": "#E5E7EB"})
    gen = wb.add_format({"border": 1, "border_color": "#E5E7EB"})
    dt = wb.add_format({"num_format": "dd-mmm-yyyy", "align": "center", "border": 1, "border_color": "#E5E7EB"})
    num = wb.add_format({"num_format": "0", "align": "center", "border": 1, "border_color": "#E5E7EB"})
    text_cols = {"Client Code", "Main Code", "Mobile", "Name", "SMS Message"}
    date_cols = {"Date of Birth", "Birthday On"}

    ws.set_row(0, 24)
    for j, col in enumerate(view.columns):
        ws.write(0, j, col, head)
        series = view[col]
        longest = max([len(col)] + [len(str(x)) for x in series.head(2000).dropna()])
        width = 14 if col in date_cols else min(max(longest + 3, 12), 60)
        ws.set_column(j, j, width)
        for i, val in enumerate(series, start=1):
            if pd.isna(val):
                ws.write_blank(i, j, None, gen)
            elif col in text_cols:
                ws.write_string(i, j, str(val), txt)
            elif col in date_cols:
                ws.write_datetime(i, j, pd.Timestamp(val).to_pydatetime(), dt)
            else:
                ws.write_number(i, j, float(val), num)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(view), 1), len(view.columns) - 1)

    info = wb.add_worksheet("Summary")
    b = wb.add_format({"bold": True})
    info.set_column(0, 0, 22)
    info.set_column(1, 1, 48)
    rows = [("Report", title), ("Reference date", ref.strftime("%d-%b-%Y")),
            ("Records", len(view)), ("Source file", source),
            ("Generated", datetime.now(TZ).strftime("%d-%b-%Y %H:%M"))]
    for i, (k, v) in enumerate(rows):
        info.write(i, 0, k, b)
        info.write(i, 1, v)
    wb.close()
    return buf.getvalue()


CSS = """
<style>
.block-container {padding-top: 1.6rem; max-width: 1250px;}
.hero {background: linear-gradient(120deg,#4C1D95 0%,#7C3AED 55%,#DB2777 100%);
       padding: 1.6rem 1.9rem; border-radius: 18px; color: #fff; margin-bottom: 1.1rem;
       box-shadow: 0 10px 30px rgba(124,58,237,.25);}
.hero h1 {margin: 0; font-size: 2rem; color:#fff; padding:0;}
.hero p {margin: .35rem 0 0 0; opacity: .92; font-size: 1.02rem;}
.kpi {border: 1px solid rgba(128,128,128,.25); background: rgba(128,128,128,.07);
      border-radius: 14px; padding: .9rem 1rem; height: 100%;}
.kpi .l {font-size: .78rem; text-transform: uppercase; letter-spacing: .06em; opacity: .7;}
.kpi .v {font-size: 1.85rem; font-weight: 700; line-height: 1.2;}
.kpi .s {font-size: .78rem; opacity: .65;}
.kpi.hot {border-color: #DB2777; background: rgba(219,39,119,.10);}
.kpi.warm {border-color: #7C3AED; background: rgba(124,58,237,.10);}
.pill {display:inline-block; padding:.15rem .7rem; border-radius:99px; font-size:.82rem;
       background: rgba(124,58,237,.15); border:1px solid rgba(124,58,237,.4);}
</style>
"""


def kpi(label: str, value, sub: str = "", cls: str = "") -> str:
    return (f'<div class="kpi {cls}"><div class="l">{label}</div>'
            f'<div class="v">{value}</div><div class="s">{sub}</div></div>')


def show_table(view: pd.DataFrame, height: int | None = None):
    cfg = {
        "Date of Birth": st.column_config.DateColumn("Date of Birth", format="DD MMM YYYY"),
        "Birthday On": st.column_config.DateColumn("Birthday On", format="ddd, DD MMM YYYY"),
        "Turning Age": st.column_config.NumberColumn("Turning Age", format="%d"),
        "Days To Go": st.column_config.NumberColumn("Days To Go", format="%d"),
    }
    kw = {"height": height} if height else {}
    st.dataframe(view, hide_index=True, width="stretch", column_config=cfg, **kw)


def birthday_block(view: pd.DataFrame, title: str, ref: date, source: str, key: str, empty_msg: str):
    if view.empty:
        st.info(empty_msg)
    else:
        show_table(view, height=min(120 + 35 * len(view), 560))
    st.download_button(
        f"⬇️  Download “{title}” as Excel ({len(view)} rows)",
        data=build_excel(view, title, ref, source),
        file_name=f"{re.sub(r'[^A-Za-z0-9_-]+', '_', title)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", key=key, disabled=view.empty,
    )


# --------------------------------------------------------------------------
# App
# --------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Duelist Birthday Desk", page_icon="🎂", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="hero"><h1>🎂 Duelist Birthday Desk</h1>'
        "<p>Upload the Duelist file, see who celebrates today, tomorrow and soon — "
        "and download the call / SMS list in one click.</p></div>",
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.header("1 · Data")
        up = st.file_uploader("Duelist file", type=["xlsb", "xlsx", "xlsm", "xls", "csv"],
                              help="Binary (.xlsb), Excel, macro-enabled (.xlsm) or CSV.")
        st.header("2 · Options")
        today = datetime.now(TZ).date()
        ref = st.date_input("Today's date", value=today,
                            help="Birthdays are calculated from this date. Change it to preview another day.")
        one_row = st.checkbox("One row per client", value=True,
                              help="Clients with several loans appear once; their Main Codes are joined.")
        fallback = st.checkbox("Use second mobile if first is empty", value=True)
        window = st.slider("Upcoming window (days)", 2, 60, 7)
        st.header("3 · SMS text")
        use_sms = st.checkbox("Add SMS message column to Excel", value=True)
        template = st.text_area("Message template", DEFAULT_SMS, height=110,
                                help="{name} is replaced with the client's name.")

    if up is None:
        c1, c2, c3 = st.columns(3)
        c1.markdown(kpi("Step 1", "Upload", "Drop the Duelist .xlsb / .xlsx in the sidebar"), unsafe_allow_html=True)
        c2.markdown(kpi("Step 2", "Review", "Today, tomorrow and upcoming birthdays"), unsafe_allow_html=True)
        c3.markdown(kpi("Step 3", "Download", "Excel with text-safe codes & mobile numbers"), unsafe_allow_html=True)
        st.stop()

    data = up.getvalue()
    sheets = list_sheets(data, up.name)
    with st.sidebar:
        sheet = st.selectbox("Sheet", sheets) if len(sheets) > 1 else sheets[0]

    headers = read_headers(data, up.name, sheet)
    detected = detect_columns(headers)
    with st.sidebar.expander("Column mapping", expanded=any(detected[f] is None for f in REQUIRED)):
        mapping = {}
        for f in ALIASES:
            opts = [NONE] + headers
            idx = opts.index(detected[f]) if detected[f] in opts else 0
            pick = st.selectbox(LABELS[f], opts, index=idx, key=f"map_{f}")
            mapping[f] = None if pick == NONE else pick

    missing = [LABELS[f] for f in REQUIRED if not mapping[f]]
    if missing:
        st.error("Please map these columns in the sidebar: " + ", ".join(missing))
        st.stop()

    cols = tuple(dict.fromkeys(v for v in mapping.values() if v))
    raw = read_columns(data, up.name, sheet, cols)
    ds = build_dataset(raw, mapping, fallback, ref)
    bd = with_birthdays(ds, ref)

    tomorrow = ref + timedelta(days=1)
    today_df, tmr_df = bd[bd["Days To Go"] == 0], bd[bd["Days To Go"] == 1]
    soon_df = bd[(bd["Days To Go"] >= 0) & (bd["Days To Go"] <= window)].sort_values(["Days To Go", "Name"])

    def prep(df):
        df = df.sort_values("Name")
        if one_row:
            df = collapse_clients(df)
        return to_view(df, template if use_sms else None, show_loans=one_row)

    def prep_sorted(df):
        df = df.sort_values(["Days To Go", "Name"])
        if one_row:
            df = collapse_clients(df).sort_values(["Days To Go", "Name"])
        return to_view(df, template if use_sms else None, show_loans=one_row)

    v_today, v_tmr, v_soon = prep(today_df), prep(tmr_df), prep_sorted(soon_df)

    # KPIs
    n_clients = ds["ClientCode"].nunique()
    k = st.columns(6)
    k[0].markdown(kpi("Records", f"{len(ds):,}", "rows in file"), unsafe_allow_html=True)
    k[1].markdown(kpi("Clients", f"{n_clients:,}", "unique client codes"), unsafe_allow_html=True)
    k[2].markdown(kpi("Valid DOB", f"{ds['DOB'].notna().mean():.0%}" if len(ds) else "0%",
                      f"{ds['DOB'].notna().sum():,} records"), unsafe_allow_html=True)
    k[3].markdown(kpi("Today", len(v_today), ref.strftime("%a, %d %b"), "warm"), unsafe_allow_html=True)
    k[4].markdown(kpi("Tomorrow", len(v_tmr), tomorrow.strftime("%a, %d %b"), "hot"), unsafe_allow_html=True)
    k[5].markdown(kpi(f"Next {window} days", len(v_soon), "incl. today"), unsafe_allow_html=True)
    st.write("")

    t1, t2, t3, t4, t5, t6 = st.tabs(
        ["🎁 Tomorrow", "🎉 Today", "📅 Upcoming", "📊 Insights", "🔎 Search all", "🧹 Data quality"])

    stamp = tomorrow.strftime("%d-%b-%Y")
    with t1:
        st.markdown(f'<span class="pill">Birthdays on {tomorrow.strftime("%A, %d %B %Y")}</span>',
                    unsafe_allow_html=True)
        st.write("")
        birthday_block(v_tmr, f"Tomorrow Birthdays {stamp}", ref, up.name, "dl_tmr",
                       "No client has a birthday tomorrow.")
    with t2:
        st.markdown(f'<span class="pill">Birthdays on {ref.strftime("%A, %d %B %Y")}</span>',
                    unsafe_allow_html=True)
        st.write("")
        birthday_block(v_today, f"Today Birthdays {ref.strftime('%d-%b-%Y')}", ref, up.name, "dl_today",
                       "No client has a birthday today.")
    with t3:
        st.caption(f"{ref.strftime('%d %b')} → {(ref + timedelta(days=window)).strftime('%d %b %Y')}")
        birthday_block(v_soon, f"Upcoming {window} Days from {ref.strftime('%d-%b-%Y')}", ref, up.name,
                       "dl_soon", "No birthdays in this window.")
    with t4:
        valid = unique_clients(bd)
        if valid.empty:
            st.info("No valid dates of birth to analyse.")
        else:
            a, b = st.columns(2)
            months = valid["DOB"].dt.month.value_counts().reindex(range(1, 13), fill_value=0)
            months.index = [calendar.month_abbr[i] for i in months.index]
            a.subheader("Birthdays by month")
            a.bar_chart(months, color="#7C3AED")
            nxt30 = (valid[valid["Days To Go"] <= 30]["Days To Go"].value_counts()
                     .reindex(range(0, 31), fill_value=0))
            b.subheader("Next 30 days")
            b.bar_chart(nxt30, color="#DB2777")
            ages = (pd.Timestamp(ref).year - valid["DOB"].dt.year)
            bins = pd.cut(ages, [-1, 17, 25, 35, 45, 55, 65, 200],
                          labels=["<18", "18-25", "26-35", "36-45", "46-55", "56-65", "65+"])
            st.subheader("Age groups")
            st.bar_chart(bins.value_counts().sort_index(), color="#4C1D95")
            busiest = valid.groupby([valid["DOB"].dt.month, valid["DOB"].dt.day]).size().sort_values(ascending=False)
            (mm, dd), cnt = busiest.index[0], busiest.iloc[0]
            st.caption(f"Busiest birthday date: {dd} {calendar.month_name[mm]} with {cnt} clients.")
    with t5:
        q = st.text_input("Search name, client code, main code or mobile")
        mcol, _ = st.columns([1, 3])
        month = mcol.selectbox("Birth month", ["All"] + list(calendar.month_name)[1:])
        res = ds.copy()
        if q:
            ql = q.strip().lower()
            mask = pd.Series(False, index=res.index)
            for c in ("Name", "ClientCode", "MainCode", "Mobile"):
                mask |= res[c].fillna("").str.lower().str.contains(ql, regex=False)
            res = res[mask]
        if month != "All":
            res = res[res["DOB"].dt.month == list(calendar.month_name).index(month)]
        st.caption(f"{len(res):,} records")
        show = res[["ClientCode", "Name", "MainCode", "Mobile", "DOB"]].rename(
            columns={"ClientCode": "Client Code", "MainCode": "Main Code", "DOB": "Date of Birth"})
        show_table(show.head(5000), height=480)
        if len(show) > 5000:
            st.caption("Showing first 5,000 matches — refine the search or export for the full list.")
        if st.button("Prepare Excel of these records"):
            st.session_state["all_xlsx"] = build_excel(show, "Extracted Clients", ref, up.name)
        if "all_xlsx" in st.session_state:
            st.download_button("⬇️  Download prepared Excel", st.session_state["all_xlsx"],
                               "Duelist_Extract.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with t6:
        mob_ok = ds["Mobile"].fillna("").str.fullmatch(r"9[678]\d{8}")
        dup_codes = ds["ClientCode"].dropna().duplicated().sum()
        q1, q2, q3, q4 = st.columns(4)
        q1.metric("Missing DOB", int(ds["DOB_raw"].isna().sum()))
        q2.metric("Unreadable / implausible DOB", int(ds["DOB_issue"].sum()))
        q3.metric("Missing mobile", int(ds["Mobile"].isna().sum()))
        q4.metric("Mobile not 10-digit 96/97/98", int((~mob_ok & ds["Mobile"].notna()).sum()))
        st.caption(f"{dup_codes:,} rows repeat a client code already seen (clients with multiple loans).")
        bad = ds[ds["DOB_issue"]][["ClientCode", "Name", "DOB_raw"]].rename(columns={"DOB_raw": "DOB as in file"})
        if not bad.empty:
            st.subheader("DOB values that could not be read")
            st.dataframe(bad.head(1000), hide_index=True, width="stretch")


if __name__ == "__main__":
    main()
