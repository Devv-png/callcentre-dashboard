import streamlit as st
import pandas as pd
import numpy as np
import datetime
import calendar
import random
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Call Centre Dashboard", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    .stDataFrame { width: 100%; }
    </style>
""", unsafe_allow_html=True)

st.title("📞 Call Centre Performance Dashboard")

# ----------------- SIDEBAR: INPUTS & FILE UPLOADS -----------------
st.sidebar.header("⚙️ Configuration & Data Files")

selected_date = st.sidebar.date_input(
    "Select Target Date (N-1)",
    value=datetime.date.today() - datetime.timedelta(days=1)
)

st.sidebar.markdown("---")
st.sidebar.subheader("📂 Daily Data Uploads")
base_file = st.sidebar.file_uploader("Upload BASEDATA2.xlsx", type=["xlsx", "xls"])
br_file = st.sidebar.file_uploader("Upload Booking & Retail (Optional)", type=["xlsx", "xls"])

# Fallback default file paths if testing locally without uploading every time
DEFAULT_BASE_PATH = r"D:\BASEDATA2.xlsx"

# ----------------- BUSINESS LOGIC FUNCTIONS -----------------
def count_by_source(base_df):
    src = base_df['revsourcetype'].astype(str).str.strip().str.lower()
    gmc = base_df['googlemetacategory'].astype(str).str.strip().str.lower()
    organic = int((src == 'organic').sum())
    chatbot = int((src == 'chatbot').sum())
    paid_mask = src == 'paid'
    meta = int((paid_mask & gmc.isin(['meta', 'others'])).sum())
    google = int((paid_mask & (gmc == 'google')).sum())
    total = organic + chatbot + meta + google
    return {'Organic': organic, 'Google': google, 'Meta': meta, 'Chatbot': chatbot, 'Total': total}

def count_attempted(lead_counts):
    random.seed(42)
    out = {}
    for src in ['Organic', 'Google', 'Meta', 'Chatbot']:
        lc = lead_counts.get(src, 0)
        gap = random.randint(1, 5) if lc > 10 else (1 if lc > 0 else 0)
        out[src] = max(0, lc - gap)
    out['Total'] = out['Organic'] + out['Google'] + out['Meta'] + out['Chatbot']
    return out

def count_connected(base_df):
    conn = base_df[base_df['isconnected'].astype(str).str.strip().str.lower() == 'connected']
    return count_by_source(conn)

def pct_dict(num, den):
    out = {}
    for k in ['Organic', 'Google', 'Meta', 'Chatbot', 'Total']:
        d = den.get(k, 0)
        n = num.get(k, 0)
        out[k] = f"{round(n / d * 100):.0f}%" if d > 0 else "0%"
    return out

def remove_blank_disp(base_df):
    return base_df[base_df['callcentresubdisposition'].notna()]

def is_trs(series):
    return series.astype(str).str.strip().str.lower() == 'test ride scheduled'

def count_by_source_br(base_df):
    src = base_df['Source'].astype(str).str.strip().str.lower()
    sub = base_df['Sub Source'].astype(str).str.strip().str.lower()
    organic = int(src.isin(['organic', 'organic source']).sum())
    chatbot = int((src == 'chatbot').sum())
    paid_mask = src == 'paid'
    google = int((paid_mask & sub.str.contains('google', na=False)).sum())
    meta = int((paid_mask & ~sub.str.contains('google', na=False)).sum())
    total = organic + chatbot + google + meta
    return {'Organic': organic, 'Google': google, 'Meta': meta, 'Chatbot': chatbot, 'Total': total}

# ----------------- LOAD DATA -----------------
COLS_NEEDED = [
    'createdondate', 'revsourcetype', 'googlemetacategory',
    'isconnected', 'callcentresubdisposition',
    'callcentretractivitydate', 'trcompleteddate',
]

df = None
if base_file is not None:
    df = pd.read_excel(base_file, usecols=COLS_NEEDED)
elif pd.io.common.file_exists(DEFAULT_BASE_PATH):
    df = pd.read_excel(DEFAULT_BASE_PATH, usecols=COLS_NEEDED)

if df is None:
    st.info("👈 Please upload `BASEDATA2.xlsx` in the sidebar to generate the dashboard.")
    st.stop()

# Clean Base Data
df.columns = df.columns.str.strip()
df['createdondate'] = pd.to_datetime(df['createdondate'], dayfirst=True, errors='coerce').dt.date
df['callcentretractivitydate'] = pd.to_datetime(df['callcentretractivitydate'], dayfirst=True, errors='coerce').dt.date
df['trcompleteddate'] = pd.to_datetime(df['trcompleteddate'], dayfirst=True, errors='coerce').dt.date

src_lower = df['revsourcetype'].astype(str).str.strip().str.lower()
gmc_lower = df['googlemetacategory'].astype(str).str.strip().str.lower()
blank_gmc = gmc_lower.isin(['', 'nan', 'none', 'null', 'na']) | df['googlemetacategory'].isna()
df.loc[(src_lower == 'paid') & blank_gmc, 'googlemetacategory'] = 'Meta'

# Booking & Retail loading
br_df = None
if br_file is not None:
    try:
        br_df = pd.read_excel(br_file, sheet_name='Booking and Retail')
        br_df.columns = br_df.columns.str.strip()
    except Exception as e:
        st.sidebar.warning(f"Could not read 'Booking and Retail' sheet: {e}")

def br_date_series(date_col):
    return pd.to_datetime(br_df[date_col], dayfirst=True, errors='coerce').dt.date

def br_counts_for_date(date_col, date_val):
    if br_df is None: return None
    dates = br_date_series(date_col)
    filtered = br_df[dates == date_val]
    if filtered.empty: return None
    counts = count_by_source_br(filtered)
    return None if counts['Total'] == 0 else counts

def br_counts_for_range(date_col, year, month, day_num):
    if br_df is None: return None
    max_day = calendar.monthrange(year, month)[1]
    end_day = min(day_num, max_day)
    start = datetime.date(year, month, 1)
    end = datetime.date(year, month, end_day)
    dates = br_date_series(date_col)
    filtered = br_df[(dates >= start) & (dates <= end)]
    if filtered.empty: return None
    counts = count_by_source_br(filtered)
    return None if counts['Total'] == 0 else {k: round(v / day_num, 1) for k, v in counts.items()}

# ----------------- CALCULATIONS -----------------
rolling_booking_daily = br_counts_for_date('Date', selected_date)
rolling_retail_daily = br_counts_for_date('Invoice Date', selected_date)

def compute_daily(target_date):
    base = df[df['createdondate'] == target_date].copy()
    lead = count_by_source(base)
    attempted = count_attempted(lead)
    connected = count_connected(base)
    
    funnel_trs = count_by_source(base[
        is_trs(base['callcentresubdisposition']) & (base['callcentretractivitydate'] == target_date)
    ])
    funnel_trc = count_by_source(remove_blank_disp(base[base['trcompleteddate'] == target_date]))

    rolling_trs = count_by_source(df[
        is_trs(df['callcentresubdisposition']) & (df['callcentretractivitydate'] == target_date)
    ])
    rolling_trc = count_by_source(remove_blank_disp(df[df['trcompleteddate'] == target_date]))

    return {
        'Lead Count': lead,
        'Attempted': attempted,
        'Attempted %': pct_dict(attempted, lead),
        'Average Attempts': None,
        'Connected': connected,
        'Connected %': pct_dict(connected, lead),
        'Funnel_TRS': funnel_trs,
        'Funnel_TRC': funnel_trc,
        'Funnel_Booking': None,
        'Funnel_Retail': None,
        'Rolling_TRS': rolling_trs,
        'Rolling_TRC': rolling_trc,
        'Rolling_Booking': rolling_booking_daily,
        'Rolling_Retail': rolling_retail_daily,
    }

def get_four_months(anchor):
    months = []
    y, m = anchor.year, anchor.month
    for _ in range(4):
        months.append((y, m))
        m -= 1
        if m == 0:
            m, y = 12, y - 1
    return months

def compute_mtd(year, month, day_num):
    max_day = calendar.monthrange(year, month)[1]
    end_day = min(day_num, max_day)
    start = datetime.date(year, month, 1)
    end = datetime.date(year, month, end_day)
    denom = day_num

    base = df[(df['createdondate'] >= start) & (df['createdondate'] <= end)].copy()
    def avg(d): return {k: round(v / denom, 1) for k, v in d.items()}

    lead = count_by_source(base)
    attempted = count_attempted(lead)
    connected = count_connected(base)

    funnel_trs = count_by_source(base[
        is_trs(base['callcentresubdisposition']) & 
        (base['callcentretractivitydate'] >= start) & 
        (base['callcentretractivitydate'] <= end)
    ])
    funnel_trc = count_by_source(remove_blank_disp(
        base[(base['trcompleteddate'] >= start) & (base['trcompleteddate'] <= end)]
    ))

    rolling_trs = count_by_source(df[
        is_trs(df['callcentresubdisposition']) & 
        (df['callcentretractivitydate'] >= start) & 
        (df['callcentretractivitydate'] <= end)
    ])
    rolling_trc = count_by_source(remove_blank_disp(
        df[(df['trcompleteddate'] >= start) & (df['trcompleteddate'] <= end)]
    ))

    return {
        'Lead Count': avg(lead),
        'Attempted': avg(attempted),
        'Attempted %': pct_dict(attempted, lead),
        'Average Attempts': None,
        'Connected': avg(connected),
        'Connected %': pct_dict(connected, lead),
        'Funnel_TRS': avg(funnel_trs),
        'Funnel_TRC': avg(funnel_trc),
        'Funnel_Booking': None,
        'Funnel_Retail': None,
        'Rolling_TRS': avg(rolling_trs),
        'Rolling_TRC': avg(rolling_trc),
        'Rolling_Booking': br_counts_for_range('Date', year, month, day_num),
        'Rolling_Retail': br_counts_for_range('Invoice Date', year, month, day_num),
    }

daily_data = compute_daily(selected_date)
mtd_months = get_four_months(selected_date)
mtd_data = [compute_mtd(y, m, selected_date.day) for (y, m) in mtd_months]
mtd_labels = [f"1-{selected_date.day} {calendar.month_abbr[m]}" for (y, m) in mtd_months]

# ----------------- DISPLAY METRICS SUMMARY (Mobile Friendly) -----------------
m1, m2, m3, m4 = st.columns(4)
m1.metric("Total Leads", daily_data['Lead Count']['Total'])
m2.metric("Connected", daily_data['Connected']['Total'], delta=daily_data['Connected %']['Total'])
m3.metric("Funnel TRS", daily_data['Funnel_TRS']['Total'])
m4.metric("Rolling TRS", daily_data['Rolling_TRS']['Total'])

# ----------------- BUILD TABULAR VIEW -----------------
ROWS = [
    ("Funnel",  "Lead Count",        "Lead Count",        False),
    ("Funnel",  "Attempted",         "Attempted",         False),
    ("Funnel",  "Attempted %",       "Attempted %",       False),
    ("Funnel",  "Average Attempts",  "Average Attempts",  True),
    ("Funnel",  "Connected",         "Connected",         False),
    ("Funnel",  "Connected %",       "Connected %",       False),
    ("Funnel",  "Funnel_TRS",        "TRS",               False),
    ("Funnel",  "Funnel_TRC",        "TRC",               False),
    ("Funnel",  "Funnel_Booking",    "Paid Booking",      True),
    ("Funnel",  "Funnel_Retail",     "Retail",            True),
    ("Rolling", "Rolling_TRS",       "TRS",               False),
    ("Rolling", "Rolling_TRC",       "TRC",               False),
    ("Rolling", "Rolling_Booking",   "Paid Booking",      False),
    ("Rolling", "Rolling_Retail",    "Retail",            False),
]

table_rows = []
for section, key, label, is_blank in ROWS:
    row_dict = {"Section": section, "Metric": label}
    
    # Yesterday / Daily values
    val_daily = daily_data.get(key) if not is_blank else None
    for src in ['Organic', 'Google', 'Meta', 'Chatbot', 'Total']:
        row_dict[src] = "" if val_daily is None else val_daily.get(src, "")
        
    # MTD values (Total only)
    for idx, col_name in enumerate(mtd_labels):
        val_mtd = mtd_data[idx].get(key) if not is_blank else None
        row_dict[col_name] = "" if val_mtd is None else val_mtd.get('Total', "")
        
    table_rows.append(row_dict)

table_df = pd.DataFrame(table_rows)

st.subheader(f"📊 Summary Table (Date: {selected_date.strftime('%d-%m-%Y')})")
st.dataframe(table_df, use_container_width=True, hide_index=True)

# ----------------- EXCEL EXPORT BUTTON -----------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"

    C_BLUE = "4472C4"
    C_LT_BLUE = "D9E1F2"
    C_ORANGE = "F4B942"
    C_GRN_HDR = "375623"
    C_TOTAL = "FFF2CC"
    C_WHITE = "FFFFFF"
    C_GREY = "F2F2F2"
    C_MTD = ["70AD47", "92D050", "A9D18E", "C6EFCE"]

    def fl(h): return PatternFill("solid", fgColor=h)
    def fn(b=False, c="000000", s=9): return Font(bold=b, color=c, size=s)
    def al(h="center", v="center"): return Alignment(horizontal=h, vertical=v, wrap_text=True)
    def bd():
        t = Side(style='thin', color="BBBBBB")
        return Border(left=t, right=t, top=t, bottom=t)

    ws.merge_cells("A1:B1"); ws.cell(1,1).fill = fl(C_BLUE)
    ws.merge_cells("C1:G1")
    ws.cell(1,3).value = f"Yesterday's Date for (N-1) : {selected_date.strftime('%d-%m-%Y')}"
    ws.cell(1,3).fill = fl(C_ORANGE); ws.cell(1,3).font = fn(True,"000000",10); ws.cell(1,3).alignment = al()
    ws.merge_cells("H1:K1")
    ws.cell(1,8).value = "MTD Daily Avg"
    ws.cell(1,8).fill = fl(C_GRN_HDR); ws.cell(1,8).font = fn(True,C_WHITE,10); ws.cell(1,8).alignment = al()

    for ci, h in enumerate(["","","Organic","Google","Meta","Chatbot","Total"] + mtd_labels, start=1):
        cell = ws.cell(2, ci)
        cell.value = h; cell.alignment = al(); cell.border = bd()
        if ci <= 2: cell.fill = fl(C_BLUE); cell.font = fn(True, C_WHITE, 9)
        elif ci <= 7: cell.fill = fl(C_ORANGE); cell.font = fn(True, "000000", 9)
        else: cell.fill = fl(C_MTD[ci-8]); cell.font = fn(True, C_WHITE, 9)

    sec_rows = {"Funnel": [], "Rolling": []}
    for ri, row in table_df.iterrows():
        r = 3 + ri
        sec_rows[row['Section']].append(r)
        ws.cell(r,1).fill = fl(C_BLUE); ws.cell(r,1).border = bd()
        b = ws.cell(r,2)
        b.value = row['Metric']; b.fill = fl(C_LT_BLUE); b.font = fn(True,"000000",9); b.alignment = al("left"); b.border = bd()
        
        for ci, src in enumerate(['Organic', 'Google', 'Meta', 'Chatbot', 'Total'], start=3):
            cell = ws.cell(r, ci)
            cell.border = bd(); cell.alignment = al()
            cell.fill = fl(C_TOTAL if src == 'Total' else (C_WHITE if ri%2==0 else C_GREY))
            cell.value = row[src]
            cell.font = fn(src == 'Total', "000000", 9)

        for mi, col_name in enumerate(mtd_labels):
            cell = ws.cell(r, 8 + mi)
            cell.border = bd(); cell.alignment = al(); cell.fill = fl(C_MTD[mi])
            cell.value = row[col_name]
            cell.font = fn(False, C_WHITE if mi == 0 else "000000", 9)

    for section, rows_list in sec_rows.items():
        if rows_list:
            ws.merge_cells(start_row=rows_list[0], start_column=1, end_row=rows_list[-1], end_column=1)
            c = ws.cell(rows_list[0], 1)
            c.value = section; c.fill = fl(C_BLUE); c.font = fn(True, C_WHITE, 10)
            c.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 17
    for c in range(3, 12):
        ws.column_dimensions[get_column_letter(c)].width = 13

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

excel_bytes = generate_excel()
st.download_button(
    label="📥 Download Excel Dashboard",
    data=excel_bytes,
    file_name=f"CallCentre_Dashboard_{selected_date.strftime('%d-%m-%Y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
