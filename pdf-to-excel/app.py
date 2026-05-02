import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import warnings
warnings.filterwarnings("ignore")

BASE_URL  = "https://srv32.hiocdis.org"
LOGIN_URL = f"{BASE_URL}/stock/login.aspx"
DATA_URL  = f"{BASE_URL}/stock/BalanceSheet_total.aspx"

ARABIC_HEADERS = [
    "كود الدواء",
    "اسم الدواء",
    "الرصيد الحالي",
    "اخر سعر توريد",
    "سعر البيع للمنتفع",
    "اخر سعر متحرك",
]

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="كشف الأرصدة", page_icon="💊", layout="wide")

st.markdown("""
<style>
    .main-title { text-align:center; color:#1F4E79; font-size:2.2rem; font-weight:700; }
    .sub-title  { text-align:center; color:#666; font-size:1rem; margin-bottom:1.5rem; }
    .stDownloadButton > button {
        background:#1F4E79; color:white; border-radius:8px;
        padding:.6rem 1.8rem; font-size:1rem; font-weight:600; width:100%;
    }
    .stDownloadButton > button:hover { background:#2E75B6; }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">💊 كشف الأرصدة الدوائية</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-title">سجّل دخولك وسيتم استخراج البيانات مباشرة من الموقع وتحويلها إلى Excel</p>',
    unsafe_allow_html=True,
)


# ── Scraper ───────────────────────────────────────────────────────────────────
def login_and_fetch(username: str, password: str) -> str:
    """Login and return the HTML of the balance sheet page."""
    session = requests.Session()
    session.headers["User-Agent"] = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
    )

    # Step 1: Load login page to grab ASP.NET hidden fields
    r = session.get(LOGIN_URL, verify=False, timeout=20)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "lxml")

    def _val(name):
        el = soup.find("input", {"name": name})
        return el["value"] if el else ""

    payload = {
        "__EVENTTARGET":        "",
        "__EVENTARGUMENT":      "",
        "__VIEWSTATE":          _val("__VIEWSTATE"),
        "__VIEWSTATEGENERATOR": _val("__VIEWSTATEGENERATOR"),
        "__EVENTVALIDATION":    _val("__EVENTVALIDATION"),
        "txtuser":              username,
        "txtpass":              password,
        "btnLogin":             "Login",
    }

    # Step 2: POST credentials
    r2 = session.post(LOGIN_URL, data=payload, verify=False, timeout=20)
    r2.raise_for_status()

    # Check if still on login page (login failed)
    if "txtuser" in r2.text or "txtpass" in r2.text:
        raise ValueError("فشل تسجيل الدخول — تحقق من اسم المستخدم وكلمة المرور")

    # Step 3: Fetch the balance sheet page
    r3 = session.get(DATA_URL, verify=False, timeout=60)
    r3.raise_for_status()
    return r3.text


def parse_table(html: str) -> pd.DataFrame:
    """Extract the data table from HTML and return a DataFrame."""
    soup  = BeautifulSoup(html, "lxml")
    table = soup.find("table")
    if table is None:
        # Try finding any table-like structure
        tables = soup.find_all("table")
        if not tables:
            raise ValueError("لم يتم العثور على جدول في صفحة الموقع")
        # Pick the largest table
        table = max(tables, key=lambda t: len(t.find_all("tr")))

    rows = []
    for tr in table.find_all("tr"):
        cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
        if cells:
            rows.append(cells)

    if not rows:
        raise ValueError("الجدول موجود لكنه فارغ")

    # Detect header row and drop it
    data_rows = []
    for row in rows:
        # Skip header rows (contain Arabic column names)
        row_text = " ".join(row)
        if any(h in row_text for h in ["كود", "اسم", "رصيد", "سعر"]):
            continue
        # Skip empty rows
        if not any(row):
            continue
        # Keep rows that have at least 4 non-empty cells
        non_empty = [c for c in row if c]
        if len(non_empty) < 4:
            continue
        data_rows.append(row)

    if not data_rows:
        raise ValueError("لا توجد صفوف بيانات في الجدول")

    # Normalize to exactly 6 columns
    normalized = []
    for row in data_rows:
        # Pad or trim to 6
        row = (row + [""] * 6)[:6]
        normalized.append(row)

    df = pd.DataFrame(normalized, columns=ARABIC_HEADERS)

    # Convert numeric columns
    for col in ARABIC_HEADERS[2:]:   # columns 3-6 are numeric
        df[col] = pd.to_numeric(df[col].str.replace(",", ""), errors="coerce").fillna(0)

    return df


# ── Excel generator ───────────────────────────────────────────────────────────
def generate_excel(df: pd.DataFrame) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Balance Sheet"
    ws.sheet_view.rightToLeft = True

    thin   = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    a_fill = PatternFill(start_color="EBF3FB", end_color="EBF3FB", fill_type="solid")
    h_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)

    # Header
    for c, h in enumerate(ARABIC_HEADERS, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = h_font
        cell.fill = h_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    # Data
    for r_i, row in enumerate(df.itertuples(index=False), 2):
        for c_i, val in enumerate(row, 1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center" if c_i != 2 else "right",
                vertical="center",
                wrap_text=(c_i == 2),
            )
            if r_i % 2 == 0:
                cell.fill = a_fill
            if c_i in (3, 4, 5, 6):
                cell.number_format = "General"

    # Totals row
    total_row = len(df) + 2
    tc = ws.cell(row=total_row, column=1, value="الإجمالي")
    tc.font = Font(bold=True)
    for c_i in (3, 4, 5, 6):
        col_letter = ["C", "D", "E", "F"][c_i - 3]
        cell = ws.cell(
            row=total_row, column=c_i,
            value=f"=SUM({col_letter}2:{col_letter}{total_row-1})",
        )
        cell.font = Font(bold=True)
        cell.number_format = "General"
        cell.fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
        cell.border = border

    # Column widths
    for col, w in zip("ABCDEF", [16, 65, 14, 18, 20, 18]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── UI ────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("🔐 تسجيل الدخول")
    username = st.text_input("اسم المستخدم", placeholder="username")
    password = st.text_input("كلمة المرور", type="password", placeholder="••••••")
    fetch_btn = st.button("🚀 جلب البيانات", use_container_width=True)
    st.markdown("---")
    st.caption(f"🌐 المصدر: {DATA_URL}")

if fetch_btn:
    if not username or not password:
        st.error("⚠️ يرجى إدخال اسم المستخدم وكلمة المرور")
    else:
        try:
            with st.spinner("🔐 جارٍ تسجيل الدخول..."):
                html = login_and_fetch(username, password)

            with st.spinner("📊 جارٍ استخراج البيانات من الجدول..."):
                df = parse_table(html)

            # ── Stats ─────────────────────────────────────────────────────
            c1, c2, c3 = st.columns(3)
            c1.metric("📦 عدد الأصناف",   f"{len(df):,}")
            c2.metric("💰 إجمالي الرصيد", f"{int(df[ARABIC_HEADERS[2]].sum()):,}")
            c3.metric("📅 تاريخ الجلب",   pd.Timestamp.now().strftime("%Y-%m-%d"))

            # ── Preview ───────────────────────────────────────────────────
            st.subheader("🔍 معاينة البيانات")
            st.dataframe(df, use_container_width=True, height=420)

            # ── Download ──────────────────────────────────────────────────
            st.subheader("⬇️ تنزيل ملف Excel")
            fname = f"BalanceSheet_{pd.Timestamp.now().strftime('%Y-%m-%d')}.xlsx"
            st.download_button(
                label="📥 تنزيل ملف Excel",
                data=generate_excel(df),
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.success(f"✅ تم استخراج **{len(df):,}** صنف بنجاح!")

        except ValueError as e:
            st.error(f"❌ {e}")
        except requests.exceptions.ConnectionError:
            st.error("❌ لا يمكن الاتصال بالموقع — تحقق من اتصال الإنترنت")
        except requests.exceptions.Timeout:
            st.error("❌ انتهت مهلة الاتصال — حاول مرة أخرى")
        except Exception as e:
            st.error(f"❌ خطأ غير متوقع: {e}")

else:
    st.markdown("""
    ---
    ### 📋 كيفية الاستخدام
    1. أدخل **اسم المستخدم** و**كلمة المرور** في الشريط الجانبي
    2. اضغط **جلب البيانات**
    3. راجع البيانات في جدول المعاينة
    4. اضغط **تنزيل ملف Excel** 📥
    ---
    > 🔄 يمكنك تشغيل هذا التطبيق يومياً للحصول على أحدث البيانات تلقائياً
    """)

st.markdown(
    "<p style='text-align:center;color:#bbb;font-size:.8rem;margin-top:2rem;'>"
    "Balance Sheet Scraper | كشف الأرصدة الدوائية</p>",
    unsafe_allow_html=True,
)
