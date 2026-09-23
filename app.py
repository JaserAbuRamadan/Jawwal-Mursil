import re
import streamlit as st
import pandas as pd

# Page Configuration & Modern Dark Theme Styling
st.set_page_config(page_title="Data Extractor Pro", page_icon="⚡", layout="centered")

st.markdown("""
    <style>
    .main {
        background-color: #0f172a;
    }
    .card {
        background-color: #1e293b;
        border: 1px solid #334155;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 15px;
    }
    .stButton>button {
        background-color: #3b82f6;
        color: white;
        font-weight: bold;
        border-radius: 6px;
        border: none;
        width: 100%;
        padding: 10px;
    }
    .stButton>button:hover {
        background-color: #2563eb;
    }
    </style>
""", unsafe_allow_html=True)

st.title("⚡ Excel Data Segmenter Pro")
st.markdown("<p style='color: #94a3b8;'>Upload your Excel file to automatically clean, filter, and sort your mobile numbers.</p>", unsafe_allow_html=True)


# ---------- Helper functions ----------

def normalize_mobile(raw) -> str | None:
    """
    Turn any of these into the same normalized LOCAL number:
      +970599000000, 970599000000, 0599000000, 599000000  ->  0599000000
    Assumption: the real subscriber number is always the LAST 9 digits,
    and we re-add a leading 0 for a consistent local-format output.
    Returns None if there aren't at least 9 digits (can't be a valid mobile).
    """
    if pd.isna(raw):
        return None
    # Excel often stores numbers as floats (599000000 -> 599000000.0).
    # Cast whole-number floats to int first so we don't pick up a stray
    # trailing digit from the ".0".
    if isinstance(raw, float) and raw.is_integer():
        raw = int(raw)
    digits = re.sub(r"\D", "", str(raw))
    if len(digits) < 9:
        return None
    local9 = digits[-9:]
    return "0" + local9


def normalize_label(raw) -> str:
    """
    Collapse label variants into 'yes' / 'no' / 'unknown'.
    Handles: 'Yes', 'USSD', 'No', 'No Login Before', 'no login', etc.
    """
    if pd.isna(raw):
        return "unknown"
    t = str(raw).strip().lower()
    if "yes" in t or "ussd" in t:
        return "yes"
    if t.startswith("no"):
        return "no"
    return "unknown"


def parse_numeric(raw) -> float:
    """Strip currency symbols/commas/spaces before converting to a number."""
    if pd.isna(raw):
        return 0.0
    cleaned = re.sub(r"[^\d.\-]", "", str(raw))
    try:
        return float(cleaned) if cleaned not in ("", "-", ".") else 0.0
    except ValueError:
        return 0.0


# ---------- 1. File Upload Card ----------
st.markdown('<div class="card">', unsafe_allow_html=True)
st.subheader("📁 1. Source File")
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

selected_sheet = None
df_preview = None
if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        selected_sheet = st.selectbox("Target Sheet", sheet_names)
        if selected_sheet:
            df_preview = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
    except Exception as e:
        st.error(f"Could not read sheets: {e}")
st.markdown('</div>', unsafe_allow_html=True)

# ---------- 2. Column Mapping Card ----------
mobile_col = numeric_col = text_col = None
if df_preview is not None:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("🧭 2. Map Your Columns")
    st.markdown("<p style='color:#94a3b8; font-size:13px;'>Column order isn't always the same, so pick them manually.</p>", unsafe_allow_html=True)

    columns = list(df_preview.columns)

    def guess_index(keywords, default=0):
        for i, c in enumerate(columns):
            if any(k in str(c).lower() for k in keywords):
                return i
        return default

    mobile_col = st.selectbox("📱 Mobile Number column", columns,
                               index=guess_index(["mobile", "phone", "number"], 0))
    numeric_col = st.selectbox("💰 CashIn / Numeric column", columns,
                                index=guess_index(["cash", "amount", "value"], min(1, len(columns) - 1)))
    text_col = st.selectbox("🏷️ Status / Label column (Yes / No / USSD)", columns,
                             index=guess_index(["status", "login", "label"], min(2, len(columns) - 1)))

    # Replace NaN with empty strings for display only — a raw NaN in the
    # preview table can crash Streamlit's frontend JSON serializer.
    safe_preview = df_preview.head(5).where(pd.notnull(df_preview.head(5)), "")
    st.dataframe(safe_preview, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ---------- 3. Process Button & Results ----------
if df_preview is not None and mobile_col and numeric_col and text_col:
    if st.button("🚀 Process & Categorize Data"):
        try:
            df = df_preview.copy()

            processed_df = pd.DataFrame({
                "Mobile": df[mobile_col].apply(normalize_mobile),
                "NumericVal": df[numeric_col].apply(parse_numeric),
                "TextVal": df[text_col].apply(normalize_label),
            }).dropna(subset=["Mobile"])

            g1 = processed_df[processed_df["TextVal"] == "yes"]["Mobile"]
            g2 = processed_df[(processed_df["TextVal"] == "no") & (processed_df["NumericVal"] < 20)]["Mobile"]
            g3 = processed_df[(processed_df["TextVal"] == "no") & (processed_df["NumericVal"] >= 20)]["Mobile"]
            unmatched = processed_df[processed_df["TextVal"] == "unknown"]

            st.success("Processing complete!")

            if len(unmatched) > 0:
                st.warning(f"⚠️ {len(unmatched)} rows had a label that wasn't recognized as Yes/USSD/No and were skipped. "
                           f"Examples: {df[text_col].dropna().astype(str).unique()[:10].tolist()}")

            groups = [
                ("Group 1: 'Yes' or 'USSD'", g1),
                ("Group 2: 'No' AND CashIn < 20", g2),
                ("Group 3: 'No' AND CashIn >= 20", g3),
            ]

            for title, group_data in groups:
                text_result = ','.join(group_data)
                st.markdown(f"""
                    <div class="card">
                        <h4 style="color: #f8fafc; margin-top: 0;">{title}</h4>
                        <p style="color: #94a3b8; font-size: 14px;">Total Items: <b>{len(group_data)}</b></p>
                    </div>
                """, unsafe_allow_html=True)
                st.text_area(f"Copy {title}", text_result, height=80, key=title)

        except Exception as e:
            st.error(f"Processing Error: {e}")
elif df_preview is not None:
    st.info("Select all three columns above to enable processing.")
