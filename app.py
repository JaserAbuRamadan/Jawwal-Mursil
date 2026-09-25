import re
import urllib.parse
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


def chunk_list(items, size):
    """Split a list into batches so a single sms: link doesn't get too long."""
    items = list(items)
    return [items[i:i + size] for i in range(0, len(items), size)]


def build_sms_link(numbers_batch, message) -> str:
    """
    Build an sms: URI that opens the phone's default Messages app with the
    given recipients and message pre-filled (Android-style '?body=' syntax).
    NOTE: MIUI's Messaging app (com.android.mms) does not split multiple
    recipients on a comma the way stock Android does - it treats the whole
    comma-joined string as one recipient. Semicolon works as the separator
    on MIUI instead.
    """
    numbers_str = ";".join(numbers_batch)
    body = urllib.parse.quote(message)
    return f"sms:{numbers_str}?body={body}"


# ---------- Regex-based auto column detection ----------
# Content is scanned FIRST (across the whole column, not just a sample) and
# does the real deciding. Header names only add a small confidence bonus —
# useful as a tiebreaker, but they never override what the data itself shows.

MOBILE_HEADER_RE = re.compile(r"mobile|phone|msisdn|contact|رقم|جوال|هاتف", re.IGNORECASE)
NUMERIC_HEADER_RE = re.compile(r"cash|amount|value|balance|price|رصيد|مبلغ", re.IGNORECASE)
LABEL_HEADER_RE = re.compile(r"status|login|label|ussd|type|حالة", re.IGNORECASE)

MOBILE_VALUE_RE = re.compile(r"^\+?\d[\d\s\-]{7,}$")
NUMERIC_VALUE_RE = re.compile(r"^[+\-]?[\d.,\s$₪]+$")
LABEL_VALUE_RE = re.compile(r"\b(yes|no|ussd)\b", re.IGNORECASE)

HEADER_BONUS = 0.2  # small nudge, not a decider
MIN_CONFIDENCE = 0.15  # below this, don't trust the guess at all


def _digit_count(v) -> int:
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return len(re.sub(r"\D", "", str(v)))


def _mobile_content_score(series: pd.Series) -> float:
    """
    A column only really looks like mobile numbers if, once normalized,
    the last-9-digit local number starts with a real Jawwal/Ooredoo prefix
    (5xx). Plain "long digit string" isn't enough — that also matches ID
    numbers, so shape alone can't tell them apart.
    """
    vals = series.dropna()
    if vals.empty:
        return 0.0
    total = len(vals)
    shape_hits = 0
    prefix_hits = 0
    for v in vals:
        norm = normalize_mobile(v)
        if norm is None:
            continue
        shape_hits += 1
        if norm[1] == "5":  # local9 starts with 5 -> 059x/056x style prefix
            prefix_hits += 1
    if shape_hits == 0:
        return 0.0
    shape_score = shape_hits / total
    prefix_score = prefix_hits / shape_hits
    # Prefix match is what actually distinguishes a phone number from an ID;
    # shape alone (just "digits, 8+ long") is weighted low on its own.
    return shape_score * (0.15 + 0.85 * prefix_score)


def _numeric_content_score(series: pd.Series) -> float:
    non_null = series.dropna()
    if non_null.empty:
        return 0.0
    if pd.api.types.is_numeric_dtype(series):
        shape_score = 1.0
    else:
        vals = non_null.astype(str).str.strip()
        shape_score = vals.apply(lambda v: bool(NUMERIC_VALUE_RE.match(v))).mean()
    nonzero_frac = (non_null.apply(parse_numeric) != 0).mean()
    # Cash/amount columns are typically short numbers (a handful of digits).
    # Long digit runs (9-13 digits) are almost always phone numbers or IDs,
    # not amounts, so penalize length heavily instead of rewarding "nonzero".
    avg_digits = non_null.apply(_digit_count).mean()
    length_score = max(0.0, min(1.0, 1 - (avg_digits - 6) / 6))
    return shape_score * (0.3 + 0.7 * nonzero_frac) * length_score


def _label_content_score(series: pd.Series) -> float:
    vals = series.dropna().astype(str).str.strip()
    if vals.empty:
        return 0.0
    return vals.apply(lambda v: bool(LABEL_VALUE_RE.search(v))).mean()


def auto_detect_columns(df: pd.DataFrame, columns: list):
    """
    Scan every column's actual values (whole column, not a sample) and score
    how well it fits each role. Header name match adds a small bonus on top
    of the content score. Then assign roles to columns so no two roles claim
    the same column, giving priority to whichever role has the clearest match.
    Returns (mobile_idx, numeric_idx, label_idx), falling back to positional
    defaults (0, 1, 2) when nothing scores above MIN_CONFIDENCE.
    """
    scores = {}
    for i, c in enumerate(columns):
        series = df[c]
        header = str(c)
        m = _mobile_content_score(series)
        n = _numeric_content_score(series)
        l = _label_content_score(series)
        if MOBILE_HEADER_RE.search(header):
            m = min(1.0, m + HEADER_BONUS)
        if NUMERIC_HEADER_RE.search(header):
            n = min(1.0, n + HEADER_BONUS)
        if LABEL_HEADER_RE.search(header):
            l = min(1.0, l + HEADER_BONUS)
        scores[i] = {"mobile": m, "numeric": n, "label": l}

    defaults = {"mobile": 0, "numeric": min(1, len(columns) - 1), "label": min(2, len(columns) - 1)}
    roles_by_confidence = sorted(
        ["mobile", "numeric", "label"],
        key=lambda r: -max(scores[i][r] for i in scores),
    )

    assigned = {}
    used = set()
    for role in roles_by_confidence:
        best_i, best_score = None, MIN_CONFIDENCE
        for i in scores:
            if i in used:
                continue
            if scores[i][role] > best_score:
                best_score, best_i = scores[i][role], i
        if best_i is None:
            best_i = defaults[role] if defaults[role] not in used else next(
                (j for j in range(len(columns)) if j not in used), 0
            )
        assigned[role] = best_i
        used.add(best_i)

    return assigned["mobile"], assigned["numeric"], assigned["label"]


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
            # Blank header cells in the Excel file can come through as NaN
            # column names, which crashes Streamlit's dataframe display
            # (it can't JSON-serialize a NaN used as a column name/key).
            df_preview.columns = [
                str(c) if pd.notna(c) else f"Column_{i}"
                for i, c in enumerate(df_preview.columns)
            ]
    except Exception as e:
        st.error(f"Could not read sheets: {e}")
st.markdown('</div>', unsafe_allow_html=True)

# ---------- 2. Column Mapping Card ----------
mobile_col = numeric_col = text_col = None
if df_preview is not None:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("🧭 2. Map Your Columns")
    st.markdown("<p style='color:#94a3b8; font-size:13px;'>Auto-detected from header names and cell contents — override manually if needed.</p>", unsafe_allow_html=True)

    columns = list(df_preview.columns)

    mobile_idx, numeric_idx, label_idx = auto_detect_columns(df_preview, columns)

    mobile_col = st.selectbox("📱 Mobile Number column", columns, index=mobile_idx)
    numeric_col = st.selectbox("💰 CashIn / Numeric column", columns, index=numeric_idx)
    text_col = st.selectbox("🏷️ Status / Label column (Yes / No / USSD)", columns, index=label_idx)

    # Cast everything to plain strings for the preview table. Mixed-type
    # object columns (e.g. some rows numeric, some text/blank) crash
    # Streamlit's Arrow conversion otherwise ("Expected bytes, got a float").
    safe_preview = df_preview.head(5).where(pd.notnull(df_preview.head(5)), "").astype(str)
    st.dataframe(safe_preview, width="stretch")
    st.markdown('</div>', unsafe_allow_html=True)

# ---------- 3. Process Button & Results ----------
# ---------- Hardcoded group messages ----------
GROUP_MESSAGES = {
    "Group 1: 'Yes' or 'USSD'": (
        "يعطيك العافية،معك جاسر من شركة جوال.\n"
        "يرجى شحن محفظتك ب 20 شيكل اليوم او في اسرع وقت.\n"
        "لضمان استمرار خدمة جوال بي\n"
        "شكراً لتعاونك."
    ),
    "Group 2: 'No' AND CashIn < 20": (
        "يعطيك العافية،معك جاسر من شركة جوال\n"
        "يرجى شحن محفظتك في اسرع وقت ب 20 شيكل\n"
        "وتفعيل خدمة ال USSD كود\n"
        "*110#\n"
        "لضمان استمرار خدمة جوال بي"
    ),
    "Group 3: 'No' AND CashIn >= 20": (
        "يعطيك العافية،معك جاسر من شركة جوال\n"
        "يرجى استخدام محفظتك في اسرع وقت ويمكنك تفعيل خدمة ال USSD كود\n"
        "*110#\n"
        "لضمان استمرار خدمة جوال بي"
    ),
}


if df_preview is not None and mobile_col and numeric_col and text_col:
    if st.button("🚀 Process & Categorize Data"):
        try:
            df = df_preview.copy()

            processed_df = pd.DataFrame({
                "Mobile": df[mobile_col].apply(normalize_mobile),
                "NumericVal": df[numeric_col].apply(parse_numeric),
                "TextVal": df[text_col].apply(normalize_label),
            }).dropna(subset=["Mobile"])

            # Guard against duplicate rows in the source sheet (same customer
            # entered twice) so nobody gets counted, or texted, more than once.
            duplicate_count = int(processed_df["Mobile"].duplicated().sum())
            processed_df = processed_df.drop_duplicates(subset=["Mobile"], keep="first")

            g1 = list(processed_df[processed_df["TextVal"] == "yes"]["Mobile"])
            g2 = list(processed_df[(processed_df["TextVal"] == "no") & (processed_df["NumericVal"] < 20)]["Mobile"])
            g3 = list(processed_df[(processed_df["TextVal"] == "no") & (processed_df["NumericVal"] >= 20)]["Mobile"])
            unmatched = processed_df[processed_df["TextVal"] == "unknown"]

            # Stash results in session_state so they survive reruns caused by
            # other widgets (like the batch-size input) instead of vanishing
            # because st.button() only returns True on the run right after
            # it's clicked.
            st.session_state["results"] = {
                "Group 1: 'Yes' or 'USSD'": g1,
                "Group 2: 'No' AND CashIn < 20": g2,
                "Group 3: 'No' AND CashIn >= 20": g3,
                "unmatched_count": len(unmatched),
                "unmatched_examples": df[text_col].dropna().astype(str).unique()[:10].tolist(),
                "duplicate_count": duplicate_count,
            }
        except Exception as e:
            st.error(f"Processing Error: {e}")

# ---------- 4. Render results (independent of the button, so batch-size
# changes and other widget interactions don't wipe the results) ----------
if "results" in st.session_state:
    results = st.session_state["results"]
    st.success("Processing complete!")

    if results["unmatched_count"] > 0:
        st.warning(f"⚠️ {results['unmatched_count']} rows had a label that wasn't recognized as Yes/USSD/No and were skipped. "
                   f"Examples: {results['unmatched_examples']}")

    if results.get("duplicate_count", 0) > 0:
        st.warning(f"⚠️ {results['duplicate_count']} rows had a mobile number that already appeared earlier in the "
                   f"sheet — only the first occurrence of each number was kept, so no one gets counted or texted twice.")

    batch_size = st.number_input(
        "Numbers per Messages batch (splitting avoids link/recipient limits on some phones)",
        min_value=1, max_value=100, value=20, step=5, key="batch_size",
    )

    for title in ["Group 1: 'Yes' or 'USSD'", "Group 2: 'No' AND CashIn < 20", "Group 3: 'No' AND CashIn >= 20"]:
        numbers = results[title]
        text_result = ','.join(numbers)
        st.markdown(f"""
            <div class="card">
                <h4 style="color: #f8fafc; margin-top: 0;">{title}</h4>
                <p style="color: #94a3b8; font-size: 14px;">Total Items: <b>{len(numbers)}</b></p>
            </div>
        """, unsafe_allow_html=True)
        st.text_area(f"Copy {title}", text_result, height=80)

        message = GROUP_MESSAGES[title]
        st.text_area(f"Message for {title} (fixed)", message, height=100, key=f"msg_{title}", disabled=True)

        if numbers:
            batches = chunk_list(numbers, batch_size)
            st.markdown("<p style='color:#94a3b8; font-size:13px;'>Tap a batch to open Messages with those numbers and the message above pre-filled:</p>", unsafe_allow_html=True)

            # Render strictly in row order (1,2,3 / 4,5,6 ...) instead of
            # round-robin column fill, and force LTR so batch order can't get
            # visually flipped by the surrounding Arabic text.
            cols_per_row = 3
            for row_start in range(0, len(batches), cols_per_row):
                row_batches = batches[row_start:row_start + cols_per_row]
                cols = st.columns(cols_per_row)  # fixed 3 slots -> even grid, last row can be partial
                for offset, batch in enumerate(row_batches):
                    batch_number = row_start + offset + 1
                    link = build_sms_link(batch, message)
                    label = f"📲 Batch {batch_number} ({len(batch)})"
                    with cols[offset]:
                        st.markdown(
                            f'<a href="{link}" target="_blank" dir="ltr" '
                            f'onclick="this.style.background=\'#16a34a\';" '
                            f'style="display:block; text-align:center; background:#1e3a8a; '
                            f'color:white; font-weight:bold; padding:10px 6px; border-radius:8px; '
                            f'text-decoration:none; margin-bottom:8px; transition:background-color 0.2s;">'
                            f'{label}</a>',
                            unsafe_allow_html=True,
                        )
elif df_preview is not None:
    st.info("Select all three columns above to enable processing.")
