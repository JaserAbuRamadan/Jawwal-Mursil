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

# 1. File Upload Card
st.markdown('<div class="card">', unsafe_allow_html=True)
st.subheader("📁 1. Source File")
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

selected_sheet = None
if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        selected_sheet = st.selectbox("Target Sheet", sheet_names)
    except Exception as e:
        st.error(f"Could not read sheets: {e}")
st.markdown('</div>', unsafe_allow_html=True)

# 2. Process Button & Results
if uploaded_file and selected_sheet:
    if st.button("🚀 Process & Categorize Data"):
        try:
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            
            if df.shape[1] < 5:
                st.error("Error: The selected sheet must have at least 5 columns.")
            else:
                # Data processing logic
                mobiles = df.iloc[:, 2].dropna().astype(str).str[2:]
                col_numeric = pd.to_numeric(df.iloc[:, 3], errors='coerce').fillna(0)
                col_text = df.iloc[:, 4].astype(str).str.strip().str.lower()

                processed_df = pd.DataFrame({
                    'Mobile': mobiles,
                    'NumericVal': col_numeric,
                    'TextVal': col_text
                }).dropna(subset=['Mobile'])

                g1 = processed_df[processed_df['TextVal'].isin(['yes', 'ussd'])]['Mobile']
                g2 = processed_df[(processed_df['TextVal'] == 'no login before') & (processed_df['NumericVal'] < 20)]['Mobile']
                g3 = processed_df[(processed_df['TextVal'] == 'no login before') & (processed_df['NumericVal'] >= 20)]['Mobile']

                st.success("Processing complete!")

                # Results Cards
                groups = [
                    ("Group 1: 'Yes' or 'USSD'", g1),
                    ("Group 2: 'No Login Before' AND CashIn < 20", g2),
                    ("Group 3: 'No Login Before' AND CashIn >= 20", g3)
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