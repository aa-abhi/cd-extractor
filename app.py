import streamlit as st
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

# ------------------------------------------------
# Page Config
# ------------------------------------------------
st.set_page_config(
    page_title="ACKO Analytics • CD Extractor",
    layout="centered"
)

# ------------------------------------------------
# Custom Styling
# ------------------------------------------------
st.markdown("""
    <style>
        .main-title {
            font-size: 34px;
            font-weight: 700;
        }
        .subtitle {
            font-size: 16px;
            color: #6c757d;
            margin-bottom: 25px;
        }
        .footer {
            margin-top: 60px;
            padding-top: 20px;
            border-top: 1px solid #2c2c2c;
            font-size: 14px;
            color: #9aa0a6;
            text-align: center;
        }
    </style>
""", unsafe_allow_html=True)

# ------------------------------------------------
# Header
# ------------------------------------------------
st.markdown('<div class="main-title">ACKO Analytics • CD Extractor</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload Excel files to extract Master Policy Holder Name and CD Balance · Supports up to 1,000 files</div>', unsafe_allow_html=True)

# ------------------------------------------------
# Core Logic
# ------------------------------------------------
def get_last_valid_value(df, column_index):
    if column_index >= len(df.columns):
        return None

    column = df.iloc[:, column_index]
    column = column.dropna()
    column = column[column.astype(str).str.strip() != ""]

    if column.empty:
        return None

    return column.iloc[-1]


def process_single_file(uploaded_file):
    try:
        details_df = pd.read_excel(
            uploaded_file,
            sheet_name="Details",
            header=None,
            engine="openpyxl"
        )

        cd_df = pd.read_excel(
            uploaded_file,
            sheet_name="CD Statement",
            header=None,
            engine="openpyxl"
        )

        master_policy_holder = get_last_valid_value(details_df, 3)
        balance = get_last_valid_value(cd_df, 6)

        # Clean balance safely
        if isinstance(balance, str):
            balance = balance.replace(",", "").strip()

        try:
            balance = float(balance)
        except:
            pass

        return {
            "File Name": uploaded_file.name,
            "Master Policy Holder Name": master_policy_holder,
            "Balance": balance
        }

    except Exception as e:
        return {
            "File Name": uploaded_file.name,
            "Master Policy Holder Name": None,
            "Balance": None,
            "Error": str(e)
        }


# ------------------------------------------------
# Upload Section
# ------------------------------------------------
uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx"],
    accept_multiple_files=True
)

if "final_df" not in st.session_state:
    st.session_state.final_df = None

if uploaded_files:

    if st.button("Process Files"):

        total_files = len(uploaded_files)
        progress_bar = st.progress(0)
        status_text = st.empty()

        results = []

        with st.spinner("Processing files..."):

            for idx, uploaded_file in enumerate(uploaded_files):

                result = process_single_file(uploaded_file)
                results.append(result)

                progress = (idx + 1) / total_files
                progress_bar.progress(progress)

                status_text.text(f"Processing file {idx + 1} of {total_files}")

        st.session_state.final_df = pd.DataFrame(results)

        progress_bar.empty()
        status_text.empty()

        st.success("Processing completed successfully ✅")
# ------------------------------------------
# Display Results If Already Processed
# ------------------------------------------
if st.session_state.final_df is not None:

    st.subheader("Preview")
    st.dataframe(st.session_state.final_df, use_container_width=True)

    csv_data = st.session_state.final_df.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="Download Result CSV",
        data=csv_data,
        file_name="cd_extraction_output.csv",
        mime="text/csv"
    )

else:
    st.warning("Upload Excel files and click 'Process Files' to begin.")
# ------------------------------------------------
# Footer
# ------------------------------------------------
st.markdown("""
    <style>
        .footer {
            margin-top: 60px;
            padding-top: 25px;
            border-top: 1px solid #2c2c2c;
            text-align: center;
            font-size: 14px;
            color: #9aa0a6;
        }
        .footer p {
            margin: 6px 0;
        }
        .footer a {
            color: #4da6ff;
            text-decoration: none;
        }
        .footer a:hover {
            text-decoration: underline;
        }
    </style>

    <div class="footer">
        <p><strong>CSV Merger v1.0</strong> · Internal use only</p>
        <p>Need help or found a bug? Contact:
            <a href="mailto:abhishek.singh@acko.tech">
                abhishek.singh@acko.tech
            </a>
        </p>
    </div>
""", unsafe_allow_html=True)
