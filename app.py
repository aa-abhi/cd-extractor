import streamlit as st
import pandas as pd
import zipfile
import io

# ------------------------------------------------
# Page Config
# ------------------------------------------------
st.set_page_config(
    page_title="ACKO Analytics • CD Extractor",
    layout="centered"
)

# ------------------------------------------------
# Header
# ------------------------------------------------
st.title("ACKO Analytics • CD Extractor")
st.caption("Upload a ZIP file containing Excel files (.xlsx)")

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


def process_excel_file(file_bytes, file_name):
    try:
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")

        if "Details" not in excel_file.sheet_names:
            return {"File Name": file_name, "Error": "Missing Details sheet"}

        if "CD Statement" not in excel_file.sheet_names:
            return {"File Name": file_name, "Error": "Missing CD Statement sheet"}

        details_df = pd.read_excel(
            excel_file,
            sheet_name="Details",
            header=None
        )

        cd_df = pd.read_excel(
            excel_file,
            sheet_name="CD Statement",
            header=None
        )

        master_policy_holder = get_last_valid_value(details_df, 3)
        balance = get_last_valid_value(cd_df, 6)

        if isinstance(balance, str):
            balance = balance.replace(",", "").strip()

        try:
            balance = float(balance)
        except (ValueError, TypeError):
            pass

        return {
            "File Name": file_name,
            "Master Policy Holder Name": master_policy_holder,
            "Balance": balance
        }

    except Exception as e:
        return {
            "File Name": file_name,
            "Master Policy Holder Name": None,
            "Balance": None,
            "Error": str(e)
        }


# ------------------------------------------------
# Upload ZIP
# ------------------------------------------------
uploaded_zip = st.file_uploader(
    "Upload ZIP File",
    type=["zip"]
)

if "final_df" not in st.session_state:
    st.session_state.final_df = None


if uploaded_zip:

    zip_size_mb = uploaded_zip.size / (1024 * 1024)
    st.info(f"ZIP Size: {zip_size_mb:.2f} MB")

    if st.button("Process ZIP"):

        results = []

        with st.spinner("Extracting ZIP..."):
            zip_bytes = uploaded_zip.read()
            zip_file = zipfile.ZipFile(io.BytesIO(zip_bytes))

            excel_files = [
                name for name in zip_file.namelist()
                if name.endswith(".xlsx")
                and not name.startswith("__MACOSX/")
                and not name.split("/")[-1].startswith("._")
            ]

        if not excel_files:
            st.error("No Excel files (.xlsx) found inside ZIP.")
            st.stop()

        total_files = len(excel_files)
        st.success(f"{total_files} Excel files found inside ZIP.")

        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, file_name in enumerate(excel_files):

            with zip_file.open(file_name) as file:
                file_bytes = file.read()

            result = process_excel_file(file_bytes, file_name)
            results.append(result)

            progress = (idx + 1) / total_files
            progress_bar.progress(progress)
            status_text.text(f"Processing file {idx + 1} of {total_files}")

        st.session_state.final_df = pd.DataFrame(results)

        progress_bar.empty()
        status_text.empty()

        st.success("Processing completed successfully ✅")


# ------------------------------------------------
# Display Results
# ------------------------------------------------
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