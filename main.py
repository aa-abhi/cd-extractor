import os
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
import multiprocessing

INPUT_FOLDER = "input_files"
OUTPUT_FILE = "final_output.csv"


def get_last_valid_value(df, column_index):
    """
    Returns last meaningful (non-empty) value from a column.
    Handles blank strings and whitespace.
    """
    if column_index >= len(df.columns):
        return None

    column = df.iloc[:, column_index]

    column = column.dropna()
    column = column[column.astype(str).str.strip() != ""]

    if column.empty:
        return None

    return column.iloc[-1]


def process_single_file(file_name):
    file_path = os.path.join(INPUT_FOLDER, file_name)

    try:
        with pd.ExcelFile(file_path, engine="openpyxl") as xls:

            if "Details" not in xls.sheet_names:
                return {"File Name": file_name, "Error": "Missing Details sheet"}

            if "CD Statement" not in xls.sheet_names:
                return {"File Name": file_name, "Error": "Missing CD Statement sheet"}

            details_df = pd.read_excel(
                xls,
                sheet_name="Details",
                header=None
            )

            cd_df = pd.read_excel(
                xls,
                sheet_name="CD Statement",
                header=None
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


def process_files():
    files = [
        f for f in os.listdir(INPUT_FOLDER)
        if f.endswith(".xlsx")
    ]

    if not files:
        print("No Excel files found.")
        return pd.DataFrame()

    max_workers = min(8, multiprocessing.cpu_count())

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        results = list(executor.map(process_single_file, files))

    return pd.DataFrame(results)


if __name__ == "__main__":
    final_df = process_files()

    if not final_df.empty:
        print("\nFinal Output:\n")
        print(final_df)

        final_df.to_csv(OUTPUT_FILE, index=False)
        print(f"\nOutput saved to {OUTPUT_FILE}")
    else:
        print("No data processed.")
