import pandas as pd

OUT_FILE = "daily_prices_feb_apr2020.xlsx"
CLEAN_FILE = "daily_prices_feb_apr2020_clean.xlsx"

def clean_and_sort(input_file=OUT_FILE, output_file=CLEAN_FILE):
    """Clean the dataset:
    - Use first row as header
    - Drop summary rows (Average, Max, Min, Modal)
    - Sort by Date
    - Save clean dataset
    """

    # Load without assuming header
    df = pd.read_excel(input_file, header=None)

    # First row contains headers (Date + column names)
    headers = df.iloc[0].tolist()
    df = df[1:]
    df.columns = headers

    # Drop summary rows
    summary_rows = ["Average price", "Maximum Price", "Minimum Price", "Modal Price"]
    # Assuming the State/UT column is the second column (adjust if needed)
    state_col = df.columns[1]
    df = df[~df[state_col].isin(summary_rows)]

    # Ensure Date is parsed properly as datetime
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")

    # Sort by actual datetime
    df = df.sort_values("Date").reset_index(drop=True)

    # Optional: format Date as string for readability
    df["Date"] = df["Date"].dt.strftime("%d-%m-%Y")

    # Save cleaned version
    df.to_excel(output_file, index=False)
    print(f"Cleaned and sorted file saved as {output_file}")


if __name__ == "__main__":
    clean_and_sort()
