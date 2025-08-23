import pandas as pd

OUT_FILE = "daily_prices_feb_apr2020.xlsx"
CLEAN_FILE = "daily_prices_feb_apr2020_clean.xlsx"

def clean_and_sort(input_file=OUT_FILE, output_file=CLEAN_FILE):
    """Clean the dataset:
    - Use first row as header
    - Sort by Date
    - Save clean dataset
    """

    # Load without assuming header
    df = pd.read_excel(input_file, header=None)

    # First row contains headers (Date + column names)
    headers = df.iloc[0].tolist()
    df = df[1:]
    df.columns = headers

    # Ensure Date is parsed properly
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
    # Fix date column format
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%d-%m-%Y")

    # Sort
    df = df.sort_values("Date").reset_index(drop=True)
    # Save cleaned version
    df.to_excel(output_file, index=False)
    print(f" Cleaned and sorted file saved as {output_file}")


if __name__ == "__main__":
    clean_and_sort()