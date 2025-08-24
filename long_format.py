import pandas as pd

# Load dataset
df = pd.read_excel("daily_prices_feb_apr2020_clean.xlsx", sheet_name="Sheet1")

# Reshape from wide to long
df_long = df.melt(
    id_vars=["Date", "States/UTs"],   # identifiers
    var_name="Commodity",             # column for commodities
    value_name="Price"                # column for prices
)

# Drop missing values (if some commodity prices are blank for a state/date)
df_long = df_long.dropna(subset=["Price"])

# Quick sanity check
print(df_long.head(20))

# Save to CSV/Excel for Stata use
df_long.to_excel("prices_long.xlsx", index=False)