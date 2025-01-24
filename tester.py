import pandas as pd

# Path to the Excel file
excel_file = "C:/Users/Thais/Documents/Tester/Mappe1.xlsx"

# Read Excel file into a pandas DataFrame
df = pd.read_excel(excel_file)

# If your Excel file doesn't have header row, you can specify:
# df = pd.read_excel(excel_file, header=None)
# Then columns would be accessible by numerical index like df.iloc[:, 0] for the first column.

# Example if columns are named: "Column1", "Column2", "Column3", "Column4", "Column5"
# We want to check if Column4 and Column5 are empty ("" or NaN) 
filtered_df = df[
    (
        (df["Column 4"].isna()) | (df["Column 4"] == "")  # Column 4 is empty
    ) &
    (
        (df["Column 5"].isna()) | (df["Column 5"] == "")  # Column 5 is empty
    )
]

# Now extract the relevant data (columns 1 and 2) for those rows
# This will give you a list of tuples (val_in_col1, val_in_col2)
result_list = list(zip(filtered_df["Column 1"], filtered_df["Column 2"]))

# Print or do something with the result
print("Rows where columns 4 and 5 are empty:")
for item in result_list:
    print(item)
