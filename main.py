import pandas as pd
import re

# Read the Excel file, specifying the column range
df = pd.read_excel('output/filtered_data.xlsx', sheet_name=0, usecols="A:A", skiprows=0)

# Convert the column values into a list if the DataFrame is not empty
if not df.empty:
    input_list = df.iloc[:, 0].tolist()  # Use tolist() to convert to list
    print(input_list)
else:
    print("DataFrame is empty. No data read.")

# Use regular expression to find the desired part from each string in the list
pattern = r'\b\d+\b'  # Match one or more digits surrounded by word boundaries
result_parts = [re.search(pattern, item).group() for item in input_list if re.search(pattern, item)]

print(result_parts)
result_df = pd.DataFrame(result_parts, columns=['Result Parts'])
result_df.to_excel('output/result_parts.xlsx', index=False)
