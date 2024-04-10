import pandas as pd
import openpyxl
import re


# Read the Excel file, specifying the column range
df = pd.read_excel('output/filtered_data.xlsx', sheet_name=0, usecols="A:A", skiprows=0)

# Print the DataFrame to inspect its structure
print(df)

# Convert the column values into a list if the DataFrame is not empty
if not df.empty:
    myList = df.iloc[:, 0].to_list()  # Use iloc to access the first column
    print(myList)
else:
    print("DataFrame is empty. No data read.")

input_list = myList

# Use regular expression to find the desired part from each string in the list
# Old pattern: r'VITA-(\w+)\*?'
result_parts = [re.search(r'VITA-\d+\b(?=\s*\([xX]?\d*\)|\s*[*]*)', item).group(1) for item in input_list if re.search(r'VITA-(\w+)\*?', item)]

print(result_parts)
result_df = pd.DataFrame(result_parts, columns=['Result Parts'])
result_df.to_excel('output/result_parts.xlsx', index=False)
