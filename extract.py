import pandas as pd
import re

# Read the Excel file, specifying the column range
df = pd.read_excel('Active+Listings+Report+04-12-2024.xlsx', usecols="D:D", skiprows=0)
print(df)

# Define a regular expression pattern to match the desired strings
pattern = r'\bVITA-\s*\d+\s*(?:\(\s*[xX]?\d*\s*\)|\s*[*]*)\b'
# Filter rows containing strings matching the pattern
filtered_df = df[df.iloc[:, 0].str.match(pattern, case=False, na=False)]

# Convert the DataFrame to a string
df_string = filtered_df.to_string(index=False, header=False)
# print(filtered_df)

# Write the string to a text file
with open('output/filtered_data.txt', 'w') as f:
    f.write(df_string)

filtered_df.to_csv('output/filtered_data.txt', index=False, header=False, sep='\t')

# Write the filtered DataFrame to an Excel file
filtered_df.to_excel('output/filtered_data.xlsx', index=False)

# Convert the column values into a list if the DataFrame is not empty
if not filtered_df.empty:
    input_list = filtered_df.iloc[:, 0].tolist()  # Use tolist() to convert to list
    print(input_list)
else:
    print("DataFrame is empty. No data read.")

# Use regular expression to find the desired part from each string in the list
pattern = r'\b\d+\b'  # Match one or more digits surrounded by word boundaries
result_parts = [re.search(pattern, item).group() for item in input_list if re.search(pattern, item)]

# print(result_parts)
result_df = pd.DataFrame(result_parts, columns=['Result Parts'])
result_df.to_excel('output/result_parts.xlsx', index=False)


# Read the Excel file into a DataFrame

# Extract column values as a list
column_values = result_df['Result Parts'].tolist()

def remove_duplicates(arr):
    unique_nums = set()
    result = []

    for num in arr:
        if num not in unique_nums:
            result.append(num)
            unique_nums.add(num)

    return result

# Remove duplicates from the column values
column_values_without_duplicates = remove_duplicates(column_values)

# Print array without duplicates
print("Array without duplicates:", column_values_without_duplicates)

# Create a DataFrame from the unique values
removed_dupes_df = pd.DataFrame(column_values_without_duplicates, columns=['Removed_Parts'])
removed_dupes_df = removed_dupes_df.astype(str)

# Write the DataFrame to an Excel file
removed_dupes_df.to_excel('output/removed_dupes.xlsx', index=False)

# Old pattern 1: r'^(AMZ,|\b(?:VITA-|vita-|amz,))'
# Old pattern 2: r'\b(?i:AMZ,|\s*VITA\s*-|vita\s*-|amz,)\s*'
# New pattern: r'VITA-\d+\b(?=\s*\([xX]?\d*\)|\s*[*]*)'