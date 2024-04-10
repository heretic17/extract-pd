import pandas as pd

# Read the Excel file, specifying the column range
df = pd.read_excel('Active+Listings+Report+04-10-2024.xlsx', usecols="D:D", skiprows=0)
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


# Old pattern 1: r'^(AMZ,|\b(?:VITA-|vita-|amz,))'
# Old pattern 2: r'\b(?i:AMZ,|\s*VITA\s*-|vita\s*-|amz,)\s*'
# New pattern: r'VITA-\d+\b(?=\s*\([xX]?\d*\)|\s*[*]*)'