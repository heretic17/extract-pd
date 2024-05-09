import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel('output/result_parts.xlsx', usecols=["Result Parts"], dtype=str)

# Extract column values as a list
column_values = df['Result Parts'].tolist()

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
result_df = pd.DataFrame(column_values_without_duplicates, columns=['Removed_Parts'])
result_df = result_df.astype(str)

# Write the DataFrame to an Excel file
result_df.to_excel('output/removed_dupes.xlsx', index=False)