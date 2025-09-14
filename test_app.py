import pandas as pd

# Load your Excel file
df = pd.read_excel("assets/Price_List.xlsx")

print(df.head())  # Display the first few rows to verify loading
print(df.columns)  # Display column names to verify structure
print(df['Part Number'])
'''
column_name = "Part Number"

# Convert to string and calculate lengths
df['Length'] = df[column_name].astype(str).map(len)

# Get max length
max_length = df['Length'].max()

# Get the part(s) with the longest length
longest_parts = df[df['Length'] == max_length][column_name].tolist()

print(f"Maximum length of Part Numbers: {max_length}")
print("Part number(s) with maximum length:")
for part in longest_parts:
    print(part)'''
