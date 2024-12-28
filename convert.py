import pandas as pd
import pyexcel as p

# Correct the file paths
xls_file = '/Users/aguschandra/Downloads/soh gc 20241226.xls'  # Full path to the .xls file
xlsx_file = '/Users/aguschandra/Downloads/soh_gc_20241226.xlsx'  # Full path for the converted .xlsx file

# Convert .xls to .xlsx
p.save_book_as(file_name=xls_file, dest_file_name=xlsx_file)
print("Conversion complete!")

# Read the .xlsx file
df = pd.read_excel(xlsx_file, engine="openpyxl")

# Display the first few rows
print(df.head())