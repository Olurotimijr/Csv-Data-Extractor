import pandas as pd
from docx import Document
import numpy as np

# Read Excel file
excel_file = 'DECENDING AND CADRE 1546.xlsx'  #name of file to be extracted from here
df = pd.read_excel(excel_file)

#print column names for debugging
print(df.columns)

# column from which data needs to be extracted
specific_data = df[['NAMES'.strip()]]

# Create a Word document
doc = Document()

# Add a table to Word and insert data
table = doc.add_table(rows=1, cols=len(specific_data.columns))
for index, row in specific_data.iterrows():
    cells = table.add_row().cells
    for col_num, value in enumerate(row):
        cells[col_num].text = str(value)

# Save the Word document
word_file = 'jordan.docx'
doc.save(word_file)

print(f"Data extracted from Excel and saved to {word_file}")


