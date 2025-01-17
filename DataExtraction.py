import pandas as pd
from docx import Document

excel_file_path = 'E:/DataExtract.xlsx'
df = pd.read_excel(excel_file_path)
data = df.iloc[:340, 1].dropna().astype(str).tolist()
doc = Document()
content = ', '.join(data)
doc.add_paragraph(content)
word_file_path = 'E:/output.docx'
doc.save(word_file_path)
print(f"Data successfully written to {word_file_path}")
