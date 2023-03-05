# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a file to convert word to pdf files
"""

# first read in all the pdf that need to be converted ----- 

import pandas as pd

df = pd.read_csv('FIG_pdf.csv', encoding='latin1')

print(df)

# conver them to docx files ---
import os
from pdf2docx import parse
import pandas as pd

# define a function to convert pdf to docx
def pdf_to_docx(pdf_path, docx_path):
    parse(pdf_path, docx_path)

# loop through each file and convert pdf to docx
for index, row in df.iterrows():
    full_dir = row['full_dir']
    if os.path.isfile(full_dir) and full_dir.endswith('.pdf'):
        docx_path = full_dir[:-4] + '.docx'
        pdf_to_docx(full_dir, docx_path)

