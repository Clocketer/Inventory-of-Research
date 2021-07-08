import streamlit as st
import pandas as pd
from urllib.error import URLError
import base64
import io
import xlrd
from xlutils.copy import copy
#Files
excel_file = 'Inventory of Research.xls'
sheet_name1 = 'Sheet1'
st.title('Inventory of Researches')
df = pd.read_excel(excel_file)
rb = xlrd.open_workbook(excel_file)
workbook = xlrd.open_workbook(excel_file)
wb = copy( rb )

getdata_ExcNo = list(df['No. '])

st.write(getdata_ExcNo)
