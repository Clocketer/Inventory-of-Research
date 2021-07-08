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
getdata_Theme = list(df['Theme'])
getdata_AS = list(df['Area/Subject'])
getdata_Title = list(df['Title'])
getdata_Prop = list(df['Proponent/s'])
getdata_SO = list(df['School/ Office'])
getdata_Fin = list(df['Findings'])
getdata_Conclu = list(df['Conclusions'])
getdata_Recomm = list(df['Recommendations '])


dataset_navigation = st.selectbox('Navigation:', ('Search', 'View Research', 'Download Original File', 'Add Entry', 'Edit Entry'))
