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

if dataset_navigation == 'Search':
    st.subheader('Use the Sidebar to Search efficiently.')
    #sidebar
    dataset_column = st.sidebar.selectbox('Search by:', ('Theme', 'Area/Subject', 'Title', 'Proponent/s', 'School/ Office'))
    #Lists
    def getdata_Theme():
        identity_data = pd.read_excel(excel_file, sheet_name = sheet_name1, usecols = 'A:I', header = 0)
        return identity_data.set_index("Theme")

    def getdata_AS():
        identity_data = pd.read_excel(excel_file, sheet_name = sheet_name1, usecols = 'A:I', header = 0)
        return identity_data.set_index("Area/Subject")

    def getdata_Title():
        identity_data = pd.read_excel(excel_file, sheet_name = sheet_name1, usecols = 'A:I', header = 0)
        return identity_data.set_index("Title")

    def getdata_Proponent():
        identity_data = pd.read_excel(excel_file, sheet_name = sheet_name1, usecols = 'A:I', header = 0)
        return identity_data.set_index("Proponent/s")

    def getdata_SO():
        identity_data = pd.read_excel(excel_file, sheet_name = sheet_name1, usecols = 'A:I', header = 0)
        return identity_data.set_index("School/ Office")

    #sidebar command
    def fix(dataset_column):
        if dataset_column == 'Theme':
            return getdata_Theme()
        
        elif dataset_column == 'Area/Subject':
            return getdata_AS()

        elif dataset_column == 'Title':
            return getdata_Title()
        
        elif dataset_column == 'Proponent':
            return getdata_Proponent()
        
        elif dataset_column == 'School/ Office':
            return getdata_SO()

    #MAIN
    try:
        df = fix(dataset_column)
        Name = st.multiselect('Choose:', list(df.index))
        if not Name :
           st.error("Please select at least one.")
        else:
            data = df.loc[Name]
            st.write("test here", data.sort_index())

            data = data.T.reset_index()
            data = pd.melt(data, id_vars=["index"]).rename(columns={"Name": "Color"})
    #ERROR
    except URLError as e:
        st.error(
            """
            **This Web app requires access from the Owner.**

            Connection error: %s
        """
            % e.reason
        )
