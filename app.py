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
        
elif dataset_navigation == 'View Research':
    st.subheader('Use the Sidebar to Search efficiently.')
    def index():
        st.sidebar.write('Available: 1-' , len(getdata_ExcNo))
        exce = st.sidebar.text_input('Input Excel')
        if st.sidebar.checkbox("Find/View"):
            return exce
        else:
            st.sidebar.write('Enter No. from the table, then put check')
    num=int(index())
    st.write('Excel No:', num)
    st.write('Theme:',getdata_Theme[num])
    st.write('Area/Subject:', getdata_AS[num])
    st.write('Title:', getdata_Title[num])
    st.write('Proponent/s:', getdata_Prop[num])
    st.write('School/Office:', getdata_SO[num])
    st.write('Findings:')
    st.write(getdata_Fin[num])
    st.write('Conclusions:')
    st.write(getdata_Conclu[num])
    st.write('Recommendations:')
    st.write(getdata_Recomm[num])    


elif dataset_navigation == 'Download Original File':
    def getdata_No():
        identity_data = pd.read_excel(excel_file, sheet_name = sheet_name1, usecols = 'A:I', header = 0)
        return identity_data.set_index("No. ")
    df=getdata_No()
    st.dataframe(df)
    download=st.button('Download Excel File')
    if download:
          'Click here to start downloading!'
          df_download = df
          csv = df_download.to_csv(index=False)
          b64 = base64.b64encode(csv.encode()).decode()  # some strings
          linko= f'<a href="data:file/csv;base64,{b64}" download="Inventory of Research.csv">Download csv file</a>'
          st.markdown(linko, unsafe_allow_html=True)    

elif dataset_navigation == 'Add Entry':

    st.write('This entry will be at Excel No.', len(getdata_ExcNo)+1)
    form = st.form(key='my_form')
    dataset_theme = form.selectbox('Search: On Theme:',
                                            ('Teaching and Learning', 'Human Resources and Development', 'Governance', 'Inclusive Education'))
    if dataset_theme == 'Teaching and Learning':
        theme = 'Teaching and Learning'
    elif dataset_theme == 'Human Resources and Development':
        theme = 'Human Resources and Development'
    elif dataset_theme == 'Inclusive Education':
        theme = 'Inclusive Education'
    elif dataset_theme == 'Governance':
        theme = 'Governance'
    
    ent_as=form.text_input('Enter Area/Subject')
    
    ent_title=form.text_input('Enter Title:')
    
    ent_prop=form.text_input('Enter Proponent/s')
    
    ent_so=form.text_input('Enter School/Office')
    
    ent_fin=form.text_input('Enter Findings:')
    
    ent_conclu=form.text_input('Enter Conclusions:')
    
    ent_recomm=form.text_input('Enter Recommendations:')
    
        
    submit_button = form.form_submit_button(label='Submit')
    if submit_button:
        rb = xlrd.open_workbook(excel_file)
        wb = copy( rb )
        w_sheet = wb.get_sheet(0)
        w_sheet.write( len(getdata_ExcNo)+1,0,len(getdata_ExcNo)+1)
        w_sheet.write( len(getdata_ExcNo)+1,1,theme)
        w_sheet.write( len(getdata_ExcNo)+1,2,ent_as)
        w_sheet.write( len(getdata_ExcNo)+1,3,ent_title)
        w_sheet.write( len(getdata_ExcNo)+1,4,ent_prop)
        w_sheet.write( len(getdata_ExcNo)+1,5,ent_so)
        w_sheet.write( len(getdata_ExcNo)+1,6,ent_fin)
        w_sheet.write( len(getdata_ExcNo)+1,7,ent_conclu)
        w_sheet.write( len(getdata_ExcNo)+1,8,ent_recomm)
        wb.save(excel_file)
        st.write('Table Updated')

elif dataset_navigation == 'Edit Entry':
    st.subheader('Use the Sidebar to Search efficiently.')
    st.write('Check below for the current input after entering Excel No on the sidebar')
    def index():
        st.sidebar.write('Available: 1-' , len(getdata_ExcNo))
        exce = st.sidebar.text_input('Input Excel')
        if st.sidebar.checkbox("Find/View"):
            return exce
        else:
            st.sidebar.write('Enter No. from the table, then put check')
    num=int(index())-1
    st.write('This entry will be at Excel No.', num+1)
    form = st.form(key='my_form')
    dataset_theme = form.selectbox('Search: On Theme:',
                                            ('Teaching and Learning', 'Human Resources and Development', 'Governance', 'Inclusive Education'))
    if dataset_theme == 'Teaching and Learning':
        theme = 'Teaching and Learning'
    elif dataset_theme == 'Human Resources and Development':
        theme = 'Human Resources and Development'
    elif dataset_theme == 'Inclusive Education':
        theme = 'Inclusive Education'
    elif dataset_theme == 'Governance':
        theme = 'Governance'
    st.write('Theme:', getdata_Theme[num])
    ent_as=form.text_input('Enter Area/Subject')
    st.write('Area/ Subject', getdata_AS[num])
    ent_title=form.text_input('Enter Title:')
    st.write('Title:', getdata_Title[num])
    ent_prop=form.text_input('Enter Proponent/s')
    st.write('Proponent/s:', getdata_Prop[num])
    ent_so=form.text_input('Enter School/Office')
    st.write('School/ Office:', getdata_SO[num])
    ent_fin=form.text_input('Enter Findings:')
    st.write('Findings:', getdata_Fin[num])
    ent_conclu=form.text_input('Enter Conclusions:')
    st.write('Conclusion:', getdata_Conclu[num])
    ent_recomm=form.text_input('Enter Recommendations:')
    st.write('Recommendations:', getdata_Recomm[num])
        
    submit_button = form.form_submit_button(label='Submit')
    if submit_button:
        
        rb = xlrd.open_workbook(excel_file)
        wb = copy( rb )
        w_sheet = wb.get_sheet(0)
        w_sheet.write( num+1,0,num)
        w_sheet.write( num+1,1,theme)
        w_sheet.write( num+1,2,ent_as)
        w_sheet.write(num+1,3,ent_title)
        w_sheet.write(num+1,4,ent_prop)
        w_sheet.write( num+1,5,ent_so)
        w_sheet.write( num+1,6,ent_fin)
        w_sheet.write( num+1,7,ent_conclu)
        w_sheet.write(num+1,8,ent_recomm)
        wb.save(excel_file)
        st.write('Table Updated')
st.write('Good Day')
