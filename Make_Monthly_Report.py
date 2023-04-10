import pandas as pd
from smartsheet_dataframe import get_report_as_df
import streamlit as st
from tkinter import Tk
from tkinter.filedialog import asksaveasfilename
import streamlit.components.v1 as components
import os

st.set_page_config(layout="wide")


def grab_monthly_report():
    global excel_file_path, writer
    ACCESS_TOKEN = 'r84vLFkD7bYJaE3eXaSB3Z8mWwD4wgTm3zucM'
    MonthlyID = '4734175795275652'

    MonthlyDF = get_report_as_df(token=ACCESS_TOKEN,report_id=MonthlyID)
    # Path to the Excel file

    # Name of the sheet containing the data
    sheet_name = 'Monthly Report'

    # Name of the column to group by
    group_by_column = 'Client Team'

    excel_file = MonthlyDF

    excel_file.drop(["Primary"], axis=1, inplace=True)
    excel_file.drop(["parent_id"], axis=1, inplace=True)
    excel_file.drop(["row_id"], axis=1, inplace=True)
    excel_file["Date"] = pd.to_datetime(excel_file['Date'])
    excel_file['Date'] = excel_file['Date'].dt.strftime('%m/%d/%Y')

    good_list = ["AIRED","6","7","8"]
    ok_list = ["4","5"]
    bad_list = ["DIDN'T AIR","1","2","3"]

    def highlight_aired(row):
        value = row.loc['Game Result']
        op_value = row.loc['How Would you Rate the Quality of the Game?']
        if value in good_list or op_value in good_list and not value:
            color = '#BAFFC9' #Green
        elif value in ok_list or op_value in ok_list and not value:
            color = '#fdffba'
        elif value in bad_list or op_value in bad_list and not value:
            color = '#ffbaba'
        else:
            color = 'None'
        return ['background-color: {}'.format(color) for r in row]


    # Get the unique values in the specified column
    unique_values = excel_file[group_by_column].unique()
    unique_values = sorted(unique_values.astype(str))
    unique_values = [i for i in unique_values if i]

    #Create a Pandas Excel writer using the original Excel file
    writer = pd.ExcelWriter(excel_file_path, engine='openpyxl')

    # Loop through the unique values and create a new sheet for each one
    for value in unique_values:
        try:
            # Filter the data for the current value
            filtered_data = excel_file[excel_file[group_by_column] == value]
            
            # Write the filtered data to a new sheet
            sheet_name = f'{value}'

            filtered_data.style.apply(highlight_aired, axis=1).to_excel(writer, sheet_name=sheet_name, index=False)
        except:
            pass
    # Save the changes to the Excel file
    writer.close()
    
#Tkinter is desktop GUI which can't be used on server
# root = Tk()
# root.withdraw()
# root.attributes("-topmost", True)

def select_folder():
    global dl_file,dl_loc,excel_file_path
    #Tkinter is desktop GUI which can't be used on server
#     global excel_file_path
#     folder_selected = asksaveasfilename(initialfile = 'Monthly_Report.xlsx',
# defaultextension=".xlsx",filetypes=[("Excel File","*.xlsx")],parent=root)
#     excel_file_path = folder_selected
    excel_file_path = os.path.join(dl_loc,dl_file)

def make_report():
    select_folder()
    if excel_file_path:
        grab_monthly_report()
        with col3:
            st.markdown('Downloaded')

# embed streamlit docs in a streamlit app
components.iframe("https://app.smartsheet.com/b/publish?EQBCT=04ddab69560e4685857be8e772dfc018",height=700,scrolling=True)

col1, col2, col3 , col4, col5 = st.columns(5)
with col1:
    pass
with col2:
    pass
with col5:
    pass
with col4:
    pass
with col3:
    dl_loc = st.text_input('Input Download Location','')
    dl_file = "Monthly_Report.xlsx"
    if st.button('Download Monthly Report'):
        make_report()
    else:
        pass
