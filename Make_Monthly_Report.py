import pandas as pd
from smartsheet_dataframe import get_report_as_df
import streamlit as st
import tempfile
import os

st.set_page_config(layout="wide")
st.markdown("<h1 style='text-align: center'>Monthly Report</h1><br><br>", unsafe_allow_html=True)

dirpath = tempfile.mkdtemp()
excel_filename = "Monthly_Report.xlsx"
excel_file_path = os.path.join(dirpath,excel_filename)

def grab_monthly_report():
    global excel_file_path
    ACCESS_TOKEN = 'r84vLFkD7bYJaE3eXaSB3Z8mWwD4wgTm3zucM'
    MonthlyID = '4734175795275652'

    MonthlyDF = get_report_as_df(token=ACCESS_TOKEN,report_id=MonthlyID)
    # Path to the Excel file
    #excel_file_path = 'D:/Users/Santi/Downloads/Monthly_Report_Test_12.xlsx'

    # Name of the sheet containing the data
    sheet_name = 'Monthly Report'

    # Name of the column to group by
    group_by_column = 'Client Team'

    # Read the Excel file
    #excel_file = pd.read_excel(excel_file_path, sheet_name=sheet_name)

    excel_file = MonthlyDF

    excel_file.drop(["Primary"], axis=1, inplace=True)
    excel_file.drop(["parent_id"], axis=1, inplace=True)
    excel_file.drop(["row_id"], axis=1, inplace=True)
    excel_file["Date"] = pd.to_datetime(excel_file['Date'])
    excel_file['Date'] = excel_file['Date'].dt.strftime('%m/%d/%Y')

    aired_list = ["AIRED","1","2","3","4","5","6","7","8"]

    def highlight_aired(row):
        value = row.loc['Game Result']
        if value in aired_list:
            color = '#BAFFC9' # Red
        else:
            color = 'None' # Blue
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

#excel_file.to_excel(excel_file_path)

#excel_file.style.set_properties(**{'color':'black'})

grab_monthly_report()

st.download_button(label="Download Monthly Report",data=excel_file_path,file_name="Monthly_Report.xlsx",mime='text/csv')