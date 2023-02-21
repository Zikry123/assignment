#----------------------------------------------------DEPENDENCIES--------------------------------------------------------------------------

import PySimpleGUI as sg
import pandas as pd
import numpy as np
from pathlib import Path
import os
import xlsxwriter
import datetime as dt
import math
#--------------------------------------------------------GLOBAL VARIABLE------------------------------------------------------------------

column_values = [] #Array that stores name of employees
project_values = [] #Array that stores Project code
project_code_error = [] #Names of employees where the project code is wrong
project_error_dict = {} #Dictionary which stores the NRIC of the employees as keys and name as value  where the project code is wrong
df_f = None #Variable that stores the dataframe that contains the main data after it has been cleaned
nric_chosen = '' #The variable that stores the nric that is chosen in the UI
bool_no_project_no = False #To check if the project number field is empty. If a valid project number is added, it becomes true. Else, it is false
bool_no_percentage = False #To check if the percentage field is empty. If a valid percentage is added, it becomes true. Else, it is false
bool_percentage_number = False #Check if the percentage field is a legit number. If it is a number, it is true. Else, it is false
data_dict = {} #Dictionary that stores the array of project number and percentage
dict_nric_project = {} #Dictionary that stores the nric as keys and project number is values
dict_nric_percent = {} #Dictionary that stores the nric as keys and percentage as numbers
nric_added = [] #Array that stores the nric that has successfully been added into the UI
bool_add = True
folder_path = False #Check if folder path is legit
project_percent_arr = [] #Array that stores the data for project and percentage for each nric
#---------------------------------------------------------------FUNCTIONS------------------------------------------------------------------

"""
This functions ensures that the user always saves the work in an 
instance where the is accidental closure of the app
"""
def input_directory(folder_path):
    while not folder_path:
        folder_path = sg.popup_get_folder("Select a folder to save your work")
    return folder_path

"""
Upload the main data in the UI. 
"""
def read_excel_file(file_excel_path,sheet_name):
    df = pd.read_excel(file_excel_path,sheet_name,header=None)
    try:
        start_row = df[df.iloc[:, 0].str.split().str.len() == 1].index[0]
        df = pd.read_excel(file_excel_path,sheet_name, header=start_row)
        df_f = pd.read_excel(file_excel_path,sheet_name, header=5).head(df[df['Department'].isnull()].index[0]) 
        df_f = df_f.fillna('-') 
        column_values_name = df_f['Employee Name'].values.tolist()
        sg.popup_no_titlebar("Upload successful")
    except (KeyError, ValueError, IndexError):
        sg.popup_error("Incorrect format")
        return [], None
    return column_values_name, df_f

"""
Upload the project code validation file into the UI
"""
def read_proj_file(file_excel_path, sheet_name):
    df = pd.read_excel(file_excel_path,sheet_name)
    try:
        project_values = df['Segment Code (Cost Centre)'].values.tolist()
        sg.popup_no_titlebar("Upload successful")
    except (KeyError, ValueError, IndexError):
        sg.popup_error("Incorrect format")
        return []
    return project_values

"""
Uploads the excel file where the work will be saved
"""
def read_data_file(file_excel_path, sheet_name):
    dict_nric_project = {}
    dict_nric_percent = {}
    df = pd.read_excel(file_excel_path,sheet_name)
    try:
        nric_values = df['NRIC No'].unique().tolist()
        for nric in nric_values:
            dict_nric_project[nric] = df[df['NRIC No'] == nric]['Project number'].tolist()
            dict_nric_percent[nric] = df[df['NRIC No'] == nric]['Percentage'].tolist()
        sg.popup_no_titlebar("Upload successful")
    except (KeyError, ValueError, IndexError):
        sg.popup_error("Incorrect format")
        return [], {}, {}
    return nric_values, dict_nric_project, dict_nric_percent
    
"""
Create first row of the app that consist of the X button, percentage and project fields
"""
def create_row(row_counter,project_values):
    row =  [sg.pin(
        sg.Col([[
            sg.Button("X", border_width=0, button_color=(sg.theme_text_color(), sg.theme_background_color()), key=('-DEL-', row_counter)),
            sg.Input(size=(15,1), key=('-PERCENTAGE-', row_counter),enable_events=True),
            sg.Text("Percentage"),
            sg.Combo(values = project_values, readonly=True, size=(20,1),enable_events= True, key=('-PROJECTNO-', row_counter)),
            sg.Text("Project Number")
            ]],
        key=('-ROW-', row_counter)
        ))]
    return row

"""
Create new row of the app that consist of the X button, percentage and project fields when the user clicks on the + sign
"""
def create_row_new(row_counter,value_percent, value_project, project_values):
    row =  [sg.pin(
        sg.Col([[
            sg.Button("X", border_width=0, button_color=(sg.theme_text_color(), sg.theme_background_color()), key=('-DEL-', row_counter)),
            sg.Input(size=(15,1),default_text=value_percent, key=('-PERCENTAGE-', row_counter),enable_events=True),
            sg.Text("Percentage"),
            sg.Combo(values = project_values, readonly=True,default_value= value_project,size=(20,1),enable_events= True, key=('-PROJECTNO-', row_counter)),
            sg.Text("Project Number")
            ]],
        key=('-ROW-', row_counter)
        ))]
    return row

"""
Return the percent array where the user inputs the data for a specific NRIC.
If no data is added to that NRIC, returns an empty array
"""
def get_percent_col(nric):
    if nric in dict_nric_percent:
        return dict_nric_percent[nric]
    else:
        return []

"""
Return the project array where the user inputs the data for a specific NRIC.
If no data is added to that NRIC, returns an empty array
"""
def get_project_col(nric):
    if nric in dict_nric_project:
        return dict_nric_project[nric]
    else:
        return []

"""
Function that calculates the numerical value of each column
based on the percentage. For example if basic wage = 500 and percent = 50% for project A and percent = 50% for project B
the basic wage = 250 for project A and 250 for project B
"""
def calc_result(percent, value):
    if isinstance(percent,(int,float)):
        result = (value * percent)/100
        return result
    else:
        return value

"""
Check if percentage is a float or integer. If it is, append the percentage to the name of the employee
"""
def conditions(s):
    if isinstance(s['percent_col'], (float, int)):
        return s['Employee Name'] + '-' + str(int(s['percent_col'])) + '%'
    else:
        return s['Employee Name']

"""
Function that returns the completed excel file that consist of the split data
"""
def get_excel(excel_file_path, output_folder, df,i,nric_added):
    if len(nric_added) == 0:
        sg.popup_error("Please choose a name and fill in the fields before clicking submit to generate the excel file")
    else:
        filename = Path(excel_file_path).stem
        highlight = df['NRIC No'].isin(nric_added)
        while os.path.exists(Path(output_folder) / f"{filename}_{i}_updated.xlsx"):
            i += 1
        writer = pd.ExcelWriter(Path(output_folder) / f"{filename}_{i}_updated.xlsx",engine='xlsxwriter')
        df.to_excel(writer, index = False,sheet_name = 'Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        yellow_fmt = workbook.add_format({'bg_color': 'yellow'})
        for row, _ in df.iterrows():
            if highlight.iloc[row]:
                worksheet.set_row(row + 1, None, yellow_fmt)
        for column in df:
            column_length = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_length)

        writer.close()
        sg.popup_no_titlebar("Success")

"""
Check if filepath is valid
"""
def is_valid_file(filepath):
    if filepath != '':
        return True
    else:
        sg.popup_error("Filepath not found. Please select a filepath for the input file and output file")
        return False

"""
Function that checks which nric has the incorrect project code
"""
def add_x(df_f, project_values):
    project_error_dict = {}
    if df_f is not None and len(project_values) != 0:
        for i in range(len(df_f['NRIC No'])):
            nric_no = df_f['NRIC No'][i]
            project_code = df_f.loc[df_f['NRIC No'] == nric_no, 'Department'].iloc[0]
            if project_code not in project_values:
                chosen_emp = df_f.loc[df_f['NRIC No'] == nric_no, 'Employee Name'].iloc[0]
                project_error_dict[nric_no] = chosen_emp
    elif df_f is not None and len(project_values) == 0:
        sg.popup("Please upload the project code file")
    elif df_f is None and len(project_values) != 0:
        sg.popup("Please upload the excel payroll file")
    else:
        sg.popup("Please upload both the excel payroll file and project code file")

    return  project_error_dict

"""
Function that updates the table in the UI as the user changes the nric_chosen in the UI
"""
def update_table(nric_added, nric_chosen, dict_nric_percent, dict_nric_project):
    if nric_chosen in nric_added:
        project_percent_arr = []
        for i in range(len(dict_nric_project[nric_chosen])):
            data_row = []
            data_row.append(dict_nric_project[nric_chosen][i])
            data_row.append(int(dict_nric_percent[nric_chosen][i]))
            project_percent_arr.append(data_row)

            
    else:
        project_percent_arr = []
    return project_percent_arr

"""
Function to saves the work of the user when the submit button is clicked
"""
def autosave(dict_nric_project, dict_nric_percent, folder_path):
    data_df_project_number = pd.DataFrame(columns=['NRIC No', 'Project number'])
    data_df_percent = pd.DataFrame(columns=['NRIC No', 'Percentage'])
    for key, value in dict_nric_project.items():
            for val in value:
                data_df_project_number = pd.concat([data_df_project_number, pd.DataFrame.from_records([{'NRIC No': key, 'Project number': val}])])
    for key, value in dict_nric_percent.items():
        for val in value:
            data_df_percent = pd.concat([data_df_percent, pd.DataFrame.from_records([{'NRIC No': key, 'Percentage': val}])])
    data_df_project_number['Percentage'] = data_df_percent['Percentage']
    data_df_project_number = data_df_project_number.reset_index()
    workbook = xlsxwriter.Workbook(folder_path + '/project_code_percentage.xlsx')
    sheet = workbook.add_worksheet()
    sheet.write(0,0,"NRIC No")
    sheet.write(0,1,"Project number")
    sheet.write(0,2,"Percentage")
    for i, row in data_df_project_number.iterrows():
        sheet.write(i + 1, 0, row['NRIC No'])
        sheet.write(i + 1, 1, row['Project number'])
        sheet.write(i + 1, 2, row['Percentage'])
    workbook.close()

folder_path = input_directory(folder_path)

#---------------------------------------------------------APP LAYOUT-----------------------------------------------------------------------

"""
Layout of the UI
"""
layout = [  
            [sg.Text('Search Name'),sg.Input(key='-SEARCH-',enable_events=True, size = (15,1)),sg.Combo(values = column_values, readonly=True, size=(20,1),enable_events= True, key='-COMBO-'), sg.Combo(values = project_code_error , readonly=True, size=(20,1),enable_events= True, key='-VALIDATE-', text_color= 'red'), sg.Text("NRIC"),sg.Input(key='-SHOWNRIC-',disabled=True, size = (25,5))],
            [sg.Text("Input Payroll File"), sg.Input(key = '-IN-',enable_events=True,disabled=True, size = (15,5)), sg.FileBrowse(file_types=(("Excel Files","*xls*"),))],
            [sg.Text("Input Project Code Validation File"), sg.Input(key = '-IN_PROJ-',enable_events=True,disabled=True, size = (15,5)), sg.FileBrowse(file_types=(("Excel Files","*xls*"),))],
            [sg.Text("Input NRIC_Proj_Percent File"), sg.Input(key = '-DATA-',enable_events=True,disabled=True, size = (15,5)), sg.FileBrowse(file_types=(("Excel Files","*xls*"),))],
            [sg.Text("Output File"), sg.Input(key = '-OUT-',enable_events = True,disabled=True), sg.FolderBrowse()],
            [sg.Column([create_row(0,project_values)], k='-ROW_PANEL-'), sg.Table(values = project_percent_arr,
            headings = ["Project","Percent"], max_col_width= 90, auto_size_columns= False, 
            display_row_numbers= False, justification = "right", key = "-Table-", row_height = 35, background_color="dark green", num_rows = 10)],
            [sg.Text("Exit", enable_events=True, key='-EXIT-', tooltip='Exit Application'),
            sg.Text("Clear Input", enable_events=True, key='-REFRESH-', tooltip='Clear input'),
            sg.Text('+', enable_events=True, k='-ADD_ITEM-', tooltip='Add Another Item'),
            sg.Text('Submit',enable_events=True,key = '-SUBMIT-', tooltip='Submit data'),
            sg.Button('Generate Excel File'),
            sg.Button('Validate')]
        ]

window = sg.Window('Payroll App', 
    layout,  use_default_focus=False, font='15', background_color='#5C4033')

row_counter = 0 #Row number for each row where each row consist of project number and percentage
del_arr = [] #The array that stores the row that got deleted when the user clicks the X
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == '-EXIT-':  #Exits the UI
        break
    if event == '-ADD_ITEM-': #Adds a row when the user clicks +
        row_counter += 1
        window.extend_layout(window['-ROW_PANEL-'], [create_row(row_counter, project_values)])
    elif event == '-IN-': #The user inserts the main data file
        column_values,df_f = read_excel_file(file_excel_path=values['-IN-'],sheet_name="Sheet1")
        window['-COMBO-'].update(values = column_values)

    elif event == '-IN_PROJ-':#The user inserts the project code validation file
        project_values = read_proj_file(file_excel_path= values['-IN_PROJ-'], sheet_name = "Sheet1")
        window[('-PROJECTNO-',0)].update(values = project_values)

    elif event == '-DATA-': #Adds the saved work into the UI
        nric_added, dict_nric_project, dict_nric_percent = read_data_file(file_excel_path=values['-DATA-'],sheet_name="Sheet1")

    elif event == '-SEARCH-':#Function that search the name
        search_term = values['-SEARCH-']
        if len(search_term) > 0:
            filtered_choices = [choice for choice in column_values if search_term.upper() in choice[0:len(search_term)] and len(search_term) > 0]
        else:
            filtered_choices = column_values
        window['-COMBO-'].update(values = filtered_choices)

    elif event == '-REFRESH-': #Refresh and remove all the input for only that NRIC when the clear input button is clicked
        for key,val in values.items():
            if ("-PERCENTAGE-"  in key[0]) | ("-PROJECTNO-" in key[0]):
                window[key].update('')

    elif event == 'Generate Excel File': #Function that generates the file that splits the data
        if df_f is None:
            sg.popup("Please upload excel file")
        elif (is_valid_file(values['-IN-']) and (is_valid_file(values['-OUT-']))):
            df_f['project_col'] = df_f['NRIC No'].apply(lambda x: get_project_col(x))
            df_f['percent_col'] = df_f['NRIC No'].apply(lambda x: get_percent_col(x))
            df_f_explode = df_f.explode(['project_col','percent_col'])
            df_f_explode = df_f_explode.fillna('-')
            df_f_explode['Employee Name'] = df_f_explode.apply( conditions, axis=1)
            df_f_explode['Department'] = np.where(df_f_explode['project_col'] != '-', df_f_explode['project_col'], df_f_explode['Department'])
            df_f_explode['Date of Hire'] = pd.to_datetime(df_f_explode['Date of Hire'],errors= 'coerce').dt.strftime("%d/%m/%Y")
            df_f_explode['Date of Birth'] = pd.to_datetime(df_f_explode['Date of Birth'],errors= 'coerce').dt.strftime("%d/%m/%Y")
            for col_names_mod in df_f_explode.iloc[:,27:83].columns:
                df_f_explode[col_names_mod] = df_f_explode.apply(lambda row:calc_result(row['percent_col'], row[col_names_mod]),axis=1)
            df_f_explode = df_f_explode.drop(columns = ['project_col','percent_col'])
            get_excel(values['-IN-'], values['-OUT-'], df_f_explode,0,nric_added)

    elif event == 'Validate':#Function that shows the name of employees that have incorrect project code
        project_error_dict= add_x(df_f, project_values) #Names of employees where the project code is wrong
        project_code_error = list(project_error_dict.values())
        window['-VALIDATE-'].update(values = project_code_error)

    elif event == '-VALIDATE-': #Function that removes that name of the employee from the list of employees with incorrect project code. Executes when the user clicks on the dropdown list consist of employee names with incorrect project code
        for key,val in values.items():
            if ("-PERCENTAGE-"  in key[0]) | ("-PROJECTNO-" in key[0]):
                window[key].update('')
        window['-COMBO-'].update(value = values['-VALIDATE-'])
        window['-SHOWNRIC-'].update(df_f.loc[df_f['Employee Name'] == values['-VALIDATE-'], 'NRIC No'].iloc[0])
        nric_chosen = df_f.loc[df_f['Employee Name'] == values['-VALIDATE-'], 'NRIC No'].iloc[0]
        if nric_chosen in project_error_dict:
            window['-VALIDATE-'].update(value = project_error_dict[nric_chosen])
        else:
            window['-VALIDATE-'].update(value = '')
        
        window['-Table-'].update(values = update_table(nric_added, nric_chosen, dict_nric_percent, dict_nric_project))
                
    elif event == '-COMBO-': #Code executes when the user clicks on the dropdown list of employee names
        name_chosen = values['-COMBO-']
        nric_chosen = df_f.loc[df_f['Employee Name'] == name_chosen, 'NRIC No'].iloc[0]
        if nric_chosen in project_error_dict:
            window['-VALIDATE-'].update(value = project_error_dict[nric_chosen])
        else:
            window['-VALIDATE-'].update(value = '')
        window['-SHOWNRIC-'].update(nric_chosen)
        for key,val in values.items():
            if ("-PERCENTAGE-"  in key[0]) | ("-PROJECTNO-" in key[0]):
                window[key].update('')
        window['-Table-'].update(values = update_table(nric_added, nric_chosen, dict_nric_percent, dict_nric_project))

    elif event[0] == '-DEL-': #Function when user clicks on the X button which removes the row
        del_arr.append(event[1])
        window[('-ROW-', event[1])].update(visible=False)
        window[('-PERCENTAGE-', event[1])].update('')
        window[('-PROJECTNO-', event[1])].update('')
    elif event == '-SUBMIT-': #Function when user clicks the submit button
        bool_percentage_number = False
        bool_no_project_no = False
        bool_no_percentage = False
        if nric_chosen == '':
            sg.popup("Please choose NRIC") #Message appear when NRIC is not chosen when the user clicks the submit button
        else:
            if nric_chosen in project_error_dict:
                del project_error_dict[nric_chosen]
                project_code_error = list(project_error_dict.values())
                window['-VALIDATE-'].update(values = project_code_error)
            if nric_chosen in nric_added:
                bool_add = sg.popup_yes_no("You have already keyed in the data for this NRIC. Do you want to overwrite?") == 'Yes' #Message appear when the user wants to overwrite the data that is saved for that NRIC
            if bool_add:
                percentage_arr = []
                project_arr = []
                for i in range(row_counter+1):
                    try:
                        if i not in del_arr and values[('-PERCENTAGE-',i)] != '':
                            percentage_num = float(values['-PERCENTAGE-',i])
                    except ValueError:
                        sg.popup_error("Percentage column should be a number") #Message appear when the percentage is not a number
                        bool_percentage_number = True
                        break
                    else:
                        if values[('-PERCENTAGE-',i)] == '' and i not in del_arr and values[('-PROJECTNO-',i)] != '':
                            sg.popup_error('Percentage field should not be empty.') #Message appear when percentage column is empty but project column is not empty
                            bool_no_project_no = True
                            break
                        elif values['-PROJECTNO-',i] == '' and i not in del_arr and values[('-PERCENTAGE-',i)] != '':
                            sg.popup_error('Project number field should not be empty') #Message appear when the project number column is empty but percentage column is not empty
                            bool_no_percentage = True
                            break
                        else:
                            percentage_arr.append(values[('-PERCENTAGE-',i)])
                            project_arr.append(values[('-PROJECTNO-',i)])
                data_dict['-PERCENTAGE-'] = percentage_arr
                data_dict['-PROJECTNO-'] = project_arr
                percent_total = 0
                for j in range(row_counter+1):
                    try:
                        percentage_num = float(values['-PERCENTAGE-',j])
                        percent_total += percentage_num
                    except:
                        pass
                if percent_total != 100:
                    sg.popup("Total percentage is not 100%. Please check again") #Message appear the total percentage is not 100% for that NRIC
                else:
                    if ((not bool_no_project_no) and (not bool_no_percentage) and (not bool_percentage_number)):
                        sg.popup("Upload Successful")
                        nric_added.append(nric_chosen)
                        percentage_arr_convert = []
                        project_arr_convert = []
                        for k in range(len(percentage_arr)):
                            try:
                                percentage_convert = float(percentage_arr[k])
                                percentage_arr_convert.append(percentage_convert)
                                project_arr_convert.append(project_arr[k])
                            except ValueError:
                                pass
                            dict_nric_project[nric_chosen] = project_arr_convert
                            dict_nric_percent[nric_chosen] = percentage_arr_convert
                    else:
                        sg.popup_error("There are some errors in the inputs. Please check again") #Message appear when there are errors in the percentage and project column
            else:
                bool_add = True
        autosave(dict_nric_project, dict_nric_percent, folder_path) #Saves the data input when the user clicks on the submit button
        window['-Table-'].update(values = update_table(nric_added, nric_chosen, dict_nric_percent, dict_nric_project))
window.close()
