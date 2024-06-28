#ASTC GUI proto
# import shutil
# from kivy.app import App
# from kivy.uix.boxlayout import BoxLayout
# from kivy.uix.textinput import TextInput
# from kivy.uix.button import Button
# from kivy.uix.label import Label

# import openpyxl 
import string
import time
import pandas as pd
from pandas import ExcelWriter
import numpy as np
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl import cell
from openpyxl.utils import get_column_letter
import xlsxwriter
from os import listdir, walk
from os.path import isfile, join
import tkinter as tk
from tkinter import *


# Import Report generatorv1.py
from Report_generatorv1 import create_report, calc_ASTC_val,calc_AIIC_val, pull_testdata, sanitize_filepath 

### # FUNCTIONS FOR PULLING EXCEL DATA ########
import matplotlib.pyplot as plt

# this is a new function combines RT and ASTC datapulls
def RAW_SLM_datapull(self, find_datafile, datatype):
    # pass datatype as '-831_Data.' or '-RT_Data.' to pull the correct data
    # if the self method doesn't pass correctly, may need to re-write
    raw_testpaths = {
        'D': self.slm_data_d_path,
        'E': self.slm_data_e_path,
        # 'A': self.slm_data_a_path
    }
    datafiles = {}
    for key, path in raw_testpaths.items():
        datafiles[key] = [f for f in listdir(path) if isfile(join(path, f))]

    if find_datafile[0] in datafiles:
        datafile_num = datatype + find_datafile[1:] + '.xlsx'
        slm_found = [x for x in datafiles[find_datafile[0]] if datafile_num in x]
        slm_found[0] = raw_testpaths[find_datafile[0]] + slm_found[0]  # If this line errors, the test file is mislabeled or doesn't exist 

    print(slm_found[0])
    if datatype == '-831_Data.':
        srs_data = pd.read_excel(slm_found[0], sheet_name='OBA')
    elif datatype == '-RT_Data.':
        srs_data = pd.read_excel(slm_found[0], sheet_name='Summary')  # data must be in Summary tab for RT meas.
    # srs_data = pd.read_excel(slm_found[0], sheet_name='OBA')  # data must be in OBA tab
    # potentially need a write to excel here...similar to previous function
    # just for 
    return srs_data


#     return srs_data
 # creating a subfunction to return the dataframe from the testplan test list
 # pass it the current test number and the test list
def pull_testplan_data(self,curr_test):
    # AIIC_test = curr_test['AIIC']
    # NIC_test = curr_test['NIC'] 
    # ASTC_test = curr_test['ASTC']

    if int(curr_test['source room vol']) >= 5300 or int(curr_test['receive room vol']) >= 5300:
        NICreporting_Note = 'The receiver and/or source room had a volume exceeding 150 m3 (5,300 cu. ft.), and the absorption of the receiver and/or source room was greater than the maximum allowed per E336-16, Paragraph 9.4.1.2.'
    elif int(curr_test['source room vol']) <= 833 or int(curr_test['receive room vol']) <= 833:
        NICreporting_Note = 'The receiver and/or source room has a volume less than the minimum volume requirement of 25 m3 (883 cu. ft.).'
    else:
        NICreporting_Note = '---'

    room_properties = pd.DataFrame(
        {
            "Source Room Name": curr_test['Source_Room'],
            "Recieve Room Name": curr_test['Receiving_Room'],
            "Testdate": curr_test['Test_Date'],
            "ReportDate": curr_test['Report_Date'],
            "Project Name": curr_test['Project_Name'],
            "Test number": curr_test['Test_Label'],
            "Source Vol" : curr_test['source_room_vol'],
            "Recieve Vol": curr_test['receive_room_vol'],
            "Partition area": curr_test['partition_area'],
            "Partition dim.": curr_test['partition_dim'],
            "Source room Finish" : curr_test['source_room_finish'],
            "Recieve room Finish": curr_test['receive_room_finish'],
            "Srs Floor Descrip.": curr_test['srs_floor'],
            "Srs Ceiling Descrip.": curr_test['srs_ceiling'],
            "Srs Walls Descrip.": curr_test['srs_Walls'],
            "Rec Floor Descrip.": curr_test['rec_floor'],
            "Rec Ceiling Descrip.": curr_test['rec_ceiling'],
            "Rec Walls Descrip.": curr_test['rec_Wall'],          
            "Tested Assembly": curr_test['tested_assembly'],
            "Expected Performance": curr_test['expected_performance'],
            "Annex 2 used?": curr_test['Annex_2_used?'],
            "Test assem. type": curr_test['Test_assembly_Type'],
            "AIIC_test": curr_test['AIIC_test'],
            "NIC_test": curr_test['NIC_test'],
            "ASTC_test": curr_test['ASTC_test'],
            "NIC reporting Note": NICreporting_Note
        },
        index=[0]
        )
    
    
    if curr_test['AIIC_test'] == 1:
        print("AIIC testing enabled, copying data...")
### IIC variables  #### When extending to IIC data
        find_source = curr_test['Source']
        find_rec = curr_test['Recieve '] #trailing whitespace? be sure to verify this is consistent in the excel file
        find_BNL = curr_test['BNL']
        find_RT = curr_test['RT']
        find_posOne = curr_test['Position1']
        find_posTwo = curr_test['Position2']
        find_posThree = curr_test['Position3']
        find_posFour = curr_test['Position4']
        find_poscarpet = curr_test['Carpet']
        find_Tapsrs = curr_test['SourceTap']
        srs_data = RAW_SLM_datapull(self,find_source,'-831_Data.')
        recive_data = RAW_SLM_datapull(self,find_rec,'-831_Data.')
        bkgrnd_data = RAW_SLM_datapull(self,find_BNL,'-831_Data.')
        rt = RAW_SLM_datapull(self,find_RT,'-RT_Data.')
        AIIC_pos1 = RAW_SLM_datapull(self,find_posOne,'-831_Data.')
        AIIC_pos2 = RAW_SLM_datapull(self,find_posTwo,'-831_Data.')
        AIIC_pos3 = RAW_SLM_datapull(self,find_posThree,'-831_Data.')
        AIIC_pos4 = RAW_SLM_datapull(self,find_posFour,'-831_Data.')
        AIIC_carpet = RAW_SLM_datapull(self,find_poscarpet,'-831_Data.')
        AIIC_source = RAW_SLM_datapull(self,find_Tapsrs,'-831_Data.')
        
        single_AIICtest_data = {
            'srs_data': pd.DataFrame(srs_data),
            'recive_data': pd.DataFrame(recive_data),
            'bkgrnd_data': pd.DataFrame(bkgrnd_data),
            'rt': pd.DataFrame(rt),
            'AIIC_pos1': pd.DataFrame(AIIC_pos1),
            'AIIC_pos2': pd.DataFrame(AIIC_pos2),
            'AIIC_pos3': pd.DataFrame(AIIC_pos3),
            'AIIC_pos4': pd.DataFrame(AIIC_pos4),
            'AIIC_source': pd.DataFrame(AIIC_source),
            'AIIC_carpet': pd.DataFrame(AIIC_carpet),
            'room_properties': pd.DataFrame(room_properties)
        }
        return single_AIICtest_data
    
    elif curr_test['ASTC_test'] == 1:
        srs_data = RAW_SLM_datapull(self,curr_test['Source'],'-831_Data.')
        recive_data = RAW_SLM_datapull(self, curr_test['Recieve '],'-831_Data.')
        bkgrnd_data = RAW_SLM_datapull(self,curr_test['BNL'],'-831_Data.')
        rt = RAW_SLM_datapull(self,curr_test['RT'],'-RT_Data.')

        single_ASTCtest_data = {
            'srs_data': pd.DataFrame(srs_data),
            'recive_data': pd.DataFrame(recive_data),
            'bkgrnd_data': pd.DataFrame(bkgrnd_data),
            'rt': pd.DataFrame(rt),
            'room_properties': pd.DataFrame(room_properties)
            }
        return single_ASTCtest_data
        
       

    elif curr_test['NIC_test'] == 1:
        srs_data = RAW_SLM_datapull(self,curr_test['Source'],'-831_Data.')
        recive_data = RAW_SLM_datapull(self, curr_test['Recieve '],'-831_Data.')
        bkgrnd_data = RAW_SLM_datapull(self,curr_test['BNL'],'-831_Data.')
        rt = RAW_SLM_datapull(self,curr_test['RT'],'-RT_Data.')

        single_NICtest_data = {
            'srs_data': pd.DataFrame(srs_data),
            'recive_data': pd.DataFrame(recive_data),
            'bkgrnd_data': pd.DataFrame(bkgrnd_data),
            'rt': pd.DataFrame(rt),
            'room_properties': pd.DataFrame(room_properties)
            }
        return single_NICtest_data
    

    
################ ## # # ###############
    # End of function section #
#######################################

#### THIS LIBRARY IMPORT NEEDS TO REMAIN HERE- Otherwise kivy errors result. 
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.checkbox import CheckBox
from kivy.uix.recycleview import RecycleView
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.gridlayout import GridLayout

class TestListPopup(GridLayout):
    def __init__(self, test_list, **kwargs):
        super().__init__(**kwargs)
        self.cols = len(test_list.columns)  # Set the number of columns based on DataFrame columns

        for column in test_list.columns:
            self.add_widget(Label(text=column))  # Add column headers

        for index, row in test_list.iterrows():
            for column in test_list.columns:
                # Add TextInput for each cell in the DataFrame
                text_input = TextInput(text=str(row[column]), multiline=False)
                self.add_widget(text_input)
# class TestListPopup(GridLayout):
#     def __init__(self, test_list, **kwargs):
#         super(TestListPopup, self).__init__(**kwargs)
#         self.cols = len(test_list.columns)  # Set the number of columns based on DataFrame columns

#         for column in test_list.columns:
#             self.add_widget(Label(text=column))  # Add column headers

#         for index, row in test_list.iterrows():
#             for column in test_list.columns:
#                 # Add TextInput for each cell in the DataFrame
#                 text_input = TextInput(text=str(row[column]), multiline=False)
#                 self.add_widget(text_input)

# GPT code - kivy GUI 
class FileLoaderApp(App):
    def build(self):
        # Initialize path variables as instance variables
        self.test_plan_path = ''
        self.report_template_path = ''
        self.slm_data_d_path = ''
        self.slm_data_e_path = ''
        self.report_output_folder_path = ''

        # Layout
        layout = BoxLayout(orientation='vertical', spacing=10, padding=10)
        # put use description on the top of the window. 
        # ie - need to have excel closed before running
        # format of folders to output to - change this eventually for the program to
        # modifiy the folder structure itself. rev 2 - plot of STC vals, removing excel, ect
        # Label for the first text entry box
        layout.add_widget(Label(text='Test Plan path and filename:'))

        # Text Entry Boxes
        self.first_text_input = TextInput(multiline=False, hint_text='File Path 1')
        self.first_text_input.bind(on_text_validate=self.on_text_validate)
        layout.add_widget(self.first_text_input)
        
        # Label for the second text entry box
        layout.add_widget(Label(text='Excel Report Template Path, will be PDF report export folder:'))

        # Text Entry Box for the second path
        self.second_text_input = TextInput(multiline=False, hint_text='File Path 2')
        self.second_text_input.bind(on_text_validate=self.on_text_validate)
        layout.add_widget(self.second_text_input)

        # Label for the third text entry box
        layout.add_widget(Label(text='SLM Data Meter 1'))

        # Text Entry Box for the third path
        self.third_text_input = TextInput(multiline=False, hint_text='File Path 3')
        self.third_text_input.bind(on_text_validate=self.on_text_validate)
        layout.add_widget(self.third_text_input)

        # Label for the fourth text entry box
        layout.add_widget(Label(text='SLM Data Meter 2'))

        # Text Entry Box for the fourth path
        self.fourth_text_input = TextInput(multiline=False, hint_text='File Path 4')
        self.fourth_text_input.bind(on_text_validate=self.on_text_validate)
        layout.add_widget(self.fourth_text_input)

        # Label for the fifth text entry box
        layout.add_widget(Label(text='Excel Report per test Output Folder path: '))

        # Text Entry Box for the fifth path
        self.fifth_text_input = TextInput(multiline=False, hint_text='File Path 5')
        self.fifth_text_input.bind(on_text_validate=self.on_text_validate)
        layout.add_widget(self.fifth_text_input)

        # CheckBox for AIIC Testing
        # self.aiic_testing_checkbox = CheckBox(active=False, size_hint=(None, None))
        # self.aiic_testing_checkbox.bind(active=self.on_aiic_testing_checkbox_active)
        # layout.add_widget(Label(text='AIIC Testing'))
        # layout.add_widget(self.aiic_testing_checkbox)

        # Load Data Button
        load_button = Button(text='Load Data', on_press=self.load_data)
        layout.add_widget(load_button)

        # Text entry box for calculating single test results
        layout.add_widget(Label(text='Single Test Results:'))
        self.single_test_text_input = TextInput(multiline=False, hint_text='Test Number')
        layout.add_widget(self.single_test_text_input)

        # Calculate Single Test Button
        calculate_single_test_button = Button(text='Calculate Single Test', on_press=self.calculate_single_test)
        layout.add_widget(calculate_single_test_button)

        # Output Reports Button
        output_button = Button(text='Output Reports', on_press=self.output_reports)
        layout.add_widget(output_button)

        # Status Text Box
        self.status_label = Label(text='Status: Ready')
        layout.add_widget(self.status_label)

        return layout

    def on_text_validate(self, instance):
        # Check if the widget is a TextInput before updating instance variables
        if isinstance(instance, TextInput):
            # Update instance variables with the entered text
            if instance is self.first_text_input:
                self.test_plan_path = sanitize_filepath(instance.text)
                instance.text = self.test_plan_path
            elif instance is self.second_text_input:
                self.report_template_path = sanitize_filepath(instance.text)
                instance.text = self.report_template_path
            elif instance is self.third_text_input:
                self.slm_data_d_path = sanitize_filepath(instance.text)
                instance.text = self.slm_data_d_path
            elif instance is self.fourth_text_input:
                self.slm_data_e_path = sanitize_filepath(instance.text)
                instance.text = self.slm_data_e_path
            elif instance is self.fifth_text_input:
                self.report_output_folder_path = sanitize_filepath(instance.text)
                instance.text = self.report_output_folder_path

    def on_aiic_testing_checkbox_active(self, instance, value):
        # Handle checkbox state change
        if value:
            print("AIIC Testing Enabled")
            # You can assign values to curr_test here or add logic based on checkbox state
        else:
            print("AIIC Testing Disabled")
            
    def show_test_list_popup(self, instance):
        # Create a popup with the TestListPopup content
        test_list_popup = TestListPopup(test_list=self.test_list)
        popup = Popup(title='Test List Editor', content=test_list_popup, size_hint=(None, None), size=(400, 400))
        popup.open()

    def load_data(self, instance):
        # # Access the text from all text boxes
        # seems like just a debug step, may not need this once full output report PDF gen. is working since it pulls directly from the SLM datafiles. 
        text_input_fields = [self.first_text_input, self.second_text_input, self.third_text_input, self.fourth_text_input, self.fifth_text_input]
        sanitized_values = [sanitize_filepath(field.text) for field in text_input_fields]
        # first_text_input_value = sanitize_filepath(self.first_text_input.text)
        # second_text_input_value = sanitize_filepath(self.second_text_input.text)
        # third_text_input_value = sanitize_filepath(self.third_text_input.text)
        # fourth_text_input_value = sanitize_filepath(self.fourth_text_input.text)
        # fifth_text_input_value = sanitize_filepath(self.fifth_text_input.text)

        #refactoring to use a dictionary for the input values
        input_values = {
            'test_plan_path': sanitized_values[0],
            'report_template_path': sanitized_values[1],
            'slm_data_d_path': sanitized_values[2],
            'slm_data_e_path': sanitized_values[3],
            'report_output_folder_path': sanitized_values[4]
        }
        
        for attr, value in input_values.items():
            setattr(self, attr, value)
        # self.test_plan_path = first_text_input_value
        # self.report_template_path = second_text_input_value
        # self.slm_data_d_path = third_text_input_value
        # self.slm_data_e_path = fourth_text_input_value
        # self.report_output_folder_path = fifth_text_input_value

            # Update instance variables with the text from the text input boxes

        print('Value from the first text box:', sanitized_values[0])
        print('Value from the second text box:', sanitized_values[1])
        print('Value from the third text box:', sanitized_values[2])
        print('Value from the fourth text box:', sanitized_values[3])
        print('Value from the fifth text box:', sanitized_values[4])

        # Display a message in the status label
        self.status_label.text = 'Status: Loading Data...'
        # Add logic to load data from file paths
        print('Arguments received by load_data:', instance, self.test_plan_path, self.report_template_path, self.slm_data_d_path, self.slm_data_e_path, self.report_output_folder_path)

        # For demonstration purposes, let's just print the file paths
        print('File Paths:', [sanitized_values[0], sanitized_values[1], sanitized_values[2], sanitized_values[3], sanitized_values[4]])
                # Access the text from all text boxes
        testplan_path = self.test_plan_path
        rawReportpath = self.report_template_path
        rawDtestpath = self.slm_data_d_path
        rawEtestpath = self.slm_data_e_path
        outputfolder = self.report_output_folder_path

        self.D_datafiles = [f for f in listdir(rawDtestpath) if isfile(join(rawDtestpath,f))]
        self.E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]
        self.test_list = pd.read_excel(testplan_path)
        testnums = self.test_list['Test Label'] ## Determines the labels and number of excel files copied
        project_name =  self.test_list['Project Name'] # need to access a cell within template the project name to copy to final reports
        project_name = project_name.iloc[1]
        # reports =  [f for f in listdir(rawReportpath) if isfile(join(rawReportpath,f))]
        # Display a message in the status label
        self.status_label.text = 'Status: All test files loaded, ready to generate reports'
        # self.status_label.text = 'Status: Copying Excel template per testplan...'
        # for i in range(len(testnums)):
            # shutil.copy(rawReportpath+reports[0], outputfolder+'23-154 ASTC-NIC Test Report_'+testnums[i]+'_.xlsx')
        # Add logic to generate reports - old excel copy method
        # for index, testnum in enumerate(testnums):
        #     shutil.copy(rawReportpath + reports[0], outputfolder + f'23-154 ASTC-NIC Test Report_{testnum}_.xlsx')    
        
        # self.status_label.text = 'Status: Data Loaded'

    def calculate_single_test(self, instance):
        # NEED TO DEBUG
        # Access the text from the single test text input box
        single_test_text_input_value = self.single_test_text_input.text
        # Display a message in the status label
        self.status_label.text = f'Status: Calculating Single Test {single_test_text_input_value}...'
        # open another window to display the test results
        window = tk.Tk()
        window.title(f'Single Test Results for: {single_test_text_input_value}')
        window.geometry("400x400")
        # Create a status text box
        status_text_box = tk.Text(window, height=10, width=40)
        status_text_box.pack()
        # single test find 
        testplan_path = self.test_plan_path
        testplanfile = pd.read_excel(testplan_path) # 
        rawReportpath = self.report_output_folder_path
        reports =  [f for f in listdir(rawReportpath) if isfile(join(rawReportpath,f))]

        # find_test = '0.1.1'
        find_test = single_test_text_input_value
        mask = testplanfile.applymap(lambda x: find_test in x if isinstance(x,str) else False).to_numpy()
        indices = np.argwhere(mask) 
        # print(indices)
        index = indices[0,0]
        # print(index)
        status_text_box.insert(tk.END,testplanfile.iloc[index])
        foundtest = testplanfile.iloc[index]
        # print the found test in the status box
        status_text_box.insert(tk.END, foundtest) # must come before mainloop
        report_string = '_'+foundtest+'_' 
        status_text_box.insert(tk.END, f"Report string: '{report_string}'")
        status_text_box.insert(tk.END, f"Report list: '{reports}'")
        curr_report_file = [x for x in reports if report_string in x]
        print('Current report file: ',curr_report_file)
        print(curr_report_file[0]) #print the name of the report file being used
        print('Current Test:', foundtest)


        window.mainloop()

        # Add logic to calculate single test results
        print('Single Test:', single_test_text_input_value)
        # Display a message in the status label
        self.status_label.text = f'Status: Single Test {single_test_text_input_value} Calculated'

    def output_reports(self, instance):
        testplan_path = self.test_plan_path
        rawReportpath = self.report_output_folder_path

        test_list = pd.read_excel(testplan_path)
        self.status_label.text = ('Test list:',test_list)
        testnums = test_list['Test Label'] ## Determines the labels and number of excel files copied
        project_name =  test_list['Project Name'] # need to access a cell within template the project name to copy to final reports
        # project_name =  test_list['Project Name'] # need to access a cell within template the project name to copy to final reports
        project_name = project_name.iloc[1]
        
        reports =  [f for f in listdir(rawReportpath) if isfile(join(rawReportpath,f))]

        # Main loop through all test data
        for i in range(len(testnums)):
            # list entry with all test data
            curr_test = test_list.iloc[i]
            print('Current Test:', curr_test)
            aiic_testing_enabled = self.aiic_testing_checkbox.active
            find_report = curr_test['Test Label'] # using the test entry to find the report file name
            report_string = '_'+find_report+'_' 
            print('Report string: ', report_string)
            print('Report list: ', reports)
            curr_report_file = [x for x in reports if report_string in x]
            print('Current report file: ',curr_report_file)
            print(curr_report_file[0]) #print the name of the report file being used
            print('Current Test:', curr_test)
            # single_test_dataframe  = pull_testplan_data(self,curr_test)
            AIIC_test = curr_test['AIIC']
            NIC_test = curr_test['NIC'] 
            ASTC_test = curr_test['ASTC']
           
            # print("writing to report file:")
            print("Test report file name: ", curr_report_file[0])
            # print(raw_report)

            ## need logic here to determine if NIC or ASTC or AIIC , and report one or the other 
            if NIC_test == 1:
                single_NICtest_data = pull_testplan_data(self,curr_test)
                create_report(self,curr_test, single_NICtest_data,test_type='NIC') # still need NIC report format
            elif ASTC_test == 1:
                single_ASTCtest_data = pull_testplan_data(self,curr_test)
                create_report(self,curr_test, single_ASTCtest_data, test_type='ASTC')
            elif AIIC_test == 1:
                single_AIICtest_data = pull_testplan_data(self,curr_test)
                create_report(self,curr_test, single_AIICtest_data, test_type='AIIC')
            else:
                print('No test type selected, skipping test...')
                # some sort of error checking needed here
                continue
            
        print('Generating Reports...')
        self.status_label.text = 'Status: Reports Generated'

## this report display window should appear after the reports are generated, after the output_reports function has run. do i need to pass it more info? TBD.
def display_report_window(report_paths, testplan, test_results):
    # Create a new window
    window = tk.Tk()
    window.title("Report Results")

    # Create a table for the testplan
    tree = ttk.Treeview(window)
    tree["columns"] = list(testplan.columns)
    for i in tree["columns"]:
        tree.heading(i, text=i)
    for index, row in testplan.iterrows():
        tree.insert("", "end", values=list(row))
    tree.pack()

    # Create a canvas for the plots
    fig = Figure(figsize=(5, 5), dpi=100)
    canvas = FigureCanvasTkAgg(fig, master=window)
    canvas.draw()
    canvas.get_tk_widget().pack()

    # Function to update the plot when a row is selected
    def on_select(event):
        selected_item = tree.selection()[0]
        test_id = tree.item(selected_item)["values"][0]
        ax = fig.add_subplot(111)
        ax.clear()
        ax.plot(test_results[test_id])
        canvas.draw()

    tree.bind("<<TreeviewSelect>>", on_select)

    # Create a listbox for the report file paths
    listbox = tk.Listbox(window)
    for path in report_paths:
        listbox.insert("end", path)
    listbox.pack()

    window.mainloop()


if __name__ == '__main__':
    FileLoaderApp().run()

