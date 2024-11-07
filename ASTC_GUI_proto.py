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
from os import listdir, walk
from os.path import isfile, join
import tkinter as tk
from tkinter import *

import matplotlib.pyplot as plt
# Import Report generatorv1.py
# from Report_generatorv1 import create_report,calc_AIIC_val, RAW_SLM_datpull, sanitize_filepath,

from config import *
from data_processor import *
# this is a new function combines RT and ASTC datapulls

 # creating a subfunction to return the dataframe from the testplan test list
 # pass it the current test number and the test list
 
### # FUNCTIONS FOR PULLING EXCEL DATA ########
## moving this to data_processor.py

# def pull_testplan_data(self,curr_test):
#     #pulling data from the test plan
#     if int(curr_test['source room vol']) >= 5300 or int(curr_test['receive room vol']) >= 5300:
#         NICreporting_Note = 'The receiver and/or source room had a volume exceeding 150 m3 (5,300 cu. ft.), and the absorption of the receiver and/or source room was greater than the maximum allowed per E336-16, Paragraph 9.4.1.2.'
#     elif int(curr_test['source room vol']) <= 833 or int(curr_test['receive room vol']) <= 833:
#         NICreporting_Note = 'The receiver and/or source room has a volume less than the minimum volume requirement of 25 m3 (883 cu. ft.).'
#     else:
#         NICreporting_Note = '---'

#     room_properties = pd.DataFrame(
#         {
#             "Site Name": curr_test['Site_Name'],
#             "Client Name": curr_test['Client_Name'],
#             "Source Room Name": curr_test['Source_Room'],
#             "Recieve Room Name": curr_test['Receiving_Room'],
#             "Testdate": curr_test['Test_Date'],
#             "ReportDate": curr_test['Report_Date'],
#             "Project Name": curr_test['Project_Name'],
#             "Test number": curr_test['Test_Label'],
#             "Source Vol" : curr_test['source_room_vol'],
#             "Recieve Vol": curr_test['receive_room_vol'],
#             "Partition area": curr_test['partition_area'],
#             "Partition dim.": curr_test['partition_dim'],
#             "Source room Finish" : curr_test['source_room_finish'],
#             "Recieve room Finish": curr_test['receive_room_finish'],
#             "Srs Floor Descrip.": curr_test['srs_floor'],
#             "Srs Ceiling Descrip.": curr_test['srs_ceiling'],
#             "Srs Walls Descrip.": curr_test['srs_Walls'],
#             "Rec Floor Descrip.": curr_test['rec_floor'],
#             "Rec Ceiling Descrip.": curr_test['rec_ceiling'],
#             "Rec Walls Descrip.": curr_test['rec_Wall'],          
#             "Tested Assembly": curr_test['tested_assembly'],
#             "Expected Performance": curr_test['expected_performance'],
#             "Annex 2 used?": curr_test['Annex_2_used?'],
#             "Test assem. type": curr_test['Test_assembly_Type'],
#             "AIIC_test": curr_test['AIIC_test'],
#             "NIC_test": curr_test['NIC_test'],
#             "ASTC_test": curr_test['ASTC_test'],
#             "DTC_test": curr_test['DTC_test'],
#             "NIC reporting Note": NICreporting_Note
#         },
#         index=[0]
#         )
    
#     ### 
#     if curr_test['AIIC_test'] == 1:
#         print("AIIC testing enabled, copying data...")
# ### IIC variables  #### When extending to IIC data
#         find_source = curr_test['Source']
#         find_rec = curr_test['Recieve '] #trailing whitespace? be sure to verify this is consistent in the excel file
#         find_BNL = curr_test['BNL']
#         find_RT = curr_test['RT']
#         find_posOne = curr_test['Position1']
#         find_posTwo = curr_test['Position2']
#         find_posThree = curr_test['Position3']
#         find_posFour = curr_test['Position4']
#         find_poscarpet = curr_test['Carpet']
#         find_Tapsrs = curr_test['SourceTap']
#         srs_data = RAW_SLM_datapull(self,find_source,'-831_Data.')
#         recive_data = RAW_SLM_datapull(self,find_rec,'-831_Data.')
#         bkgrnd_data = RAW_SLM_datapull(self,find_BNL,'-831_Data.')
#         rt = RAW_SLM_datapull(self,find_RT,'-RT_Data.')
#         AIIC_pos1 = RAW_SLM_datapull(self,find_posOne,'-831_Data.')
#         AIIC_pos2 = RAW_SLM_datapull(self,find_posTwo,'-831_Data.')
#         AIIC_pos3 = RAW_SLM_datapull(self,find_posThree,'-831_Data.')
#         AIIC_pos4 = RAW_SLM_datapull(self,find_posFour,'-831_Data.')
#         AIIC_carpet = RAW_SLM_datapull(self,find_poscarpet,'-831_Data.')
#         AIIC_source = RAW_SLM_datapull(self,find_Tapsrs,'-831_Data.')
        
#         single_AIICtest_data = {
#             'srs_data': pd.DataFrame(srs_data),
#             'recive_data': pd.DataFrame(recive_data),
#             'bkgrnd_data': pd.DataFrame(bkgrnd_data),
#             'rt': pd.DataFrame(rt),
#             'AIIC_pos1': pd.DataFrame(AIIC_pos1),
#             'AIIC_pos2': pd.DataFrame(AIIC_pos2),
#             'AIIC_pos3': pd.DataFrame(AIIC_pos3),
#             'AIIC_pos4': pd.DataFrame(AIIC_pos4),
#             'AIIC_source': pd.DataFrame(AIIC_source),
#             'AIIC_carpet': pd.DataFrame(AIIC_carpet),
#             'room_properties': pd.DataFrame(room_properties)
#         }
#     else:
#         single_AIICtest_data = None
#         # return single_AIICtest_data
    
#     if curr_test['ASTC_test'] == 1:
#         srs_data = RAW_SLM_datapull(self,curr_test['Source'],'-831_Data.')
#         recive_data = RAW_SLM_datapull(self, curr_test['Recieve '],'-831_Data.')
#         bkgrnd_data = RAW_SLM_datapull(self,curr_test['BNL'],'-831_Data.')
#         rt = RAW_SLM_datapull(self,curr_test['RT'],'-RT_Data.')

#         single_ASTCtest_data = {
#             'srs_data': pd.DataFrame(srs_data),
#             'recive_data': pd.DataFrame(recive_data),
#             'bkgrnd_data': pd.DataFrame(bkgrnd_data),
#             'rt': pd.DataFrame(rt),
#             'room_properties': pd.DataFrame(room_properties)
#             }
#     else:
#         single_ASTCtest_data = None
#         # return single_ASTCtest_data
        
#     if curr_test['NIC_test'] == 1:
#         srs_data = RAW_SLM_datapull(self,curr_test['Source'],'-831_Data.')
#         recive_data = RAW_SLM_datapull(self, curr_test['Recieve '],'-831_Data.')
#         bkgrnd_data = RAW_SLM_datapull(self,curr_test['BNL'],'-831_Data.')
#         rt = RAW_SLM_datapull(self,curr_test['RT'],'-RT_Data.')

#         single_NICtest_data = {
#             'srs_data': pd.DataFrame(srs_data),
#             'recive_data': pd.DataFrame(recive_data),
#             'bkgrnd_data': pd.DataFrame(bkgrnd_data),
#             'rt': pd.DataFrame(rt),
#             'room_properties': pd.DataFrame(room_properties)
#             }
#         # return single_NICtest_data
#     else:
#         single_NICtest_data = None
    
#     if curr_test['DTC_test'] == 1:

#         srs_data = RAW_SLM_datapull(self,curr_test['Source'],'-831_Data.')
#         recive_data = RAW_SLM_datapull(self, curr_test['Recieve '],'-831_Data.')
#         bkgrnd_data = RAW_SLM_datapull(self,curr_test['BNL'],'-831_Data.')
#         rt = RAW_SLM_datapull(self,curr_test['RT'],'-RT_Data.')
#         srs_door_open = RAW_SLM_datapull(self,curr_test['Source_Door_Open'],'-831_Data.')
#         srs_door_closed = RAW_SLM_datapull(self,curr_test['Source_Door_Closed'],'-831_Data.')
#         recive_door_open = RAW_SLM_datapull(self, curr_test['Recieve_Door_Open '],'-831_Data.')
#         recive_door_closed = RAW_SLM_datapull(self, curr_test['Recieve_Door_Closed '],'-831_Data.')


#         single_DTCtest_data = {
#             'srs_data': pd.DataFrame(srs_data),
#             'recive_data': pd.DataFrame(recive_data),
#             'bkgrnd_data': pd.DataFrame(bkgrnd_data),
#             'rt': pd.DataFrame(rt),
#             'srs_door_open': pd.DataFrame(srs_door_open),
#             'srs_door_closed': pd.DataFrame(srs_door_closed),
#             'recive_door_open': pd.DataFrame(recive_door_open),
#             'recive_door_closed': pd.DataFrame(recive_door_closed),
#             'room_properties': pd.DataFrame(room_properties)
#             }
#     else:
#         single_DTCtest_data = None
#         # return single_DTCtest_data
#     return room_properties, single_NICtest_data, single_DTCtest_data, single_ASTCtest_data, single_AIICtest_data
    

    
################ ## # # ###############
    # GUI Interface - kivy #
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
        self.test_plan_path = TextInput(multiline=False, hint_text='test_plan_path')
        self.test_plan_path.bind(on_text_validate=self.on_text_validate)
        layout.add_widget(self.test_plan_path)

        layout.add_widget(Label(text='SLM Data Meter 1'))

        # Text Entry Box for SLM data 1
        self.slm_data_d_path = TextInput(multiline=False, hint_text='SLM Data Meter 1')
        self.slm_data_d_path.bind(on_text_validate=self.on_text_validate)
        layout.add_widget(self.slm_data_d_path)

        # Label for the fourth text entry box
        layout.add_widget(Label(text='SLM Data Meter 2'))

        # Text Entry Box for the fourth path
        self.slm_data_e_path = TextInput(multiline=False, hint_text='SLM Data Meter 2')
        self.slm_data_e_path.bind(on_text_validate=self.on_text_validate)
        layout.add_widget(self.slm_data_e_path)

        # Label for the fifth text entry box
        layout.add_widget(Label(text='Excel Report per test Output Folder path: '))

        # Text Entry Box for the fifth path
        self.output_folder_path = TextInput(multiline=False, hint_text='Output Path')
        self.output_folder_path.bind(on_text_validate=self.on_text_validate)
        layout.add_widget(self.output_folder_path)

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
        output_button = Button(text='Output All Reports', on_press=self.output_all_reports)
        layout.add_widget(output_button)

        # Status Text Box
        self.status_label = Label(text='Status: Ready')
        layout.add_widget(self.status_label)

        # Debug check box for testing
        self.debug_check_box = CheckBox(active=True)
        self.debug_label = Label(text='Debug')
        layout.add_widget(self.debug_check_box)
        layout.add_widget(self.debug_label)

        return layout

    def on_text_validate(self, instance):
        # Check if the widget is a TextInput before updating instance variables
        if isinstance(instance, TextInput):
            # Update instance variables with the entered text
            if instance is self.test_plan_path:
                self.test_plan_path = sanitize_filepath(instance.text)
                instance.text = self.test_plan_path
            elif instance is self.output_folder_path:
                self.output_folder_path = sanitize_filepath(instance.text)
                instance.text = self.report_template_path
            elif instance is self.third_text_input:
                self.slm_data_d_path = sanitize_filepath(instance.text)
                instance.text = self.slm_data_d_path
            elif instance is self.fourth_text_input:
                self.slm_data_e_path = sanitize_filepath(instance.text)
                instance.text = self.slm_data_e_path
            elif instance is self.single_test_text_input:
                self.single_test_text_input = sanitize_filepath(instance.text)
                instance.text = self.single_test_text_input
            
    def show_test_list_popup(self, instance):
        # Create a popup with the TestListPopup content
        test_list_popup = TestListPopup(test_list=self.test_list)
        popup = Popup(title='Test List Editor', content=test_list_popup, size_hint=(None, None), size=(400, 400))
        popup.open()

    def load_data(self, instance):
        # # Access the text from all text boxes
        # seems like just a debug step, may not need this once full output report PDF gen. is working since it pulls directly from the SLM datafiles. 
        text_input_fields = [self.test_plan_path, self.output_folder_path, self.slm_data_d_path, self.slm_data_e_path, self.fifth_text_input]
        sanitized_values = [sanitize_filepath(field.text) for field in text_input_fields]

        #refactoring to use a dictionary for the input values
        input_values = {
            'test_plan_path': sanitized_values[0],
            'report_template_path': sanitized_values[1],
            'slm_data_d_path': sanitized_values[2],
            'slm_data_e_path': sanitized_values[3],
            'report_output_folder_path': sanitized_values[4],
            'single_test_text_input': sanitized_values[5]
        }
        
        for attr, value in input_values.items():
            setattr(self, attr, value)
            # Update instance variables with the text from the text input boxes

        print('Value from the first text box:', sanitized_values[0])
        print('Value from the second text box:', sanitized_values[1])
        print('Value from the third text box:', sanitized_values[2])
        print('Value from the fourth text box:', sanitized_values[3])
        print('Value from the fifth text box:', sanitized_values[4])
        print('Value from the single test text box:', sanitized_values[5])

        # Display a message in the status label
        self.status_label.text = 'Status: Loading Data...'
        # Add logic to load data from file paths
        print('Arguments received by load_data:', instance, self.test_plan_path, self.report_template_path, self.slm_data_d_path, self.slm_data_e_path, self.report_output_folder_path)

        # For demonstration purposes, let's just print the file paths
        print('File Paths:', [sanitized_values[0], sanitized_values[1], sanitized_values[2], sanitized_values[3], sanitized_values[4], sanitized_values[5]])
                # Access the text from all text boxes
        testplan_path = self.test_plan_path
        # rawReportpath = self.report_template_path
        rawDtestpath = self.slm_data_d_path
        rawEtestpath = self.slm_data_e_path
        outputfolder = self.report_output_folder_path

        self.D_datafiles = [f for f in listdir(rawDtestpath) if isfile(join(rawDtestpath,f))]
        self.E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]
        # self.test_plan = pd.read_excel(testplan_path)
        # self.room_properties, self.test_types = load_test_plan(testplan_path)
        testnums = self.test_list['Test Label'] ## Determines the labels and number of excel files copied
        project_name =  self.test_list['Project Name'] # need to access a cell within template the project name to copy to final reports
        project_name = project_name.iloc[1]
        ### debug print outs here for the datafiles ### 
        print(self.D_datafiles)
        print(self.E_datafiles)
        # Display a message in the status label
        self.status_label.text = 'Status: All test files loaded, ready to generate reports'
        # open a popup window to display the test list
        self.show_test_list_popup(self.test_list)
        #

    @classmethod
    def output_all_reports(self, instance):
        # access the text from the load_data function
        print('Arguments received by output_reports:', instance, self.test_plan_path, self.report_template_path, self.slm_data_d_path, self.slm_data_e_path, self.report_output_folder_path)

        testplan_path = self.test_plan_path
        reportOutputfolder = self.report_output_folder_path

        test_list = pd.read_excel(testplan_path)
        self.status_label.text = ('Test list:', test_list)
        testnums = test_list['Test Label']
        # pull the debug checkbox from the GUI and set it equal to debug
        debug = 1 if self.debug_check_box.active else 0

        if debug == 1:  
            print('Debug mode enabled')
            print('Test list:', test_list)
            print('Test numbers:', testnums)
            print('Test plan path:', testplan_path)
            print('Report output folder:', reportOutputfolder)
        
        # this will iterate through all the tests in the test plan and generate the reports
        for i in range(len(testnums)):
            curr_test = test_list.iloc[i]

            if debug == 1:
                print('Current Test:', curr_test)
            self.status_label.text = 'Status: Reporting test: ' + curr_test['Test Label']
            
            # Determine the test type and generate the report
            # loop through the test types and generate the report for the first one that is enabled
            self.room_properties, self.test_types, self.test_data = load_test_plan(testplan_path)

            selected_test_type = next((test_type for test_type in TestType if curr_test[test_type.value] == 1), None)

            # for all our test types, if the test type is enabled, generate the report
            for test_type in TestType:
                if curr_test[test_type.value] == 1:
                    self.status_label.text = f'Status: Generating report for {test_type.value} test...'
                    # write a loop entry here to make a debug function where i can see/edit the data
                    if debug == 1:
                        ## open a popup window to display the room properties, test types, and test data
                        self.show_test_properties_popup(self.room_properties, self.test_types, self.test_data)
                        # pause the program until the user closes the popup
                        self.popup.wait_window()
                        ## create a debug popup window to edit the room properties, test types, and test data
                        self.edit_test_properties_popup(self.room_properties, self.test_types, self.test_data)
                    room_properties, test_types, test_data = load_test_plan(self, curr_test, test_type.value)
                    # test data for report debugging goes here before printing to PDF
                    ####### function below for calculating the results from the loaded test plan 
                    if test_type == TestType.DTC:
                        DTC_report_data =self.calc_DTC_data(self, curr_test, test_data) # does it need to have self?
                        # now pass the report data to the create_report function
                        create_report(self, curr_test, DTC_report_data, reportOutputfolder, test_type=selected_test_type.value)
                    elif test_type == TestType.NIC:
                        NIC_report_data = self.calc_NIC_data(self, curr_test, test_data)
                        create_report(self, curr_test, NIC_report_data, reportOutputfolder, test_type=selected_test_type.value)
                    elif test_type == TestType.AIIC:
                        AIIC_report_data = self.calc_AIIC_data(self, curr_test, test_data)
                        create_report(self, curr_test, AIIC_report_data, reportOutputfolder, test_type=selected_test_type.value)
                    elif test_type == TestType.ASTC:
                        ASTC_report_data = self.calc_ASTC_data(self, curr_test, test_data)
                        create_report(self, curr_test, ASTC_report_data, reportOutputfolder, test_type=selected_test_type.value)
                    ## maybe write a log file generator here
                    # write a function here to print a report number and potentially interuupt or restart the test load. 
                    # print(test_data)

                    create_report(self, curr_test, test_data, reportOutputfolder, test_type=selected_test_type.value)
                    # print the test number label and test type]
                    print(f'Test number: {curr_test["Test Label"]}, Test type: {selected_test_type.value}')
            else:
                print('No test type selected, skipping test...')
                
        

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

        # rawReportpath = self.report_output_folder_path
        # reports =  [f for f in listdir(rawReportpath) if isfile(join(rawReportpath,f))]

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

        print('Current Test:', foundtest)
        
        # output reports function takes in the entire instance, we need to pass it the foundtest 
        FileLoaderApp.output_reports(self, instance)

        window.mainloop()

        # Add logic to calculate single test results
        print('Single Test:', single_test_text_input_value)
        # Display a message in the status label
        self.status_label.text = f'Status: Calculating Single Test {single_test_text_input_value}...'
        room_properties, test_types, test_data = load_test_plan(testplan_path)
        create_report(self, foundtest, test_data, reportOutputfolder, test_type=selected_test_type.value)
        self.status_label.text = f'Status: Single Test {single_test_text_input_value} Calculated'

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

