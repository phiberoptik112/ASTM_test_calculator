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

### # FUNCTIONS FOR PULLING EXCEL DATA ########
import matplotlib.pyplot as plt

def plot_curves(STCCurve, ASTC_curve):
    plt.figure(figsize=(10, 6))
    plt.plot(STCCurve, label='STC Curve')
    plt.plot(ASTC_curve, label='ASTC Curve')
    plt.xlabel('Index')
    plt.ylabel('Value')
    plt.grid(True)
    plt.tick_params(axis='x', rotation=45)
    plt.title('STC Curve vs ASTC Curve')
    plt.legend()
    plt.show()

# this is a new function combines RT and ASTC datapulls
def pull_testdata(self, find_datafile, datatype):
    # pass datatype as '-831_Data.' or '-RT_Data.' to pull the correct data
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

# Pulling SLM data function definition OLD 
# def write_testdata(self,find_datafile, reportfile, newsheetname):
#     rawDtestpath = self.slm_data_d_path
#     rawEtestpath = self.slm_data_e_path
#     rawReportpath = self.report_output_folder_path
#     D_datafiles = [f for f in listdir(rawDtestpath) if isfile(join(rawDtestpath,f))]
#     E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]
#     # excel = win32com.client.Dispatch("Excel.Application")
#     if find_datafile[0] =='D':
#         datafile_num = find_datafile[1:]
#         datafile_num = '-831_Data.'+datafile_num+'.xlsx'
#         slm_found = [x for x in D_datafiles if datafile_num in x]
#         slm_found[0] = rawDtestpath+slm_found[0]# If this line errors, the test file is mislabled or doesn't exist 
#         # print(srs_slm_found)
#     elif find_datafile[0] == 'E':
#         datafile_num = find_datafile[1:]
#         datafile_num = '-831_Data.'+datafile_num+'.xlsx'
#         slm_found = [x for x in E_datafiles if datafile_num in x]
#         slm_found[0] = rawEtestpath+slm_found[0]# If this line errors, the test file is mislabled or doesn't exist 

#     print(slm_found[0])

#     srs_data = pd.read_excel(slm_found[0],sheet_name='OBA') # data must be in OBA tab
#     with ExcelWriter(
#     rawReportpath+reportfile,
#     mode="a",
#     engine="openpyxl",
#     if_sheet_exists="replace",
#     ) as writer:
#         srs_data.to_excel(writer, sheet_name=newsheetname) #writes to report file
#     time.sleep(1)

#     return srs_data
 
## OLD function for specific paths For RT pull
# def write_RTtestdata(self, find_datafile, reportfile,newsheetname):
#     rawDtestpath = self.slm_data_d_path
#     rawEtestpath = self.slm_data_e_path
#     rawReportpath = self.report_output_folder_path
#     D_datafiles = [f for f in listdir(rawDtestpath) if isfile(join(rawDtestpath,f))]
#     E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]

#     if find_datafile[0] =='D':
#         datafile_num = find_datafile[1:]
#         datafile_num = '-RT_Data.'+datafile_num+'.xlsx'
#         slm_found = [x for x in D_datafiles if datafile_num in x]
#         slm_found[0] = rawDtestpath+slm_found[0]# If this line errors, the test file is mislabled or doesn't exist 
#         # print(srs_slm_found)
#     elif find_datafile[0] == 'E':
#         datafile_num = find_datafile[1:]
#         datafile_num = '-RT_Data.'+datafile_num+'.xlsx'
#         slm_found = [x for x in E_datafiles if datafile_num in x]
#         slm_found[0] = rawEtestpath+slm_found[0] # If this line errors, the test file is mislabled or doesn't exist 

#     print(slm_found[0])

#     srs_data = pd.read_excel(slm_found[0],sheet_name='Summary')
#     # data must be in Summary tab for RT meas.
#     # could reduce this function by also passing the sheet to be read into the args.
#     with ExcelWriter(
#     rawReportpath+reportfile,
#     mode="a",
#     engine="openpyxl",
#     if_sheet_exists="replace",
#     ) as writer:
#         srs_data.to_excel(writer, sheet_name=newsheetname) 
#     time.sleep(1)

#     return srs_data
 # creating a subfunction to return the dataframe from the testplan test list
 # pass it the current test number and the test list
def pull_testplan_data(self,curr_test):

    AIIC_test = curr_test['AIIC']
    NIC_test = curr_test['NIC'] 
    ASTC_test = curr_test['ASTC']

    room_properties = pd.DataFrame(
    {
        "Source Room Name": curr_test['Source Room'],
        "Recieve Room Name": curr_test['Receiving Room'],
        "Testdate": curr_test['Test Date'],
        "ReportDate": curr_test['Report Date'],
        "Project Name": project_name,
        "Test number": find_report,
        "Source Vol" : curr_test['source room vol'],
        "Recieve Vol": curr_test['receive room vol'],
        "Partition area": curr_test['partition area'],
        "Partition dim.": curr_test['partition dim'],
        "Source room Finish" : curr_test['source room finish'],
        "Recieve room Finish": curr_test['receive room finish'],
        "Srs Floor Descrip.": curr_test['srs_floor'],
        "Srs Ceiling Descrip.": curr_test['srs_ceiling'],
        "Srs Walls Descrip.": curr_test['srs_Walls'],
        "Rec Floor Descrip.": curr_test['rec_floor'],
        "Rec Ceiling Descrip.": curr_test['rec_ceiling'],
        "Rec Walls Descrip.": curr_test['rec_Wall'],          
        "Tested Assembly": curr_test['tested assembly'],
        "Expected Performance": curr_test['expected performance'],
        "Annex 2 used?": curr_test['Annex 2 used?'],
        "Test assem. type": curr_test['Test assembly Type'],
        "NIC reporting Note": NICreporting_Note
    },
    index=[0]
    )
    if self.aiic_testing_checkbox.active:
        print("AIIC testing enabled, copying data...")
### IIC variables  #### When extending to IIC data
        find_posOne = curr_test['Position1']
        find_posTwo = curr_test['Position2']
        find_posThree = curr_test['Position3']
        find_posFour = curr_test['Position4']
        find_poscarpet = curr_test['Carpet']
        find_Tapsrs = curr_test['SourceTap']
        srs_data = pull_testdata(self,find_source,'-831_Data.')
        recive_data = pull_testdata(self,find_rec,'-831_Data.')
        bkgrnd_data = pull_testdata(self,find_BNL,'-831_Data.')
        rt = pull_testdata(self,find_RT,'-RT_Data.')
        AIIC_pos1 = pull_testdata(self,find_posOne,'-831_Data.')
        AIIC_pos2 = pull_testdata(self,find_posTwo,'-831_Data.')
        AIIC_pos3 = pull_testdata(self,find_posThree,'-831_Data.')
        AIIC_pos4 = pull_testdata(self,find_posFour,'-831_Data.')
        AIIC_carpet = pull_testdata(self,find_poscarpet,'-831_Data.')
        AIIC_source = pull_testdata(self,find_Tapsrs,'-831_Data.')
        
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
    
    elif NIC_test == 1 or ASTC_test == 1:
        print("AIIC data not enabled")
        print("ASTC or NIC test enabled, copying data...")
        srs_data = pull_testdata(self,curr_test['Source'],'-831_Data.')
        recive_data = pull_testdata(self, curr_test['Recieve '],'-831_Data.')
        bkgrnd_data = pull_testdata(self,curr_test['BNL'],'-831_Data.')
        rt = pull_testdata(self,curr_test['RT'],'-RT_Data.')

        single_ASTCtest_data = {
            'srs_data': pd.DataFrame(srs_data),
            'recive_data': pd.DataFrame(recive_data),
            'bkgrnd_data': pd.DataFrame(bkgrnd_data),
            'rt': pd.DataFrame(rt),
            'room_properties': pd.DataFrame(room_properties)
            }
        single_NICtest_data = {
            'srs_data': pd.DataFrame(srs_data),
            'recive_data': pd.DataFrame(recive_data),
            'bkgrnd_data': pd.DataFrame(bkgrnd_data),
            'rt': pd.DataFrame(rt),
            'room_properties': pd.DataFrame(room_properties)
            }
        return single_ASTCtest_data, single_NICtest_data
    

    
def sanitize_filepath(filepath):
    ##"""Sanitize a file path by replacing forward slashes with backslashes."""
    filepath = filepath.replace('T:', '//DLA-04/Shared/')
    filepath = filepath.replace('\\','/')
    # need to add a line to append a / at the end of the filename
    return filepath
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
        super(TestListPopup, self).__init__(**kwargs)
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
        self.aiic_testing_checkbox = CheckBox(active=False, size_hint=(None, None))
        self.aiic_testing_checkbox.bind(active=self.on_aiic_testing_checkbox_active)
        layout.add_widget(Label(text='AIIC Testing'))
        layout.add_widget(self.aiic_testing_checkbox)

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
        # find_test = '0.1.1'
        find_test = single_test_text_input_value
        mask = testplanfile.applymap(lambda x: find_test in x if isinstance(x,str) else False).to_numpy()
        indices = np.argwhere(mask) 
        # print(indices)
        index = indices[0,0]
        # print(index)
        print(testplanfile.iloc[index])
        foundtest = testplanfile.iloc[index]
        # print the found test in the status box
        status_text_box.insert(tk.END, foundtest) # must come before mainloop


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
                        #### write room dimensions ####
            
            #moving into function for getting room test data and properties
            srs_roomName = curr_test['Source Room']
            rec_roomName = curr_test['Receiving Room']
            testdate = curr_test['Test Date'] 
            reportdate = curr_test['Report Date']
            source_vol = curr_test['source room vol']
            rec_vol = curr_test['receive room vol']
            partition_area = curr_test['partition area']
            partition_dim = curr_test['partition dim']
            source_rm_finish = curr_test['source room finish']
            rec_rm_finish = curr_test['receive room finish']
            srs_floor_descrip = curr_test['srs_floor']
            srs_ceiling_descrip = curr_test['srs_ceiling']
            srs_walls_descrip = curr_test['srs_Walls']
            rec_floor_descrip = curr_test['rec_floor']
            rec_ceiling_descrip = curr_test['rec_ceiling']
            rec_walls_descrip = curr_test['rec_Wall']
            tested_assem = curr_test['tested assembly']
            expected_perf = curr_test['expected performance']
            annex_two = curr_test['Annex 2 used?']
            test_assem_type = curr_test['Test assembly Type']

            # need to add these into the testplan excel sheet
            AIIC_test = curr_test['AIIC']
            NIC_test = curr_test['NIC'] 
            ASTC_test = curr_test['ASTC']

            # this code will not work, each test may either AIIC and ASTC, NIC and ASTC , or NIC
            #need to build in logic for creat report for which test report to generate, and do it separately.

            # if AIIC_test == 1:
            #     test_type = 'AIIC'
            # elif NIC_test == 1:
            #     test_type = 'NIC'
            # elif ASTC_test == 1:
            #     test_type = 'ASTC'

            if int(source_vol) >= 5300 or int(rec_vol) >= 5300:
                NICreporting_Note = 'The receiver and/or source room had a volume exceeding 150 m3 (5,300 cu. ft.), and the absorption of the receiver and/or source room was greater than the maximum allowed per E336-16, Paragraph 9.4.1.2.'
            elif int(source_vol) <= 833 or int(rec_vol) <= 833:
                NICreporting_Note = 'The receiver and/or source room has a volume less than the minimum volume requirement of 25 m3 (883 cu. ft.).'
            else:
                NICreporting_Note = '---'

            room_properties = pd.DataFrame(
                {
                    "Source Room Name": srs_roomName,
                    "Recieve Room Name": rec_roomName,
                    "Testdate": testdate,
                    "ReportDate": reportdate,
                    "Project Name": project_name,
                    "Test number": find_report,
                    "Source Vol" : source_vol,
                    "Recieve Vol": rec_vol,
                    "Partition area": partition_area,
                    "Partition dim.": partition_dim,
                    "Source room Finish" : source_rm_finish,
                    "Recieve room Finish": rec_rm_finish,
                    "Srs Floor Descrip.": srs_floor_descrip,
                    "Srs Ceiling Descrip.": srs_ceiling_descrip,
                    "Srs Walls Descrip.": srs_walls_descrip,
                    "Rec Floor Descrip.": rec_floor_descrip,
                    "Rec Ceiling Descrip.": rec_ceiling_descrip,
                    "Rec Walls Descrip.": rec_walls_descrip,          
                    "Tested Assembly": tested_assem,
                    "Expected Performance": expected_perf,
                    "Annex 2 used?": annex_two,
                    "Test assem. type": test_assem_type,
                    "NIC reporting Note": NICreporting_Note
                },
                index=[0]
            )

            if aiic_testing_enabled:
                print("AIIC testing enabled, copying data...")
        ### IIC variables  #### When extending to IIC data
                find_posOne = curr_test['Position1']
                find_posTwo = curr_test['Position2']
                find_posThree = curr_test['Position3']
                find_posFour = curr_test['Position4']
                find_poscarpet = curr_test['Carpet']
                find_Tapsrs = curr_test['SourceTap']
                # '-831_Data.'
                # '-RT_Data.'
                ### older datapull from excel and write to each report excel doc
                # srs_data = write_testdata(self,find_source,curr_report_file[0],'ASTC Source')
                # recive_data = write_testdata(self,find_rec,curr_report_file[0],'ASTC Receive')
                # bkgrnd_data = write_testdata(self,find_BNL,curr_report_file[0],'BNL')
                # rt = write_RTtestdata(self,find_RT,curr_report_file[0],'RT')
                # AIIC_pos1 = write_testdata(self,find_posOne,curr_report_file[0],'AIIC POS 1')
                # AIIC_pos2 = write_testdata(self,find_posTwo,curr_report_file[0],'AIIC POS 2')
                # AIIC_pos3 = write_testdata(self,find_posThree,curr_report_file[0],'AIIC POS 3')
                # AIIC_pos4 = write_testdata(self,find_posFour,curr_report_file[0],'AIIC POS 4')
                # AIIC_carpet = write_testdata(self,find_poscarpet,curr_report_file[0],'AIIC CARPET')
                # AIIC_source = write_testdata(self,find_Tapsrs,curr_report_file[0],'AIIC Source')
                
                srs_data = pull_testdata(self,find_source,'-831_Data.')
                recive_data = pull_testdata(self,find_rec,'-831_Data.')
                bkgrnd_data = pull_testdata(self,find_BNL,'-831_Data.')
                rt = pull_testdata(self,find_RT,'-RT_Data.')
                AIIC_pos1 = pull_testdata(self,find_posOne,'-831_Data.')
                AIIC_pos2 = pull_testdata(self,find_posTwo,'-831_Data.')
                AIIC_pos3 = pull_testdata(self,find_posThree,'-831_Data.')
                AIIC_pos4 = pull_testdata(self,find_posFour,'-831_Data.')
                AIIC_carpet = pull_testdata(self,find_poscarpet,'-831_Data.')
                AIIC_source = pull_testdata(self,find_Tapsrs,'-831_Data.')
                
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
            else:
                print("AIIC data not enabled")

            ## ASTC variables - test file numbers 
            find_source = curr_test['Source']
            find_rec = curr_test['Recieve ']
            find_BNL = curr_test['BNL']
            find_RT = curr_test['RT']
            
            # old method to pull data from excel and write to report excel doc
            # reving to pull_testdata function
            # srs_data = write_testdata(self,find_source,curr_report_file[0],'ASTC Source')
            # recive_data = write_testdata(self,find_rec,curr_report_file[0],'ASTC Receive')
            # bkgrnd_data = write_testdata(self,find_BNL,curr_report_file[0],'BNL')
            # rt = write_RTtestdata(self,find_RT,curr_report_file[0],'RT')

            srs_data = pull_testdata(self,find_source,'-831_Data.')
            recive_data = pull_testdata(self,find_rec,'-831_Data.')
            bkgrnd_data = pull_testdata(self,find_BNL,'-831_Data.')
            rt = pull_testdata(self,find_RT,'-RT_Data.')
     
            # ## UNTESTED, needs validating
            single_ASTCtest_data = {
            'srs_data': pd.DataFrame(srs_data),
            'recive_data': pd.DataFrame(recive_data),
            'bkgrnd_data': pd.DataFrame(bkgrnd_data),
            'rt': pd.DataFrame(rt),
            'room_properties': pd.DataFrame(room_properties)
            }
            single_NICtest_data = {
            'srs_data': pd.DataFrame(srs_data),
            'recive_data': pd.DataFrame(recive_data),
            'bkgrnd_data': pd.DataFrame(bkgrnd_data),
            'rt': pd.DataFrame(rt),
            'room_properties': pd.DataFrame(room_properties)
            }
            raw_report = rawReportpath+curr_report_file[0]
    
            
            if int(source_vol) >= 5300 or int(rec_vol) >= 5300:
                NICreporting_Note = 'The receiver and/or source room had a volume exceeding 150 m3 (5,300 cu. ft.), and the absorption of the receiver and/or source room was greater than the maximum allowed per E336-16, Paragraph 9.4.1.2.'
            elif int(source_vol) <= 833 or int(rec_vol) <= 833:
                NICreporting_Note = 'The receiver and/or source room has a volume less than the minimum volume requirement of 25 m3 (883 cu. ft.).'
            else:
                NICreporting_Note = '---'


            # append all this data or write individully to a dataframe, then save as new sheet
            # refer to this sheet with the SLM Data sheet to propigate to report.
 
            #### write this all to a proper reference template. ###
            #### REWORK THIS TO WRITE OUT TO REPORT_GENERATOR function
            with ExcelWriter(
            raw_report,
            mode="a",
            engine="openpyxl",
            if_sheet_exists="replace",
            ) as writer:
                room_properties.to_excel(writer, sheet_name='room_properties') 
            

            print("pausing for excel...")
            time.sleep(3)
            #### open excel file and write the reports to PDF ###
            # lol wil output wherever the file explorer in VS is currently. Probably should fix this. 
        
           
            ## - new funnction from stc_calc_dev3 needs to be plugged in here - create_report- passing to report generator v1.py
           
            print("writing to report file:")
            print("Test report file name: ", curr_report_file[0])
            print(raw_report)

            
            ## need logic here to determine if NIC or ASTC or AIIC , and report one or the other 
            if curr_test['ASTC'] == 1:
                create_report(curr_test, room_properties, single_ASTCtest_data, 'ASTC')
            elif curr_test['NIC'] == 1:
                create_report(curr_test, room_properties, single_NICtest_data, 'NIC') # still need NIC report format
            elif curr_test['AIIC'] == 1:
                create_report(curr_test, room_properties, single_AIICtest_data, 'AIIC')
            

        
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

