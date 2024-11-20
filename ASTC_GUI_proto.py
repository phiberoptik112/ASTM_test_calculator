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
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, KeepInFrame, PageBreak 

import matplotlib.pyplot as plt
# Import Report generatorv1.py
# from Report_generatorv1 import create_report,calc_AIIC_val, RAW_SLM_datpull, sanitize_filepath,

from config import *
from data_processor import *
from base_test_reporter import *
# this is a new function combines RT and ASTC datapulls

 # creating a subfunction to return the dataframe from the testplan test list
 # pass it the current test number and the test list
 
### # FUNCTIONS FOR PULLING EXCEL DATA ########
## moved to data_processor.py
    
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

print("GUI import complete")

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
        self.single_test_input_box = TextInput(multiline=False, hint_text='Test Number')
        layout.add_widget(self.single_test_input_box)

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

        # Add test input populate button if debug mode is checked
        if self.debug_check_box.active:
            test_input_button = Button(text='Test Input Populate', on_press=self.populate_test_inputs)
            layout.add_widget(test_input_button)

        return layout
    def populate_test_inputs(self, instance):
        """Populate test input fields with default test paths"""
        # Set default test paths
        self.test_plan_path.text = "/Users/jakepfitsch/Documents/Documents - Jake’s iMac/Python_projects/STC_tester_reporter/Exampledata/TestPlan_ASTM_testingv2.xlsx"
        self.output_folder_path.text = "/Users/jakepfitsch/Documents/Documents - Jake’s iMac/Python_projects/STC_tester_reporter/Exampledata/testeroutputs"
        self.slm_data_d_path.text = "/Users/jakepfitsch/Documents/Documents - Jake’s iMac/Python_projects/STC_tester_reporter/Exampledata/RawData/A_Meter" 
        self.slm_data_e_path.text = "/Users/jakepfitsch/Documents/Documents - Jake’s iMac/Python_projects/STC_tester_reporter/Exampledata/RawData/E_Meter"
        # self.single_test_input_box.text = "1"
        
        # Update status label
        self.status_label.text = "Status: Test inputs populated"
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
            elif instance is self.single_test_input_box:
                self.single_test_input_box = sanitize_filepath(instance.text)
                instance.text = self.single_test_input_box
            
    def show_test_list_popup(self, instance):
        # Create a popup with the TestListPopup content
        test_list_popup = TestListPopup(test_list=self.test_list)
        popup = Popup(title='Test List Editor', content=test_list_popup, size_hint=(None, None), size=(400, 400))
        popup.open()

    def load_data(self, instance):
        # # Access the text from all text boxes
        # seems like just a debug step, may not need this once full output report PDF gen. is working since it pulls directly from the SLM datafiles. 
        try:
            text_input_fields = [self.test_plan_path, self.output_folder_path, self.slm_data_d_path, self.slm_data_e_path, self.single_test_input_box]
            sanitized_values = [sanitize_filepath(field.text) for field in text_input_fields]

        #refactoring to use a dictionary for the input values
            input_values = {
                'test_plan_path': sanitized_values[0],
                'slm_data_d_path': sanitized_values[1],
                'slm_data_e_path': sanitized_values[2],
                'report_output_folder_path': sanitized_values[3],
                'single_test_input_box': sanitized_values[4]
            }
        
            for attr, value in input_values.items():
                setattr(self, attr, value)
                # Update instance variables with the text from the text input boxes

            print('Value from the first text box:', sanitized_values[0])
            print('Value from the second text box:', sanitized_values[1])
            print('Value from the third text box:', sanitized_values[2])
            print('Value from the fourth text box:', sanitized_values[3])
            print('Value from the single test text box:', sanitized_values[4])

            # Display a message in the status label
            self.status_label.text = 'Status: Loading Data...'
            # Add logic to load data from file paths
            print('Arguments received by load_data:', instance, self.test_plan_path, self.slm_data_d_path, self.slm_data_e_path, self.report_output_folder_path)

            # For demonstration purposes, let's just print the file paths
            print('File Paths:', [sanitized_values[0], sanitized_values[1], sanitized_values[2], sanitized_values[3], sanitized_values[4]])
                    # Access the text from all text boxes
            testplan_path = self.test_plan_path
            rawDtestpath = self.slm_data_d_path
            rawEtestpath = self.slm_data_e_path
            outputfolder = self.report_output_folder_path

            self.D_datafiles = [f for f in listdir(rawDtestpath) if isfile(join(rawDtestpath,f))]
            self.E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]
            # self.room_properties, self.test_types = load_test_plan(testplan_path)

            ### each line of the testplan loaded from testplan_path is a test that may require up to 4 test reports to be generated.
            ### TestData -> BaseTestReport subclasses -> PDF Report ###

            ### debug print outs here for the datafiles ### 
            print(self.D_datafiles)
            print(self.E_datafiles)
            print('Arguments received by output_reports:', instance, self.test_plan_path, 
              self.slm_data_d_path, self.slm_data_e_path, self.report_output_folder_path)
            print('--=-=----=-=-=-=-=-=-=-=-=-=-=-=-=-=-===-=-=-'
            )
            testplan_path = self.test_plan_path
            report_output_folder = self.report_output_folder_path
            test_list = pd.read_excel(testplan_path)
            debug = 1 if self.debug_check_box.active else 0

        # List to store all report data objects
            report_data_objects = []

        # Process each test in the test plan
            for _, curr_test in test_list.iterrows():
                if debug:
                    print('Current Test:', curr_test)
                self.status_label.text = f'Status: Processing test: {curr_test["Test_Label"]}'
                print('Creating RoomProperties instance')
                # Create RoomProperties instance

                room_props = RoomProperties(
                    site_name=curr_test['Site_Name'],
                    client_name=curr_test['Client_Name'],
                    source_room=curr_test['Source Room'],
                    receive_room=curr_test['Receiving Room'], 
                    test_date=curr_test['Test Date'],
                    report_date=curr_test['Report Date'],
                    project_name=curr_test['Project Name'],
                    test_label=curr_test['Test_Label'],
                    source_vol=curr_test['source room vol'],
                    receive_vol=curr_test['receive room vol'],
                    partition_area=curr_test['partition area'],
                    partition_dim=curr_test['partition dim'],
                    source_room_finish=curr_test['source room finish'],
                    source_room_name=curr_test['Source Room'],
                    receive_room_finish=curr_test['receive room finish'],
                    receive_room_name=curr_test['Receiving Room'],
                    srs_floor=curr_test['srs floor descrip.'],
                    srs_ceiling=curr_test['srs ceiling descrip.'],
                    srs_walls=curr_test['srs walls descrip.'],
                    rec_floor=curr_test['rec floor descrip.'],
                    rec_ceiling=curr_test['rec ceiling descrip.'],
                    rec_walls=curr_test['rec walls descrip.'],
                    annex_2_used=curr_test['Annex 2 used?'],
                    tested_assembly=curr_test['tested assembly'],  ## potentially redunant - expect to remove
                    test_assembly_type=curr_test['Test assembly Type'],
                    expected_performance=curr_test['expected performance']
                )
                # Initialize total_test_data at class level or method start
                self.total_test_data = pd.DataFrame()

                # Create list of enabled tests based on curr_test values
                enabled_tests = []
                if curr_test['AIIC'] == 1:
                    print('AIIC enabled')
                    enabled_tests.append(TestType.AIIC)
                if curr_test['ASTC'] == 1:
                    print('ASTC enabled')
                    enabled_tests.append(TestType.ASTC)
                if curr_test['NIC'] == 1:
                    print('NIC enabled')
                    enabled_tests.append(TestType.NIC)
                if curr_test['DTC'] == 1:
                    print('DTC enabled')
                    enabled_tests.append(TestType.DTC)

                # Process each enabled test
                for test_type in enabled_tests:
                    try:
                        self.status_label.text = f'Status: Processing {test_type.value} test...'
                        raw_data = self.load_test_data(curr_test, test_type, room_props)
                        
                        # Ensure raw_data is a DataFrame
                        if not isinstance(raw_data, pd.DataFrame):
                            raw_data = pd.DataFrame(raw_data)
                        
                        # Add metadata
                        raw_data['test_label'] = curr_test['Test_Label']
                        raw_data['test_type'] = test_type.value
                        
                        # Append to total_test_data
                        self.total_test_data = pd.concat([self.total_test_data, raw_data], 
                                                        ignore_index=True)
                        
                    except Exception as e:
                        print(f"Error processing {test_type.value} test: {str(e)}")

                    # Display a message in the status label
                    self.status_label.text = 'Status: All test files loaded, ready to generate reports'
                    # open a popup window to display the test list
                    self.show_test_list_popup(self.test_list)
        except Exception as e:
            print(f"Error loading data: {str(e)}")
            self.status_label.text = f'Status: Error loading data: {str(e)}'
    ## TestData -> BaseTestReport subclasses -> PDF Report ###
    def output_all_reports(self, instance):
        print('Arguments received by output_reports:', instance, self.test_plan_path, 
              self.slm_data_d_path, self.slm_data_e_path, self.report_output_folder_path)

        testplan_path = self.test_plan_path
        report_output_folder = self.report_output_folder_path
        test_list = pd.read_excel(testplan_path)
        debug = 1 if self.debug_check_box.active else 0

        # List to store all report data objects
        report_data_objects = []

        # Process each test in the test plan
        for _, curr_test in test_list.iterrows():
            if debug:
                print('Current Test:', curr_test)
            self.status_label.text = f'Status: Processing test: {curr_test["Test_Label"]}'

            # Create RoomProperties instance
            room_props = RoomProperties(
                site_name=curr_test['Site_Name'],
                client_name=curr_test['Client_Name'],
                source_room=curr_test['Source Room'],
                receive_room=curr_test['Receiving Room'], 
                test_date=curr_test['Test Date'],
                report_date=curr_test['Report Date'],
                project_name=curr_test['Project Name'],
                test_label=curr_test['Test_Label'],
                source_vol=curr_test['source room vol'],
                receive_vol=curr_test['receive room vol'],
                partition_area=curr_test['partition area'],
                partition_dim=curr_test['partition dim'],
                source_room_finish=curr_test['source room finish'],
                source_room_name=curr_test['Source Room'],
                receive_room_finish=curr_test['receive room finish'],
                receive_room_name=curr_test['Receiving Room'],
                srs_floor=curr_test['srs floor descrip.'],
                srs_ceiling=curr_test['srs ceiling descrip.'],
                srs_walls=curr_test['srs walls descrip.'],
                rec_floor=curr_test['rec floor descrip.'],
                rec_ceiling=curr_test['rec ceiling descrip.'],
                rec_walls=curr_test['rec walls descrip.'],
                annex_2_used=curr_test['Annex 2 used?'],
                tested_assembly=curr_test['tested assembly'],  ## potentially redunant - expect to remove
                test_assembly_type=curr_test['Test assembly Type'],
                expected_performance=curr_test['expected performance']
            )
            # Create list of enabled test types based on binary values in test columns
            enabled_tests = []
            test_type_columns = {
                'AIIC_test': TestType.AIIC,
                'ASTC_test': TestType.ASTC, 
                'NIC_test': TestType.NIC,
                'DTC_test': TestType.DTC
            }
            
            for col, test_type in test_type_columns.items():
                if curr_test[col] == 1:
                    enabled_tests.append(test_type)
            # Process each enabled test type
            for test_type in enabled_tests:
                self.status_label.text = f'Status: Generating {test_type.value} report...'                    
                # Create appropriate TestData instance
                raw_data = self.load_test_data(curr_test, test_type, room_props)

                # When test_type is AIIC:
                # raw_data will be AIICTestData with base_data + position data
                # When test_type is DTC:
                # raw_data will be DTCTestData with base_data + door measurements
                # When test_type is ASTC:
                # raw_data will be ASTCTestData with base_data only
                # When test_type is NIC:
                # raw_data will be NICTestData with base_data only

                if test_type == TestType.ASTC:
                    ### need to validate the raw data - this is actualy the ASTCTestData instance
                    report = ASTCTestReport.create_report(
                        test_data=raw_data,
                        room_properties=room_props,
                        output_folder=report_output_folder,
                        test_type=test_type
                    )
                    ## save all reports to the report output folder
                    report.save_report()
                    if debug:
                        self.show_test_properties_popup(report)
                elif test_type == TestType.AIIC:
                    report = AIICTestReport.create_report(
                        test_data=raw_data,
                        room_properties=room_props,
                        output_folder=report_output_folder,
                        test_type=test_type
                    )
                    ## save all reports to the report output folder
                    report.save_report()
                    if debug:
                        self.show_test_properties_popup(report)
                elif test_type == TestType.NIC:
                    report = NICTestReport.create_report(
                        test_data=raw_data,
                        room_properties=room_props,
                        output_folder=report_output_folder,
                        test_type=test_type
                    )
                    ## save all reports to the report output folder
                    report.save_report()
                    if debug:
                        self.show_test_properties_popup(report)
                elif test_type == TestType.DTC:
                    report = DTCTestReport.create_report(
                        test_data=raw_data,
                        room_properties=room_props,
                        output_folder=report_output_folder,
                        test_type=test_type
                    )
                    ## save all reports to the report output folder
                    report.save_report()
                    if debug:
                        self.show_test_properties_popup(report)

                # Create ReportData instance ### GENERIC< NEEDED?? ###
                # self.report_data = ReportData(
                #     room_properties=room_props,   
                #     test_data=test_data,
                #     test_type=test_type
                # )
                # Generate report
                # report_path = report_data.generate_report()
                # report_data_objects.append(report_data)

                print(f'Generated {test_type.value} report for test {curr_test["Test Label"]}')


        self.status_label.text = 'Status: All reports generated successfully'
        return report_data_objects

    def load_test_data(self, curr_test: pd.Series, test_type: TestType, room_props: RoomProperties) -> Union[AIICTestData, ASTCTestData, NICTestData, DTCtestData]:
        """Load and format raw test data based on test type"""

        # Base data dictionary for all test types
        base_data = {
            'srs_data': pd.DataFrame(self.RAW_SLM_datapull(curr_test['Source'], '-831_Data.')),
            'recive_data': pd.DataFrame(self.RAW_SLM_datapull(curr_test['Recieve '], '-831_Data.')),
            'bkgrnd_data': pd.DataFrame(self.RAW_SLM_datapull(curr_test['BNL'], '-831_Data.')),
            'rt': pd.DataFrame(self.RAW_SLM_datapull(curr_test['RT'], '-RT_Data.'))
        }

        # Create appropriate test data instance based on type
        if test_type == TestType.AIIC:
            aiic_data = base_data.copy()
            aiic_data.update({
                'AIIC_pos1': pd.DataFrame(self.RAW_SLM_datapull(curr_test['Position1'], '-831_Data.')),
                'AIIC_pos2': pd.DataFrame(self.RAW_SLM_datapull(curr_test['Position2'], '-831_Data.')),
                'AIIC_pos3': pd.DataFrame(self.RAW_SLM_datapull(curr_test['Position3'], '-831_Data.')),
                'AIIC_pos4': pd.DataFrame(self.RAW_SLM_datapull(curr_test['Position4'], '-831_Data.')),
                'AIIC_source': pd.DataFrame(self.RAW_SLM_datapull(curr_test['SourceTap'], '-831_Data.')),
                'AIIC_carpet': pd.DataFrame(self.RAW_SLM_datapull(curr_test['Carpet'], '-831_Data.'))
            })
            return AIICTestData(room_properties=room_props, test_data=aiic_data)

        elif test_type == TestType.DTC:
            dtc_data = base_data.copy()
            dtc_data.update({
                'srs_door_open': pd.DataFrame(self.RAW_SLM_datapull(curr_test['Source_Door_Open'], '-831_Data.')),
                'srs_door_closed': pd.DataFrame(self.RAW_SLM_datapull(curr_test['Source_Door_Closed'], '-831_Data.')),
                'recive_door_open': pd.DataFrame(self.RAW_SLM_datapull(curr_test['Receive_Door_Open'], '-831_Data.')),
                'recive_door_closed': pd.DataFrame(self.RAW_SLM_datapull(curr_test['Receive_Door_Closed'], '-831_Data.'))
            })
            return DTCtestData(room_properties=room_props, test_data=dtc_data)

        elif test_type == TestType.ASTC:
            return ASTCTestData(room_properties=room_props, test_data=base_data)

        elif test_type == TestType.NIC:
            return NICTestData(room_properties=room_props, test_data=base_data)

        else:
            raise ValueError(f"Unsupported test type: {test_type}")

    def calculate_single_test(self, instance):
        """Calculate and generate report for a single test based on test label input"""
        try:
            # Get test label from input
            test_label = self.single_test_text_input.text
            self.status_label.text = f'Status: Processing test {test_label}...'

            # Create results window
            window = tk.Tk()
            window.title(f'Test Results: {test_label}')
            window.geometry("600x400")
            
            status_text = tk.Text(window, height=20, width=70)
            status_text.pack(padx=10, pady=10)

            # Load and find test in test plan
            try:
                test_plan_df = pd.read_excel(self.test_plan_path)
                mask = test_plan_df.applymap(lambda x: test_label in str(x) if pd.notna(x) else False)
                test_row_idx = mask.any(axis=1).idxmax()
                curr_test = test_plan_df.iloc[test_row_idx]
                
                status_text.insert(tk.END, "Found test entry:\n")
                status_text.insert(tk.END, f"{curr_test.to_string()}\n\n")
            except Exception as e:
                raise ValueError(f"Could not find test {test_label} in test plan: {str(e)}")

            # Create room properties dataclass instance
            room_props = RoomProperties(
                site_name=curr_test['Site_Name'],
                client_name=curr_test['Client_Name'],
                source_room=curr_test['Source Room'],
                receive_room=curr_test['Receiving Room'], 
                test_date=curr_test['Test Date'],
                report_date=curr_test['Report Date'],
                project_name=curr_test['Project Name'],
                test_label=curr_test['Test_Label'],
                source_vol=curr_test['source room vol'],
                receive_vol=curr_test['receive room vol'],
                partition_area=curr_test['partition area'],
                partition_dim=curr_test['partition dim'],
                source_room_finish=curr_test['source room finish'],
                source_room_name=curr_test['Source Room'],
                receive_room_finish=curr_test['receive room finish'],
                receive_room_name=curr_test['Receiving Room'],
                srs_floor=curr_test['srs floor descrip.'],
                srs_ceiling=curr_test['srs ceiling descrip.'],
                srs_walls=curr_test['srs walls descrip.'],
                rec_floor=curr_test['rec floor descrip.'],
                rec_ceiling=curr_test['rec ceiling descrip.'],
                rec_walls=curr_test['rec walls descrip.'],
                annex_2_used=curr_test['Annex 2 used?'],
                tested_assembly=curr_test['tested assembly'],  ## potentially redunant - expect to remove
                test_assembly_type=curr_test['Test assembly Type'],
                expected_performance=curr_test['Expected Performance']
            )

            # Determine test type and create appropriate test data object
            test_type = TestType(curr_test['Test Type'])
            test_data = self.load_test_data(curr_test, test_type, room_props)

            status_text.insert(tk.END, f"Generating {test_type.value} report...\n")

            # Generate report
            report = create_report(curr_test, test_data, self.report_output_folder_path)
            
            status_text.insert(tk.END, f"Report generated successfully: {report.get_doc_name()}\n")
            self.status_label.text = f'Status: Test {test_label} report generated successfully'

            window.mainloop()

        except Exception as e:
            error_msg = f"Error processing test {test_label}: {str(e)}"
            self.status_label.text = f'Status: Error - {error_msg}'
            if 'status_text' in locals():
                status_text.insert(tk.END, f"ERROR: {error_msg}\n")
            print(error_msg)


    ### PRIMARILY DEBUG ###
    def excel_import(self):  
        ## PRIMARILY DEUBUG #####
        # import the excel file from the testplan_path
        self.test_plan = pd.read_excel(self.test_plan_path)
        testnums = self.test_list['Test Label'] ## Determines the labels and number of excel files copied
        project_name =  self.test_list['Project Name'] # need to access a cell within template the project name to copy to final reports
        project_name = project_name.iloc[1]
        # Initialize list to store ReportData objects
        report_data_list = []

        # Iterate through each test in testnums
        for idx, test_row in testnums.iterrows():
            # Load test plan data for this row

            ### previous func before classes ### 
            # room_properties, test_types, test_data = load_test_plan(self.test_plan_path, self)

            
            # Check which test types are enabled for this test
            for test_type in TestType:
                if test_types[test_type]:
                    # Create ReportData object for this test type
                    report_data = ReportData(
                        room_properties=room_properties,
                        test_data=test_data,
                        test_type=test_type
                    )
                    report_data_list.append(report_data)
                    
                    # Create popup window for this report data
                    report_window = tk.Toplevel()
                    report_window.title(f"Test {test_row['Test Label']} - {test_type.value}")
                    report_window.geometry("600x400")
                    
                    # Create text widget to display report data
                    text_widget = tk.Text(report_window, height=20, width=70)
                    text_widget.pack(padx=10, pady=10)
                    
                    # Insert report data details
                    text_widget.insert(tk.END, f"Test Label: {test_row['Test Label']}\n")
                    text_widget.insert(tk.END, f"Test Type: {test_type.value}\n\n")
                    text_widget.insert(tk.END, f"Room Properties:\n")
                    text_widget.insert(tk.END, f"Site Name: {room_properties.site_name}\n")
                    text_widget.insert(tk.END, f"Source Room: {room_properties.source_room}\n")
                    text_widget.insert(tk.END, f"Receive Room: {room_properties.receive_room}\n")
                    text_widget.insert(tk.END, f"Test Date: {room_properties.test_date}\n")
                    
                    # Make text widget read-only
                    text_widget.configure(state='disabled')

        print(f"Generated {len(report_data_list)} report data windows")
    def assign_room_properties(self, curr_test: pd.Series):
        """Assign room properties to a RoomProperties instance"""
        print('current test:', curr_test)
        room_props = RoomProperties(
            site_name=curr_test['Site_Name'],
            client_name=curr_test['Client_Name'],
            source_room=curr_test['Source Room'],
            receive_room=curr_test['Receiving Room'], 
            test_date=curr_test['Test Date'],
            report_date=curr_test['Report Date'],
            project_name=curr_test['Project Name'],
            test_label=curr_test['Test_Label'],
            source_vol=curr_test['source room vol'],
            receive_vol=curr_test['receive room vol'],
            partition_area=curr_test['partition area'],
            partition_dim=curr_test['partition dim'],
            source_room_finish=curr_test['source room finish'],
            source_room_name=curr_test['Source Room'],
            receive_room_finish=curr_test['receive room finish'],
            receive_room_name=curr_test['Receiving Room'],
            srs_floor=curr_test['srs floor descrip.'],
            srs_ceiling=curr_test['srs ceiling descrip.'],
            srs_walls=curr_test['srs walls descrip.'],
            rec_floor=curr_test['rec floor descrip.'],
            rec_ceiling=curr_test['rec ceiling descrip.'],
            rec_walls=curr_test['rec walls descrip.'],
            annex_2_used=curr_test['Annex 2 used?'],
            tested_assembly=curr_test['tested assembly'],  ## potentially redunant - expect to remove
            test_assembly_type=curr_test['Test assembly Type'],
            expected_performance=curr_test['expected performance']
        )
        return room_props
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

