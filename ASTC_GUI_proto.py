#ASTC GUI proto



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

from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, KeepInFrame, PageBreak 

import matplotlib.pyplot as plt
# Import Report generatorv1.py
# from Report_generatorv1 import create_report,calc_AIIC_val, RAW_SLM_datpull, sanitize_filepath,

from config import *
# from data_processor import *
from data_processor import (
    TestType, 
    RoomProperties,
    AIICTestData,
    ASTCTestData,
    NICTestData,
    DTCtestData,
    TestData,
    calc_atl_val,
    calc_astc_val,
    format_SLMdata,
    extract_sound_levels,
    calculate_onethird_Logavg
)

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
from kivy.uix.scrollview import ScrollView

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
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.test_data_processor = None
        self.total_test_data = pd.DataFrame()
    def build(self):
        # Initialize path variables as instance variables
        self.test_plan_path = ''
        self.slm_data_d_path = ''
        self.slm_data_e_path = ''
        self.report_output_folder_path = ''
        self.single_test_input_box = ''
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
        self.test_plan_path.text = "./Exampledata/TestPlan_ASTM_testingv2.xlsx"
        self.output_folder_path.text = "./Exampledata/testeroutputs/"
        self.slm_data_d_path.text = "./Exampledata/RawData/A_Meter/" 
        self.slm_data_e_path.text = "./Exampledata/RawData/E_Meter/"
        # self.single_test_input_box.text = "1"
        
        # Update status label
        self.status_label.text = "Status: Test inputs populated"
    def RAW_SLM_datapull(self, find_datafile, datatype):
        """Pull data from SLM files
        
        Args:
            find_datafile (str): File identifier to search for
            datatype (str): Either '-831_Data.' for regular measurements or '-RT_Data.' for reverberation time
            
        Returns:
            pd.DataFrame: Formatted measurement data
        """
        # Use the class's data file paths
        raw_testpaths = {
            'A': self.slm_data_d_path,  # Path for A meter files
            'D': self.slm_data_d_path,  # Path for D meter files
            'E': self.slm_data_e_path   # Path for E meter files
        }

        if self.debug_check_box.active:
            print(f"Looking for file with prefix {find_datafile} in paths:")
            print(raw_testpaths)

        meter_id = find_datafile[0]
        if meter_id not in raw_testpaths:
            raise ValueError(f"Unknown meter identifier: {meter_id}")

        path = raw_testpaths[meter_id]
        
        # Use the already loaded datafiles
        if meter_id in ['A', 'D']:
            datafiles = self.D_datafiles
        elif meter_id == 'E':
            datafiles = self.E_datafiles
        else:
            raise ValueError(f"Unknown meter type: {meter_id}")

        datafile_num = datatype + find_datafile[1:] + '.xlsx'
        slm_found = [x for x in datafiles if datafile_num in x]

        if not slm_found:
            raise ValueError(f"No matching files found for {datafile_num} in {path}")

        full_path = os.path.join(path, slm_found[0])
        
        try:
            if self.debug_check_box.active:
                print(f"Reading file: {full_path}")
                
            if datatype == '-831_Data.':
                # Regular measurement data - use OBA sheet
                print(f"Reading OBA sheet for measurement data")
                df = pd.read_excel(full_path, sheet_name='OBA')
            elif datatype == '-RT_Data.':
                # Reverberation time data - use Summary sheet
                print(f"Reading Summary sheet for RT data")
                df = pd.read_excel(full_path, sheet_name='Summary')
            else:
                raise ValueError(f"Unknown datatype: {datatype}")
            
            if self.debug_check_box.active:
                print(f"DataFrame shape: {df.shape}")
                # print(f"DataFrame columns: {df.columns.tolist()}")
                
            # Verify the DataFrame has expected data
            if df.empty:
                raise ValueError(f"Empty DataFrame loaded from {full_path}")
                
            return df
            
        except Exception as e:
            if self.debug_check_box.active:
                print(f"Error in RAW_SLM_datapull: {str(e)}")
                print(f"File: {find_datafile}, Datatype: {datatype}")
            raise
    
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
            
    ## not using right now 
    def show_test_list_popup(self, test_list):
        # Create popup window
        popup = Popup(title='Loaded Test List',
                     size_hint=(0.9, 0.9))  # 90% of screen size
        
        # Create main layout
        layout = BoxLayout(orientation='vertical')
        
        # Create ScrollView
        scroll_view = ScrollView(
            size_hint=(1, 1),  # Take full size of parent
            do_scroll_x=True,
            do_scroll_y=True,
            bar_width=10,
            scroll_type=['bars']
        )
        
        # Create GridLayout for the test data
        grid = GridLayout(
            cols=len(test_list.columns),
            size_hint=(None, None),  # Neither dimension is bound to parent
            spacing=10,
            padding=10,
            row_default_height=40
        )
        
        # Calculate minimum width needed for content
        col_width = 200  # Adjust this value based on your content
        grid.width = col_width * len(test_list.columns)
        grid.height = grid.row_default_height * (len(test_list) + 1)  # +1 for header
        
        # Add headers
        for col in test_list.columns:
            grid.add_widget(Label(
                text=str(col),
                size_hint=(None, None),
                size=(col_width, grid.row_default_height),
                halign='center'
            ))
        
        # Add data
        for _, row in test_list.iterrows():
            for item in row:
                grid.add_widget(Label(
                    text=str(item),
                    size_hint=(None, None),
                    size=(col_width, grid.row_default_height),
                    halign='center'
                ))
        
        # Add grid to ScrollView
        scroll_view.add_widget(grid)
        
        # Add ScrollView to layout
        layout.add_widget(scroll_view)
        
        # Add close button
        close_button = Button(
            text='Close',
            size_hint=(1, 0.1)
        )
        close_button.bind(on_press=popup.dismiss)
        layout.add_widget(close_button)
        
        # Set popup content
        popup.content = layout
        popup.open()
    ## didn't end up implementing.... could be helpful? but the debug output i've got is good enough atm
    def show_test_data_popup(self, test_type: TestType, raw_data: dict):
        """Display raw test data in a popup window"""
        popup = Popup(
            title=f'{test_type.value} Test Data Preview',
            size_hint=(0.9, 0.9)
        )
        
        # Create main layout
        layout = BoxLayout(orientation='vertical', spacing=10, padding=10)
        
        # Create ScrollView for the data
        scroll = ScrollView(size_hint=(1, 0.9))
        
        # Create content layout
        content = GridLayout(
            cols=1,
            spacing=10,
            padding=10,
            size_hint_y=None
        )
        content.bind(minimum_height=content.setter('height'))
        
        # Add header
        header = Label(
            text=f'Raw Data Preview for {test_type.value} Test',
            size_hint_y=None,
            height=40,
            bold=True
        )
        content.add_widget(header)
        
        # Add each dataframe preview
        for key, df in raw_data.items():
            # Section header
            section = Label(
                text=f'\n{key}:',
                size_hint_y=None,
                height=30,
                bold=True
            )
            content.add_widget(section)
            
            # DataFrame info
            if isinstance(df, pd.DataFrame):
                info_text = (
                    f'Shape: {df.shape}\n'
                    f'Columns: {", ".join(str(col) for col in df.columns[:5])}...\n'
                    f'\nFirst few rows:\n{df.head(3).to_string()}'
                )
            else:
                info_text = "Not a DataFrame"
                
            info_label = TextInput(
                text=info_text,
                size_hint_y=None,
                height=200,
                readonly=True,
                multiline=True
            )
            content.add_widget(info_label)
        
        # Add content to scroll view
        scroll.add_widget(content)
        layout.add_widget(scroll)
        
        # Add close button
        close_button = Button(
            text='Close',
            size_hint=(1, 0.1)
        )
        close_button.bind(on_press=popup.dismiss)
        layout.add_widget(close_button)
        
        # Set popup content and open
        popup.content = layout
        popup.open()


    def load_data(self, instance):
        self.test_data_collection = {}  # Move to class level
        try:
            # Access the text from all text boxes
            # seems like just a debug step, may not need this once full output report PDF gen. is working since it pulls directly from the SLM datafiles. 
            text_input_fields = [self.test_plan_path, self.slm_data_d_path, self.slm_data_e_path, self.output_folder_path, self.single_test_input_box]
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
            print('D_datafiles:', self.D_datafiles)
            print('E_datafiles:', self.E_datafiles)
            print('Arguments received by output_reports:', instance, self.test_plan_path, 
              self.slm_data_d_path, self.slm_data_e_path, self.report_output_folder_path)
            print('--=-=---=-=-=-=--=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-===-=-=-'
            )
            testplan_path = self.test_plan_path
            report_output_folder = self.report_output_folder_path
            self.test_list = pd.read_excel(testplan_path)
            debug = 1 if self.debug_check_box.active else 0


        # Process each test in the test plan
            for _, curr_test in self.test_list.iterrows():
                if self.debug_check_box.active:
                    print('Current Test:', curr_test['Test_Label'])
                self.status_label.text = f'Status: Processing test: {curr_test["Test_Label"]}'
                room_props = self.assign_room_properties(curr_test)
                
                # Create list of enabled tests
                curr_test_data = {}
                for test_type, column in {
                    TestType.AIIC: 'AIIC',
                    TestType.ASTC: 'ASTC',
                    TestType.NIC: 'NIC',
                    TestType.DTC: 'DTC'
                }.items():
                    if curr_test[column] == 1:
                        try:
                            self.status_label.text = f'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-Status: Loading {test_type.value} test data...-=-=-=-=-=-=-=-=-=-==-==-=-=-=-=-=-=-=-'
                            test_data = self.load_test_data(curr_test, test_type, room_props)
                            curr_test_data[test_type] = {
                                'room_properties': room_props,
                                'test_data': test_data
                                }
                            if self.debug_check_box.active:
                                print(f'Successfully loaded {test_type.value} data')
                        except Exception as e:
                            print(f"Error loading {test_type.value} test: {str(e)}")
            
                if curr_test_data:
                    self.test_data_collection[curr_test['Test_Label']] = curr_test_data

            self.status_label.text = 'Status: All test data loaded successfully'
            return True

        except Exception as e:
            print(f"Error in load_data: {str(e)}")
            self.status_label.text = f'Status: Error loading data: {str(e)}'
            return False

    ## TestData -> BaseTestReport subclasses -> PDF Report ###
    def output_all_reports(self, instance):
        """Generate reports using previously loaded test data"""
        if not self.test_data_collection:
            self.status_label.text = 'Status: No test data loaded. Please load data first.'
            return

        try:
            report_data_objects = []
            
            for test_label, test_data_dict in self.test_data_collection.items():
                if self.debug_check_box.active:
                    print(f'Generating reports for test: {test_label}')
                
                for test_type, data in test_data_dict.items():
                    self.status_label.text = f'Status: Generating {test_type.value} report for {test_label}...'
                    
                    try:
                        # Create appropriate report based on test type
                        if test_type == TestType.ASTC:
                            report = ASTCTestReport.create_report(
                                test_data=data['test_data'],
                                room_properties=data['room_properties'],
                                output_folder=self.report_output_folder_path,
                                test_type=test_type
                            )
                        elif test_type == TestType.AIIC:
                            report = AIICTestReport.create_report(
                                test_data=data['test_data'],
                                room_properties=data['room_properties'],
                                output_folder=self.report_output_folder_path,
                                test_type=test_type
                            )
                        elif test_type == TestType.NIC:
                            report = NICTestReport.create_report(
                                test_data=data['test_data'],
                                room_properties=data['room_properties'],
                                output_folder=self.report_output_folder_path,
                                test_type=test_type
                            )
                        elif test_type == TestType.DTC:
                            report = DTCTestReport.create_report(
                                test_data=data['test_data'],
                                room_properties=data['room_properties'],
                                output_folder=self.report_output_folder_path,
                                test_type=test_type
                            )
                        
                        # Save report and show debug info if enabled
                        report.save_report()
                        if self.debug_check_box.active:
                            self.show_test_properties_popup(report)
                        
                        report_data_objects.append(report)
                        print(f'Generated {test_type.value} report for test {test_label}')
                    
                    except Exception as e:
                        print(f"Error generating {test_type.value} report for {test_label}: {str(e)}")
                        continue

            self.status_label.text = 'Status: All reports generated successfully'
            return report_data_objects

        except Exception as e:
            error_msg = f"Error generating reports: {str(e)}"
            print(error_msg)
            self.status_label.text = f'Status: Error - {error_msg}'
            return None

    def load_test_data(self, curr_test: pd.Series, test_type: TestType, room_props: RoomProperties) -> Union[AIICTestData, ASTCTestData, NICTestData, DTCtestData]:
        """Load and format raw test data based on test type"""
            # First verify this test type is enabled for the current test
        test_type_columns = {
            TestType.AIIC: 'AIIC',  # Changed from 'AIIC_test'
            TestType.ASTC: 'ASTC',  # Changed from 'ASTC_test'
            TestType.NIC: 'NIC',    # Changed from 'NIC_test'
            TestType.DTC: 'DTC'     # Changed from 'DTC_test'
        }
        # Check if this test type is enabled
        test_column = test_type_columns.get(test_type)
        if test_column not in curr_test or curr_test[test_column] != 1:
            if self.debug_check_box.active:
                print(f'Test type {test_type.value} is not enabled for test {curr_test["Test_Label"]}')
            raise ValueError(f"Test type {test_type.value} is not enabled for this test")

        # If we get here, the test type is valid and enabled
        if self.debug_check_box.active:
            print(f'Loading base data for test: {curr_test["Test_Label"]} ({test_type.value})')
    
        
        try:
            # Load each DataFrame separately and verify it's valid
            if self.debug_check_box.active:
                print('Loading base data')
            base_data = {
                'srs_data': self.RAW_SLM_datapull(curr_test['Source'], '-831_Data.'),
                'recive_data': self.RAW_SLM_datapull(curr_test['Recieve '], '-831_Data.'),
                'bkgrnd_data': self.RAW_SLM_datapull(curr_test['BNL'], '-831_Data.'),
                'rt': self.RAW_SLM_datapull(curr_test['RT'], '-RT_Data.')
            }

            # Verify each DataFrame is valid
            if self.debug_check_box.active:
                print('Verifying base data')
            for key, df in base_data.items():
                if df is None or df.empty:
                    raise ValueError(f"Invalid or empty DataFrame for {key}")

            # Create appropriate test data instance based on type
            if test_type == TestType.AIIC:
                aiic_data = base_data.copy()
                # Load additional AIIC-specific data
                if self.debug_check_box.active:
                    print('Loading AIIC-specific data')
                additional_data = {
                    'AIIC_pos1': self.RAW_SLM_datapull(curr_test['Position1'], '-831_Data.'),
                    'AIIC_pos2': self.RAW_SLM_datapull(curr_test['Position2'], '-831_Data.'),
                    'AIIC_pos3': self.RAW_SLM_datapull(curr_test['Position3'], '-831_Data.'),
                    'AIIC_pos4': self.RAW_SLM_datapull(curr_test['Position4'], '-831_Data.'),
                    'AIIC_source': self.RAW_SLM_datapull(curr_test['SourceTap'], '-831_Data.'),
                    'AIIC_carpet': self.RAW_SLM_datapull(curr_test['Carpet'], '-831_Data.')
                }
                
                # Verify additional data
                if self.debug_check_box.active:
                    print('Verifying additional data')
                for key, df in additional_data.items():
                    if df is None or df.empty:
                        raise ValueError(f"Invalid or empty DataFrame for {key}")
                
                aiic_data.update(additional_data)
                return AIICTestData(room_properties=room_props, test_data=aiic_data)

            elif test_type == TestType.DTC:
                dtc_data = base_data.copy()
                additional_data = {
                    'srs_door_open': self.RAW_SLM_datapull(curr_test['Source_Door_Open'], '-831_Data.'),
                    'srs_door_closed': self.RAW_SLM_datapull(curr_test['Source_Door_Closed'], '-831_Data.'),
                    'recive_door_open': self.RAW_SLM_datapull(curr_test['Receive_Door_Open'], '-831_Data.'),
                    'recive_door_closed': self.RAW_SLM_datapull(curr_test['Receive_Door_Closed'], '-831_Data.')
                }
                
                # Verify additional data
                if self.debug_check_box.active:
                    print('Verifying additional data')
                for key, df in additional_data.items():
                    if df is None or df.empty:
                        raise ValueError(f"Invalid or empty DataFrame for {key}")
                
                dtc_data.update(additional_data)
                return DTCtestData(room_properties=room_props, test_data=dtc_data)

            elif test_type == TestType.ASTC:
                if self.debug_check_box.active:
                    print('Creating ASTCTestData instance')
                return ASTCTestData(room_properties=room_props, test_data=base_data)

            elif test_type == TestType.NIC:
                if self.debug_check_box.active:
                    print('Creating NICTestData instance')
                return NICTestData(room_properties=room_props, test_data=base_data)

            else:
                raise ValueError(f"Unsupported test type: {test_type}")

        except Exception as e:
            if self.debug_check_box.active:
                print(f"Error loading data: {str(e)}")
                print(f"Current test data: {curr_test}")
            raise

    def calculate_single_test(self, instance):
        """Calculate and generate report for a single test based on test label input"""
        try:
            # Get test label from input
            test_label = self.single_test_input_box.text
            self.status_label.text = f'Status: Processing test {test_label}...'

            # Load and find test in test plan
            try:
                test_plan_df = pd.read_excel(self.test_plan_path)
                mask = test_plan_df.applymap(lambda x: test_label in str(x) if pd.notna(x) else False)
                test_row_idx = mask.any(axis=1).idxmax()
                curr_test = test_plan_df.iloc[test_row_idx]
            except Exception as e:
                raise ValueError(f"Could not find test {test_label} in test plan: {str(e)}")

            # Create room properties instance
            room_props = self.assign_room_properties(curr_test)

            # Determine test type and create appropriate test data object
            test_type = TestType(curr_test['Test Type'])
            test_data = self.load_test_data(curr_test, test_type, room_props)

            self.status_label.text = f"Generating {test_type.value} report..."

            # Generate report
            report = create_report(curr_test, test_data, self.report_output_folder_path)
            
            # Show results in popup
            show_results_popup(test_label, test_data)
            
            self.status_label.text = f'Status: Test {test_label} report generated successfully'

        except Exception as e:
            error_msg = f"Error processing test {test_label}: {str(e)}"
            self.status_label.text = f'Status: Error - {error_msg}'
            print(error_msg)



    @classmethod
    def assign_room_properties(self, curr_test: pd.Series):
        """Assign room properties to a RoomProperties instance"""
        # print('current test:', curr_test)
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
    content = ReportDisplayWindow(report_paths, testplan, test_results)
    popup = Popup(
        title="Report Results",
        content=content,
        size_hint=(0.9, 0.9)
    )
    popup.open()

class ResultWindow(GridLayout):
    def __init__(self, test_label, test_data, **kwargs):
        super().__init__(**kwargs)
        self.cols = 1
        self.padding = 10
        self.spacing = 10

        # Title
        self.add_widget(Label(text=f'Test Results: {test_label}', size_hint_y=None, height=40))
        
        # Scrollable content
        scroll = ScrollView(size_hint=(1, None), size=(400, 300))
        content = GridLayout(cols=1, spacing=10, size_hint_y=None)
        content.bind(minimum_height=content.setter('height'))
        
        # Add test data
        content.add_widget(Label(
            text=f"Test Label: {test_label}\n"
                 f"Room Properties:\n"
                 f"Site Name: {test_data.room_properties.site_name}\n"
                 f"Source Room: {test_data.room_properties.source_room}\n"
                 f"Receive Room: {test_data.room_properties.receive_room}\n"
                 f"Test Date: {test_data.room_properties.test_date}",
            size_hint_y=None,
            height=200,
            text_size=(380, None)
        ))
        
        scroll.add_widget(content)
        self.add_widget(scroll)
        
        # Close button
        close_btn = Button(
            text='Close',
            size_hint_y=None,
            height=40
        )
        close_btn.bind(on_press=self.dismiss)
        self.add_widget(close_btn)
    
    def dismiss(self, instance):
        # Find the parent popup and dismiss it
        parent = instance
        while not isinstance(parent, Popup):
            parent = parent.parent
        parent.dismiss()

def show_results_popup(test_label, test_data):
    content = ResultWindow(test_label, test_data)
    popup = Popup(
        title=f'Test Results: {test_label}',
        content=content,
        size_hint=(0.8, 0.8)
    )
    popup.open()

class ReportDisplayWindow(GridLayout):
    def __init__(self, report_paths, testplan, test_results, **kwargs):
        super().__init__(**kwargs)
        self.cols = 1
        self.padding = 10
        self.spacing = 10

        # Create table for testplan
        table = GridLayout(cols=len(testplan.columns), size_hint_y=None)
        table.bind(minimum_height=table.setter('height'))
        
        # Add headers
        for col in testplan.columns:
            table.add_widget(Label(text=str(col)))
            
        # Add rows
        for _, row in testplan.iterrows():
            for value in row:
                table.add_widget(Label(text=str(value)))
        
        # Add table in a ScrollView
        scroll = ScrollView(size_hint=(1, 0.7))
        scroll.add_widget(table)
        self.add_widget(scroll)
        
        # Add report paths list
        paths_label = Label(
            text='Generated Reports:',
            size_hint_y=None,
            height=30
        )
        self.add_widget(paths_label)
        
        paths_layout = GridLayout(cols=1, size_hint_y=None)
        paths_layout.bind(minimum_height=paths_layout.setter('height'))
        
        for path in report_paths:
            paths_layout.add_widget(Label(
                text=str(path),
                size_hint_y=None,
                height=30
            ))
        
        paths_scroll = ScrollView(size_hint=(1, 0.3))
        paths_scroll.add_widget(paths_layout)
        self.add_widget(paths_scroll)

if __name__ == '__main__':
    FileLoaderApp().run()

