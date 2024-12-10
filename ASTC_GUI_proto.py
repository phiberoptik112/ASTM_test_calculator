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
import tempfile
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
    calculate_onethird_Logavg,
    calc_AIIC_val_claude
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
from kivy.core.window import Window
from kivy.uix.image import Image as KivyImage
from kivy.properties import ObjectProperty
import fitz  # PyMuPDF for PDF preview
import tempfile
import os

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
        # Add debug print at the start of build
        if hasattr(self, 'debug_check_box') and self.debug_check_box.active:
            print("Starting build method")

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
        layout.add_widget(Label(text='PDF Report per test Output Folder path: '))

        # Text Entry Box for the fifth path
        self.output_folder_path = TextInput(multiline=False, hint_text='Output Path')
        self.output_folder_path.bind(on_text_validate=self.on_text_validate)
        layout.add_widget(self.output_folder_path)

        # Load Data Button
        load_button = Button(text='Load Data', on_press=self.load_data)
        layout.add_widget(load_button)

        # Add debug prints around single test input box creation
        if hasattr(self, 'debug_check_box') and self.debug_check_box.active:
            print("Creating single test input box")
            
        # Text entry box for calculating single test results
        layout.add_widget(Label(text='Single Test Results:'))
        self.single_test_input_box = TextInput(
            multiline=False,
            hint_text='Test Number',
            size_hint_y=None,
            height=40
        )
        if hasattr(self, 'debug_check_box') and self.debug_check_box.active:
            print(f"Single test input box created: {type(self.single_test_input_box)}")
            
        layout.add_widget(self.single_test_input_box)

        # Button Layout for test-related buttons
        test_buttons_layout = BoxLayout(
            orientation='horizontal',
            spacing=10,
            size_hint_y=None,
            height=40
        )

        # Calculate Single Test Button
        calculate_single_test_button = Button(
            text='Calculate Single Test',
            size_hint_x=0.33
        )
        calculate_single_test_button.bind(on_press=self.calculate_single_test)
        
        # View Test Data Button
        view_test_data_button = Button(
            text='View Test Data',
            size_hint_x=0.33
        )
        view_test_data_button.bind(on_press=self.view_current_test_data)
        
        # Output Reports Button
        output_button = Button(
            text='Output All Reports',
            size_hint_x=0.33
        )
        output_button.bind(on_press=self.output_all_reports)

        # Add buttons to the horizontal layout
        test_buttons_layout.add_widget(calculate_single_test_button)
        test_buttons_layout.add_widget(view_test_data_button)
        test_buttons_layout.add_widget(output_button)
        
        # Add the button layout to the main layout
        layout.add_widget(test_buttons_layout)

        # Status Text Box
        self.status_label = Label(text='Status: Ready')
        layout.add_widget(self.status_label)

        # Debug section layout
        debug_layout = BoxLayout(
            orientation='horizontal',
            spacing=10,
            size_hint_y=None,
            height=40
        )
        
        # Debug check box
        self.debug_check_box = CheckBox(active=True)
        self.debug_label = Label(text='Debug')
        debug_layout.add_widget(self.debug_check_box)
        debug_layout.add_widget(self.debug_label)

        # Add test input populate button if debug mode is checked
        test_input_button = Button(
            text='Test Input Populate',
            size_hint_x=0.5
        )
        test_input_button.bind(on_press=self.populate_test_inputs)
        debug_layout.add_widget(test_input_button)

        layout.add_widget(debug_layout)

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
                self.test_plan_path.text = sanitize_filepath(instance.text)
            elif instance is self.output_folder_path:
                self.output_folder_path.text = sanitize_filepath(instance.text)
            elif instance is self.slm_data_d_path:
                self.slm_data_d_path.text = sanitize_filepath(instance.text)
            elif instance is self.slm_data_e_path:
                self.slm_data_e_path.text = sanitize_filepath(instance.text)
            elif instance is self.single_test_input_box:
                self.single_test_input_box.text = sanitize_filepath(instance.text)
            
    ## not using right now...?
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
                                print(f'Current test data structure:')
                                print(f"- Test type: {test_type}")
                                print(f"- Test data type: {type(test_data)}")
                                print(f"- Attributes: {dir(test_data)}")
                                # print(f"- Has room_properties: {'room_properties' in dir(test_data)}")
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
            for test_label, test_data_dict in self.test_data_collection.items():
                if self.debug_check_box.active:
                    print(f'Generating reports for test: {test_label}')
                
                for test_type, data in test_data_dict.items():
                    self.status_label.text = f'Status: Generating {test_type.value} report for {test_label}...'
                    
                    try:
                        if self.debug_check_box.active:
                            print(f"Debug: Test data structure:")
                            print(f"- Test type: {test_type}")
                            print(f"- Test data type: {type(data['test_data'])}")
                            print(f"- Attributes: {dir(data['test_data'])}")
                            print(f"- Has room_properties: {'room_properties' in dir(data['test_data'])}")
                            if hasattr(data['test_data'], 'room_properties'):
                                print(f"- Room properties type: {type(data['test_data'].room_properties)}")
                    
                        report_class = {
                            TestType.ASTC: ASTCTestReport,
                            TestType.AIIC: AIICTestReport,
                            TestType.NIC: NICTestReport,
                            TestType.DTC: DTCTestReport
                        }.get(test_type)
                        
                        if report_class:
                            report = report_class.create_report(
                                test_data=data['test_data'],
                                output_folder=self.report_output_folder_path,
                                test_type=test_type
                            )
                            
                            # Get the PDF path
                            pdf_path = os.path.join(
                                self.report_output_folder_path,
                                f"{test_label}_{test_type.value}_Report.pdf"
                            )
                            
                            print(f'Generated {test_type.value} report for test {test_label}')
                            
                            # Show results popup with PDF path
                            show_results_popup(test_label, data['test_data'], pdf_path)
                    
                    except Exception as e:
                        print(f"Error generating {test_type.value} report for {test_label}: {str(e)}")
                        continue

            self.status_label.text = 'Status: All reports generated successfully'
            return True

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

            # Create appropriate test data instance 
            # based on type
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

    def view_current_test_data(self, instance):
        """Show the test data for the currently selected test"""
        try:
            # Add more detailed debugging
            if self.debug_check_box.active:
                print("Debug info for view_current_test_data:")
                print(f"self.single_test_input_box exists: {hasattr(self, 'single_test_input_box')}")
                if hasattr(self, 'single_test_input_box'):
                    print(f"Type of single_test_input_box: {type(self.single_test_input_box)}")
                    print(f"Dir of single_test_input_box: {dir(self.single_test_input_box)}")

            if not hasattr(self, 'single_test_input_box'):
                raise ValueError("Test input box does not exist")
                
            if not isinstance(self.single_test_input_box, TextInput):
                raise ValueError(f"Test input box is wrong type: {type(self.single_test_input_box)}")
                
            # Get test label from input box
            test_label = self.single_test_input_box.text.strip()
            
            if self.debug_check_box.active:
                print(f"Test label retrieved: {test_label}")
            
            if not test_label:
                self.status_label.text = 'Status: Please enter a test label'
                return
                
            if not hasattr(self, 'test_data_collection') or not self.test_data_collection:
                self.status_label.text = 'Status: No test data loaded. Please load data first.'
                return
                
            if test_label not in self.test_data_collection:
                self.status_label.text = f'Status: Test {test_label} not found in loaded data'
                return
                
            # Get test data for the specified label
            test_data_dict = self.test_data_collection[test_label]
            
            if self.debug_check_box.active:
                print(f"Found test data for {test_label}:")
                print(f"Test types: {list(test_data_dict.keys())}")
            
            # Show popup for each test type in the test data
            for test_type, data in test_data_dict.items():
                if self.debug_check_box.active:
                    print(f"Showing data for test type: {test_type}")
                    
                pdf_path = os.path.join(
                    self.report_output_folder_path,
                    f"{test_label}_{test_type.value}_Report.pdf"
                )
                
                if self.debug_check_box.active:
                    print(f"Looking for PDF at: {pdf_path}")
                    print(f"Test data contents: {data['test_data']}")
                
                show_results_popup(test_label, data['test_data'], pdf_path)
                
            self.status_label.text = f'Status: Showing data for test {test_label}'
            
        except Exception as e:
            error_msg = f"Error viewing test data: {str(e)}"
            print(error_msg)
            if self.debug_check_box.active:
                print(f"Full error details: {e}")
                import traceback
                traceback.print_exc()
            self.status_label.text = f'Status: Error - {error_msg}'

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

class PDFPreviewWindow(GridLayout):
    def __init__(self, pdf_path, **kwargs):
        super().__init__(**kwargs)
        self.cols = 1
        self.padding = 10
        self.spacing = 10
        
        # Create scrollable view
        scroll = ScrollView(size_hint=(1, 0.9))
        content = GridLayout(cols=1, spacing=10, size_hint_y=None)
        content.bind(minimum_height=content.setter('height'))
        
        try:
            # Open PDF and convert first page to image
            doc = fitz.open(pdf_path)
            for page_num in range(min(len(doc), 4)):  # Preview up to 4 pages
                page = doc[page_num]
                pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))  # Scale down to 50%
                
                # Save to temporary image file
                tmp_img = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                pix.save(tmp_img.name)
                
                # Add image to layout
                img = KivyImage(source=tmp_img.name, size_hint_y=None)
                img.height = Window.height * 0.8  # Scale to 80% of window height
                content.add_widget(img)
                
            doc.close()
        except Exception as e:
            content.add_widget(Label(text=f"Error loading PDF: {str(e)}"))
        
        scroll.add_widget(content)
        self.add_widget(scroll)
        
        # Close button
        close_btn = Button(
            text='Close',
            size_hint_y=0.1
        )
        close_btn.bind(on_press=self.dismiss)
        self.add_widget(close_btn)
    
    def dismiss(self, instance):
        parent = instance
        while not isinstance(parent, Popup):
            parent = parent.parent
        parent.dismiss()

class ResultWindow(GridLayout):
    def __init__(self, test_label, test_data, pdf_path=None, **kwargs):
        super().__init__(**kwargs)
        self.cols = 1
        self.padding = 10
        self.spacing = 10

        # Title
        self.add_widget(Label(
            text=f'Test Results: {test_label}',
            size_hint_y=None,
            height=40,
            bold=True
        ))
        
        # Scrollable content
        scroll = ScrollView(size_hint=(1, None), size=(400, 300))
        content = GridLayout(cols=1, spacing=10, size_hint_y=None)
        content.bind(minimum_height=content.setter('height'))
        
        # Add test type and basic info
        content.add_widget(Label(
            text=f"\nTest Type: {test_data.__class__.__name__}",
            size_hint_y=None,
            height=30,
            bold=True
        ))

        # Display SLM data based on test type
        if hasattr(test_data, 'srs_data'):
            content.add_widget(Label(
                text="\nSource Room Measurements:",
                size_hint_y=None,
                height=30,
                bold=True
            ))
            # Format and display source room data
            srs_data = format_SLMdata(test_data.srs_data)
            content.add_widget(Label(
                text=str(srs_data),
                size_hint_y=None,
                height=100,
                text_size=(380, None)
            ))

        if hasattr(test_data, 'recive_data'):
            content.add_widget(Label(
                text="\nReceiver Room Measurements:",
                size_hint_y=None,
                height=30,
                bold=True
            ))
            # Format and display receiver room data
            rec_data = format_SLMdata(test_data.recive_data)
            content.add_widget(Label(
                text=str(rec_data),
                size_hint_y=None,
                height=100,
                text_size=(380, None)
            ))

        # For AIIC tests, show tapping positions
        if hasattr(test_data, 'pos1'):
            content.add_widget(Label(
                text="\nTapping Machine Positions:",
                size_hint_y=None,
                height=30,
                bold=True
            ))
            for i in range(1, 5):
                pos_data = getattr(test_data, f'pos{i}', None)
                if pos_data is not None:
                    formatted_pos = format_SLMdata(pos_data)
                    content.add_widget(Label(
                        text=f"Position {i}:\n{str(formatted_pos)}",
                        size_hint_y=None,
                        height=100,
                        text_size=(380, None)
                    ))

        # Show background noise data
        if hasattr(test_data, 'bkgrnd_data'):
            content.add_widget(Label(
                text="\nBackground Noise Measurements:",
                size_hint_y=None,
                height=30,
                bold=True
            ))
            bkg_data = format_SLMdata(test_data.bkgrnd_data)
            content.add_widget(Label(
                text=str(bkg_data),
                size_hint_y=None,
                height=100,
                text_size=(380, None)
            ))

        # Show RT data
        if hasattr(test_data, 'rt'):
            content.add_widget(Label(
                text="\nReverberation Time Data:",
                size_hint_y=None,
                height=30,
                bold=True
            ))
            rt_data = test_data.rt['Unnamed: 10'][25:41]/1000  # Convert to seconds
            content.add_widget(Label(
                text=str(rt_data),
                size_hint_y=None,
                height=100,
                text_size=(380, None)
            ))
        
        # Add room properties if available
        if hasattr(test_data, 'room_properties'):
            props = test_data.room_properties
            content.add_widget(Label(
                text="\nRoom Properties:",
                size_hint_y=None,
                height=30,
                bold=True
            ))
            for prop_name, prop_value in vars(props).items():
                content.add_widget(Label(
                    text=f"{prop_name}: {str(prop_value)}",
                    size_hint_y=None,
                    height=30,
                    text_size=(380, None)
                ))
        
        scroll.add_widget(content)
        self.add_widget(scroll)
        
        # Button layout
        button_layout = BoxLayout(
            orientation='horizontal',
            size_hint_y=None,
            height=40,
            spacing=10
        )
        
        # Close button
        close_btn = Button(
            text='Close',
            size_hint_x=0.5
        )
        close_btn.bind(on_press=self.dismiss)
        
        # Preview PDF button
        preview_btn = Button(
            text='Preview PDF',
            size_hint_x=0.5
        )
        preview_btn.bind(on_press=lambda x: self.show_pdf_preview(pdf_path))
        
        button_layout.add_widget(close_btn)
        button_layout.add_widget(preview_btn)
        self.add_widget(button_layout)
    
    def dismiss(self, instance):
        parent = instance
        while not isinstance(parent, Popup):
            parent = parent.parent
        parent.dismiss()
    
    def show_pdf_preview(self, pdf_path):
        if pdf_path and os.path.exists(pdf_path):
            content = PDFPreviewWindow(pdf_path)
            preview_popup = Popup(
                title='PDF Preview',
                content=content,
                size_hint=(0.9, 0.9)
            )
            preview_popup.open()
        else:
            error_popup = Popup(
                title='Error',
                content=Label(text='PDF file not found'),
                size_hint=(0.5, 0.3)
            )
            error_popup.open()

def show_results_popup(test_label, test_data, pdf_path=None):
    content = ResultWindow(test_label, test_data, pdf_path)
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

