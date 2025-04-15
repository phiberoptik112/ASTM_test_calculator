import traceback
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.checkbox import CheckBox
from kivy.uix.popup import Popup
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.core.window import Window
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from kivy.uix.image import Image as KivyImage
from kivy.uix.widget import Widget
from kivy.graphics import Color, Rectangle
from kivy.graphics.instructions import Canvas
from enum import Enum
from typing import Dict, List, Optional, Union
import pandas as pd
import tempfile
import os
from os import listdir
from os.path import isfile, join
import fitz
import matplotlib.pyplot as plt
from src.core.test_data_manager import TestDataManager
from src.gui.test_plan_input import TestPlanInputWindow
from src.gui.analysis_dashboard import ResultsAnalysisDashboard
from src.gui.test_plan_manager import TestPlanManagerWindow
from kivy.uix.filechooser import FileChooserListView

import logging
logging.getLogger('matplotlib.font_manager').disabled = True
logging.getLogger('matplotlib').setLevel(logging.WARNING)
from src.core.data_processor import (
    TestType, 
    RoomProperties,
    AIICTestData,
    ASTCTestData,
    NICTestData,
    DTCtestData,    
    TestData,
    calc_NR_new,
    calc_AIIC_val_claude,
    calc_atl_val,
    calc_astc_val,
    format_SLMdata,
    calculate_onethird_Logavg,
    sanitize_filepath
)
from src.core.test_processor import TestProcessor
import numpy as np
from src.reports.base_test_reporter import *

class MainWindow(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.spacing = 10
        self.padding = 10
        
        # Initialize collections
        self.test_data_collection = {}
        
        # Initialize TestDataManager with debug mode
        self.test_data_manager = TestDataManager(debug_mode=False)  # Can be tied to checkbox later
        
        # Create UI sections
        self._create_input_section()
        self._create_test_controls()
        self._create_analysis_section()
        self._create_status_section()

    def _create_input_section(self):
        """Create input fields section"""
        input_grid = GridLayout(
            cols=2,
            spacing=5,
            size_hint_y=0.4
        )
        
        # Test Plan Input
        input_grid.add_widget(Label(text='Test Plan:'))
        test_plan_layout = BoxLayout(orientation='horizontal', spacing=5)
        
        # Create text input and file picker button layout
        path_picker_layout = BoxLayout(orientation='horizontal', size_hint_x=0.7)
        
        self.test_plan_path = TextInput(
            multiline=False,
            hint_text='Path to test plan Excel file',
            size_hint_x=0.85
        )
        path_picker_layout.add_widget(self.test_plan_path)
        
        # Add file picker button
        file_picker_btn = Button(
            text='...',
            size_hint_x=0.15
        )
        file_picker_btn.bind(on_press=lambda x: self.show_file_picker(self.test_plan_path, [('Excel Files', '*.xlsx'), ('All Files', '*.*'), ('CSV Files', '*.csv')]))
        path_picker_layout.add_widget(file_picker_btn)
        
        test_plan_layout.add_widget(path_picker_layout)
        
        # Add Test Plan GUI Button
        self.test_plan_gui_button = Button(
            text='Open Test Plan GUI',
            size_hint_x=0.3
        )
        self.test_plan_gui_button.bind(on_press=self.show_test_plan_input)
        test_plan_layout.add_widget(self.test_plan_gui_button)
        
        input_grid.add_widget(test_plan_layout)
        
        # SLM Data Inputs
        input_grid.add_widget(Label(text='SLM Data Meter 1:'))
        self.slm_data_1_path = TextInput(
            multiline=False,
            hint_text='Path to first SLM data file'
        )
        input_grid.add_widget(self.slm_data_1_path)
        
        input_grid.add_widget(Label(text='SLM Data Meter 2:'))
        self.slm_data_2_path = TextInput(
            multiline=False,
            hint_text='Path to second SLM data file'
        )
        input_grid.add_widget(self.slm_data_2_path)
        
        # Output Folder
        input_grid.add_widget(Label(text='Output Folder:'))
        self.output_path = TextInput(
            multiline=False,
            hint_text='Path for output reports'
        )
        input_grid.add_widget(self.output_path)
        
        self.add_widget(input_grid)
        
        # Button Layout
        button_layout = BoxLayout(
            orientation='horizontal',
            spacing=5,
            size_hint_y=0.1
        )
        
        # Load Data Button
        self.load_button = Button(
            text='Load Data',
            size_hint_x=0.5
        )
        self.load_button.bind(on_press=self.load_data)
        button_layout.add_widget(self.load_button)
        
        # Populate Test Inputs Button
        self.populate_button = Button(
            text='Load Example Data',
            size_hint_x=0.5
        )
        self.populate_button.bind(on_press=self.populate_test_inputs)
        button_layout.add_widget(self.populate_button)
        
        self.add_widget(button_layout)

    def _create_test_controls(self):
        """Create test control section"""
        test_controls = BoxLayout(
            orientation='vertical',
            spacing=5,
            size_hint_y=0.3
        )
        
        # Single Test Input
        test_input_layout = BoxLayout(size_hint_y=0.5)
        test_input_layout.add_widget(Label(
            text='Test Number:',
            size_hint_x=0.3
        ))
        self.test_number_input = TextInput(
            multiline=False,
            hint_text='Enter test number and test type'
        )
        test_input_layout.add_widget(self.test_number_input)
        test_controls.add_widget(test_input_layout)
        
        # Test Control Buttons
        button_layout = BoxLayout(spacing=5, size_hint_y=0.5)
        
        self.calc_button = Button(text='Calculate Test')
        self.calc_button.bind(on_press=self.calculate_single_test)
        button_layout.add_widget(self.calc_button)
        
        self.view_button = Button(text='View Data')
        self.view_button.bind(on_press=self.view_test_data)
        button_layout.add_widget(self.view_button)
        
        self.report_button = Button(text='Generate Reports')
        self.report_button.bind(on_press=self.generate_reports)
        button_layout.add_widget(self.report_button)
        
        # Add Plot Data button
        self.plot_button = Button(text='Raw Loaded Data Plot')
        self.plot_button.bind(on_press=self.show_plot_selection)
        button_layout.add_widget(self.plot_button)
        
        test_controls.add_widget(button_layout)
        self.add_widget(test_controls)

    def _create_analysis_section(self):
        """Create analysis dashboard section with integrated update methods"""
        self.analysis_tabs = TabbedPanel(size_hint_y=0.5)
        
        # Test Plan Tab with integrated update functionality
        self.test_plan_tab = TabbedPanelItem(text='Test Plan')
        self.test_plan_tab.do_default_tab = True
        
        self.test_plan_scroll = ScrollView()
        self.test_plan_grid = GridLayout(
            cols=1,
            spacing=2,
            size_hint_y=None,
            padding=5
        )
        self.test_plan_grid.bind(minimum_height=self.test_plan_grid.setter('height'))
        
        self.test_plan_scroll.add_widget(self.test_plan_grid)
        self.test_plan_tab.add_widget(self.test_plan_scroll)
        self.analysis_tabs.add_widget(self.test_plan_tab)
        
        # Results Tab - reuse existing dashboard
        self.results_tab = TabbedPanelItem(text='Results Dashboard')
        self.results_dashboard = ResultsAnalysisDashboard(self.test_data_manager)
        self.results_tab.add_widget(self.results_dashboard)
        self.analysis_tabs.add_widget(self.results_tab)
        
        # Raw Data Tab - integrate with view_test_data functionality
        self.raw_data_tab = TabbedPanelItem(text='Raw Data')
        self.raw_data_layout = BoxLayout(orientation='vertical')
        
        # Add test selection controls
        test_controls = BoxLayout(
            orientation='horizontal',
            size_hint_y=None,
            height=40
        )
        self.raw_data_test_input = TextInput(
            multiline=False,
            hint_text='Enter test number and test type',
            size_hint_x=0.7
        )
        view_button = Button(
            text='View Data',
            size_hint_x=0.3
        )
        view_button.bind(on_press=self.view_test_data)
        
        test_controls.add_widget(self.raw_data_test_input)
        test_controls.add_widget(view_button)
        
        self.raw_data_layout.add_widget(test_controls)
        
        # Scrollable data view
        self.raw_data_view = ScrollView()
        self.raw_data_grid = GridLayout(
            cols=1,
            spacing=2,
            size_hint_y=None
        )
        self.raw_data_grid.bind(minimum_height=self.raw_data_grid.setter('height'))
        self.raw_data_view.add_widget(self.raw_data_grid)
        self.raw_data_layout.add_widget(self.raw_data_view)
        
        self.raw_data_tab.add_widget(self.raw_data_layout)
        self.analysis_tabs.add_widget(self.raw_data_tab)
        
        self.add_widget(self.analysis_tabs)

    def _create_status_section(self):
        """Create status and debug section"""
        status_layout = BoxLayout(
            orientation='horizontal',
            spacing=5,
            size_hint_y=0.1
        )
        
        # Status Label
        self.status_label = Label(
            text='Status: Ready',
            size_hint_x=0.7
        )
        status_layout.add_widget(self.status_label)
        
        # Debug Controls
        debug_layout = BoxLayout(size_hint_x=0.3)
        self.debug_checkbox = CheckBox(active=True)
        debug_layout.add_widget(self.debug_checkbox)
        debug_layout.add_widget(Label(text='Debug Mode'))
        status_layout.add_widget(debug_layout)
        
        self.add_widget(status_layout)

    def _show_error(self, message: str):
        """Display error message in popup with more detailed formatting"""
        content_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        # Add scrollable text for long error messages
        scroll_view = ScrollView(size_hint=(1, 0.8))
        error_label = Label(
            text=message, 
            size_hint_y=None,
            text_size=(380, None)
        )
        error_label.bind(texture_size=error_label.setter('size'))
        scroll_view.add_widget(error_label)
        content_layout.add_widget(scroll_view)
        
        # Add close button
        close_button = Button(text="Close", size_hint=(1, 0.2))
        content_layout.add_widget(close_button)
        
        popup = Popup(
            title='Error',
            content=content_layout,
            size_hint=(None, None),
            size=(400, 300)
        )
        
        close_button.bind(on_press=popup.dismiss)
        popup.open()

    def populate_test_inputs(self, instance):
        """Populate input fields with example data paths"""
        # self.test_plan_path.text = "./Exampledata/TestPlan_ASTM_testingv2.csv"
        # self.slm_data_1_path.text = "./Exampledata/RawData/A_Meter/"
        # self.slm_data_2_path.text = "./Exampledata/RawData/E_Meter/"
        # self.output_path.text = "./Exampledata/testeroutputs/"

        ## kaanapali shores
        self.test_plan_path.text = "./example_data_kaanapali/TestPlan_kaanapali_AIIC_testingv2.csv"
        self.slm_data_1_path.text = "./example_data_kaanapali/rawData/meterD/"
        self.slm_data_2_path.text = "./example_data_kaanapali/rawData/meterE/"
        self.output_path.text = "./example_data_kaanapali/reports_outputs/"
        self.status_label.text = "Status: Test inputs populated with example data"

    def load_data(self, instance):
        """Load and process all test data with improved error handling"""
        try:
            debug_mode = self.debug_checkbox.active
            
            # Get paths from input fields
            test_plan_path = self.test_plan_path.text
            slm_data_d_path = self.slm_data_1_path.text
            slm_data_e_path = self.slm_data_2_path.text
            report_output_path = self.output_path.text
            
            # Validate paths
            if not all([test_plan_path, slm_data_d_path, slm_data_e_path, report_output_path]):
                raise ValueError("All paths must be provided")
                
            # Check if the test plan file exists
            if not os.path.exists(test_plan_path):
                raise FileNotFoundError(f"Test plan file not found: {test_plan_path}")
            
            # Check if the data directories exist
            for path, name in [
                (slm_data_d_path, "SLM Data Meter 1"),
                (slm_data_e_path, "SLM Data Meter 2")
            ]:
                if not os.path.exists(path):
                    raise FileNotFoundError(f"{name} directory not found: {path}")
                if not os.path.isdir(path):
                    raise ValueError(f"{name} path must be a directory: {path}")
            
            # Create output directory if it doesn't exist
            if not os.path.exists(report_output_path):
                try:
                    os.makedirs(report_output_path, exist_ok=True)
                    if debug_mode:
                        print(f"Created output directory: {report_output_path}")
                except PermissionError:
                    raise PermissionError(f"Cannot create output directory: {report_output_path}. Please check permissions.")
            
            if debug_mode:
                print("\nLoading data with paths:")
                print(f"Test plan: {test_plan_path}")
                print(f"SLM Data D: {slm_data_d_path}")
                print(f"SLM Data E: {slm_data_e_path}")
                print(f"Output: {report_output_path}")
            
            # Set paths in test data manager
            self.test_data_manager.set_data_paths(
                test_plan_path=test_plan_path,
                meter_d_path=slm_data_d_path,
                meter_e_path=slm_data_e_path,
                output_path=report_output_path
            )
            
            # Load test plan
            try:
                if debug_mode:
                    print("\nLoading test plan...")
                self.status_label.text = 'Status: Loading test plan...'
                self.test_data_manager.load_test_plan()
                
                # Process test data
                if debug_mode:
                    print("\nProcessing test data...")
                    print(f"Test plan path: {test_plan_path}")
                    print(f"SLM data D path: {slm_data_d_path}")
                    print(f"SLM data E path: {slm_data_e_path}")
                    print(f"Report output path: {report_output_path}")
                self.status_label.text = 'Status: Processing test data...'
                self.test_data_manager.process_test_data()
                
                # Update all displays
                self.update_displays()
                
                # Add debug output to verify data was loaded
                if debug_mode:
                    print("inside debug mode, going to gather the test collection")
                    test_collection = self.test_data_manager.get_test_collection()
                    print("\nLoaded test collection:")
                    print(f"Number of tests: {len(test_collection)}")
                    for test_label, test_data in test_collection.items():
                        print(f"\nTest: {test_label}")
                        print(f"Test types: {list(test_data.keys())}")
                
                self.status_label.text = 'Status: All test data loaded successfully'
                
                # Show plot selection popup after successful load
                self.show_plot_selection(instance)
                
                return True
                
            except ValueError as e:
                error_msg = str(e)
                
                # Check for specific Excel extension issue
                if "File is not a zip file" in error_msg or "appears to be corrupted" in error_msg:
                    self._show_excel_extension_error(error_msg)
                    return False
                
                # Provide specific suggestions based on error types
                if "corrupted" in error_msg and "Excel file" in error_msg:
                    error_msg += "\n\nSuggestions:\n- Open the file in Excel/LibreOffice and save it in a new format (CSV or XLSX)\n- Check if the file is password protected\n- Try using a different Excel file"
                elif "No matching files found" in error_msg:
                    error_msg += "\n\nSuggestions:\n- Verify SLM data file naming format\n- Check paths to SLM data directories\n- Ensure data files are in the expected format (.xlsx)"
                
                raise ValueError(error_msg)
                
        except FileNotFoundError as e:
            error_msg = f"File not found: {str(e)}"
            if debug_mode:
                print(f"\nERROR: {error_msg}")
            self._show_error(error_msg + "\n\nPlease check that all paths are correct and files/directories exist.")
            self.status_label.text = f"Status: Error - {error_msg}"
            return False
        except PermissionError as e:
            error_msg = f"Permission error: {str(e)}"
            if debug_mode:
                print(f"\nERROR: {error_msg}")
            self._show_error(error_msg + "\n\nPlease check file/directory permissions.")
            self.status_label.text = f"Status: Error - {error_msg}"
            return False
        except ValueError as e:
            error_msg = f"Error loading data: {str(e)}"
            if debug_mode:
                print(f"\nERROR: {error_msg}")
                print("Stack trace:")
                import traceback
                traceback.print_exc()
            self._show_error(error_msg)
            self.status_label.text = f"Status: {error_msg}"
            return False
        except Exception as e:
            error_msg = f"Unexpected error: {str(e)}"
            if debug_mode:
                print(f"\nERROR: {error_msg}")
                print("Stack trace:")
                import traceback
                traceback.print_exc()
            self._show_error(error_msg + "\n\nPlease check the debug output for more details.")
            self.status_label.text = f"Status: Error - {error_msg}"
            return False
    # may be a redundant function, but keeping it for now
    def assign_room_properties(self, test_row: pd.Series) -> RoomProperties:
        """Create RoomProperties from test row data"""
        print('test_row:', test_row)
        return RoomProperties(
            site_name=test_row['Site_Name'],
            client_name=test_row['Client_Name'],
            source_room=test_row['Source Room'],
            receive_room=test_row['Receiving Room'], 
            test_date=test_row['Test Date'],
            report_date=test_row['Report Date'],
            project_name=test_row['Project Name'],
            test_label=test_row['Test_Label'],
            source_vol=test_row['source room vol'],
            receive_vol=test_row['receive room vol'],
            partition_area=test_row['partition area'],
            partition_dim=test_row['partition dim'],
            source_room_finish=test_row['source room finish'],
            source_room_name=test_row['Source Room'],
            receive_room_finish=test_row['receive room finish'],
            receive_room_name=test_row['Receiving Room'],
            srs_floor=test_row['srs floor descrip.'],
            srs_ceiling=test_row['srs ceiling descrip.'],
            srs_walls=test_row['srs walls descrip.'],
            rec_floor=test_row['rec floor descrip.'],
            rec_ceiling=test_row['rec ceiling descrip.'],
            rec_walls=test_row['rec walls descrip.'],
            annex_2_used=test_row['Annex 2 used?'],
            tested_assembly=test_row['tested assembly'],  ## potentially redunant - expect to remove
            test_assembly_type=test_row['Test assembly Type'],
            expected_performance=test_row['expected performance']
        )

    def calculate_single_test(self, instance):
        """Calculate results for a single test"""
        try:
            test_number = self.test_number_input.text
            if not test_number:
                raise ValueError("Please enter a test number")
                
            # Implementation similar to ASTC_GUI_proto's calculate_single_test
            # but integrated with TestDataManager
            
            self.status_label.text = 'Status: Test calculated successfully'
            
        except Exception as e:
            self.status_label.text = f'Status: Error calculating test - {str(e)}'
            if self.debug_checkbox.active:
                print(f"Error calculating test: {str(e)}")

    def view_test_data(self, instance):
        """View raw data for current test - integrated with Raw Data tab"""
        try:
            input_text = self.raw_data_test_input.text or self.test_number_input.text
            test_number, test_type = input_text.split(',') if ',' in input_text else (input_text, None)
            test_number = test_number.strip()
            if test_type:
                test_type = test_type.strip()
            if not test_number:
                raise ValueError("Please enter a test number and test type")
        
            # Clear existing data
            self.raw_data_grid.clear_widgets()
            
            # Get test data
            test_data = self.test_data_manager.get_test_data(test_number, test_type)
            if test_data:
                # Display test information
                self._add_data_section("Test Information", {
                    "Test Number": test_number,
                    "Test Type": test_type,
                    "Date": test_data.get('test_date', 'Unknown')
                })
                
                # Display raw measurements
                if 'measurements' in test_data:
                    self._add_data_section("Raw Measurements", test_data['measurements'])
                
                # Display calculated values
                if 'calculated_values' in test_data:
                    self._add_data_section("Calculated Values", test_data['calculated_values'])
                
                # Switch to raw data tab
                self.raw_data_tab.state = 'down'
            else:
                self.raw_data_grid.add_widget(Label(
                    text=f'No data found for test {test_number}',
                    size_hint_y=None,
                    height=40
                ))
                
        except Exception as e:
            self.status_label.text = f'Status: Error viewing test data - {str(e)}'
            if self.debug_checkbox.active:
                print(f"Error viewing test data: {str(e)}")

    def show_test_plan_input(self, instance):
        """Show test plan input window"""
        content = TestPlanManagerWindow(test_data_manager=self.test_data_manager)
        self.test_plan_popup = Popup(
            title='Test Plan Manager',
            content=content,
            size_hint=(0.9, 0.9)
        )
        self.test_plan_popup.open()

    def on_test_plan_save(self, room_properties: RoomProperties, test_type: TestType):
        """Handle saved test plan data"""
        try:
            # Add the new test data to the collection
            test_label = room_properties.test_label
            self.test_data_collection[test_label] = {
                test_type: {
                    'room_properties': room_properties,
                    'test_data': None  # Will be populated when data files are loaded
                }
            }
            
            if self.debug_checkbox.active:
                print(f"Added new test: {test_label} ({test_type.value})")
                print(f"Room properties: {vars(room_properties)}")
            
            self.status_label.text = f'Status: Added new test {test_label}'
            
            # Update analysis dashboard if it exists
            if hasattr(self, 'analysis_dashboard'):
                self.analysis_dashboard.update_data(self.test_data_collection)
                
        except Exception as e:
            self._show_error(f"Error saving test plan: {str(e)}")

    def preview_pdf(self, pdf_path):
        """Preview generated PDF report"""
        try:
            if os.path.exists(pdf_path):
                # Create temporary image of first page
                doc = fitz.open(pdf_path)
                page = doc[0]
                pix = page.get_pixmap()
                img_path = tempfile.mktemp(suffix='.png')
                pix.save(img_path)
                
                # Create scrollable image preview
                scroll = ScrollView(size_hint=(1, 1))
                image = KivyImage(source=img_path, size_hint_y=None)
                image.height = image.texture_size[1]
                scroll.add_widget(image)
                
                # Create popup with preview
                content = BoxLayout(orientation='vertical')
                content.add_widget(scroll)
                
                preview_popup = Popup(
                    title='PDF Preview',
                    content=content,
                    size_hint=(0.9, 0.9)
                )
                preview_popup.open()
            else:
                self._show_error('PDF file not found')
                
        except Exception as e:
            self._show_error(f'Error previewing PDF: {str(e)}')

    def show_results(self, test_data: TestData):
        """Display test results in a popup"""
        content = GridLayout(cols=1, padding=10, spacing=5)
        
        # Add test information
        props = test_data.room_properties
        info_text = (
            f"Test Label: {props.test_label}\n"
            f"Project: {props.project_name}\n"
            f"Test Date: {props.test_date}\n"
            f"Source Room: {props.source_room}\n"
            f"Receive Room: {props.receive_room}"
        )
        content.add_widget(Label(text=info_text))
        
        # Add results table
        results_table = self._create_results_table(test_data)
        content.add_widget(results_table)
        
        popup = Popup(
            title='Test Results',
            content=content,
            size_hint=(0.8, 0.8)
        )
        popup.open()

    def generate_reports(self, test_data: TestData):
        """Generate PDF report for test data"""
        try:
            if not test_data:
                raise ValueError("No test data provided")
            
            test_collection = self.test_data_manager.get_test_collection()
            
            for test_label, test_types in test_collection.items():
                print(f"\nGenerating reports for test: {test_label}")
                
                for test_type, data in test_types.items():
                    try:
                        self.status_label.text = f'Status: Generating {test_type.value} report for {test_label}...'
                        
                        # Debug: Print data before conversion
                        print("\nData before conversion:")
                        if hasattr(data['test_data'], 'calculated_values'):
                            print("Calculated values keys:", data['test_data'].calculated_values.keys())
                        else:
                            print("No calculated_values attribute found")
                        
                        # Convert the data to the expected format
                        converted_test_data = self.convert_to_test_data(data)
                        
                        # Debug: Print data after conversion
                        print("\nData after conversion:")
                        if hasattr(converted_test_data, 'calculated_values'):
                            print("Calculated values keys:", converted_test_data.calculated_values.keys())
                        else:
                            print("No calculated_values attribute found after conversion")
                        
                        # Get appropriate report class
                        report_class = {
                            TestType.ASTC: ASTCTestReport,
                            TestType.AIIC: AIICTestReport,
                            TestType.NIC: NICTestReport,
                            TestType.DTC: DTCTestReport
                        }.get(test_type)
                        
                        if report_class:
                            # Generate report using the converted test data
                            report = report_class.create_report(
                                test_data=converted_test_data,
                                output_folder=self.output_path.text,
                                test_type=test_type
                            )
                            
                            # Get PDF path
                            pdf_path = os.path.join(
                                self.output_path.text,
                                f"{test_label}_{test_type.value}_Report.pdf"
                            )
                            
                            if self.debug_checkbox.active:
                                print(f'Generated {test_type.value} report for test {test_label}')
                            
                            # Show preview
                            self.preview_pdf(pdf_path)
                            
                    except Exception as e:
                        error_msg = f"Error generating {test_type.value} report for {test_label}: {str(e)}"
                        print(error_msg)
                        if self.debug_checkbox.active:
                            traceback.print_exc()
                        continue
            
            self.status_label.text = 'Status: All reports generated successfully'
            return True
            
        except Exception as e:
            error_msg = f"Error generating reports: {str(e)}"
            if self.debug_checkbox.active:
                print(f"\nERROR: {error_msg}")
                traceback.print_exc()
            self._show_error(error_msg)
            self.status_label.text = f"Status: {error_msg}"
            return False

    def show_plot_selection(self, instance):
        """Show popup with buttons to plot data for each test, with test type selection"""
        if self.debug_checkbox.active:
            print("\n=== Show Plot Selection ===")
            print("Test Data Collection Contents:")
            for label, data in self.test_data_manager.test_data_collection.items():
                print(f"\nTest: {label}")
                print(f"Available types: {list(data.keys())}")
                for test_type, test_data in data.items():
                    print(f"- {test_type.value}:")
                    print(f"  Room properties: {test_data['room_properties']}")
                    print(f"  Test data type: {type(test_data['test_data'])}")

        # Create content layout
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)
        
        # Add title label
        content.add_widget(Label(
            text='Select tests and types to plot:',
            size_hint_y=None,
            height=40
        ))
        
        # Create scrollable list
        scroll = ScrollView(size_hint=(1, 0.8))
        main_layout = GridLayout(
            cols=1,
            spacing=10,
            size_hint_y=None
        )
        main_layout.bind(minimum_height=main_layout.setter('height'))
        
        # Dictionary to store checkbox states for each test
        self.test_type_checkboxes = {}
        
        # Add section for each test
        for test_label in self.test_data_manager.test_data_collection:
            # Debug print
            print(f"\nProcessing test: {test_label}")
            test_data = self.test_data_manager.test_data_collection[test_label]
            print(f"Test data keys: {test_data.keys()}")
            
            # Create container for this test
            test_container = BoxLayout(
                orientation='vertical',
                size_hint_y=None,
                height=120,
                padding=[10, 5]
            )
            
            # Add test label
            test_container.add_widget(Label(
                text=f'Test {test_label}',
                size_hint_y=None,
                height=30,
                bold=True
            ))
            
            # Create horizontal layout for checkboxes
            checkbox_layout = GridLayout(
                cols=3,
                spacing=10,
                size_hint_y=None,
                height=40
            )
            
            # Store checkboxes for this test
            self.test_type_checkboxes[test_label] = {}
            
            # Check for available test types using the enum values
            available_types = []
            
            if TestType.AIIC in test_data:
                available_types.append(('AIIC', TestType.AIIC))
            if TestType.ASTC in test_data:
                available_types.append(('ASTC', TestType.ASTC))
            if TestType.NIC in test_data:
                available_types.append(('NIC', TestType.NIC))
            if TestType.DTC in test_data:
                available_types.append(('DTC', TestType.DTC))
                
            # Debug print available types
            print(f"Available test types for {test_label}: {[t[0] for t in available_types]}")
            
            for type_name, test_type in available_types:
                # Create checkbox container
                type_container = BoxLayout(
                    orientation='horizontal',
                    size_hint_y=None,
                    height=30,
                    spacing=5
                )
                
                # Create checkbox
                checkbox = CheckBox(
                    size_hint_x=None,
                    width=30,
                    active=False
                )
                self.test_type_checkboxes[test_label][test_type] = checkbox
                
                # Create label
                label = Label(
                    text=type_name,
                    size_hint_x=None,
                    width=70,
                    halign='left'
                )
                
                # Add widgets to container
                type_container.add_widget(checkbox)
                type_container.add_widget(label)
                
                # Add container to checkbox layout
                checkbox_layout.add_widget(type_container)
            
            test_container.add_widget(checkbox_layout)
            
            # Add plot button
            plot_btn = Button(
                text='Plot Selected',
                size_hint_y=None,
                height=30
            )
            plot_btn.bind(
                on_press=lambda x, label=test_label: self.plot_selected_test_data(label)
            )
            test_container.add_widget(plot_btn)
            
            # Add separator
            separator = Widget(
                size_hint_y=None,
                height=2
            )
            with separator.canvas:
                Color(0.5, 0.5, 0.5, 1)
                Rectangle(pos=separator.pos, size=separator.size)
            
            def update_rect(instance, value):
                instance.canvas.children[-1].pos = instance.pos
                instance.canvas.children[-1].size = instance.size
            
            separator.bind(pos=update_rect, size=update_rect)
            test_container.add_widget(separator)
            
            main_layout.add_widget(test_container)
        
        scroll.add_widget(main_layout)
        content.add_widget(scroll)
        
        # Create close button
        close_btn = Button(
            text='Close',
            size_hint_y=None,
            height=40
        )
        
        # Create popup
        plot_popup = Popup(
            title='Plot Test Data',
            content=content,
            size_hint=(0.8, 0.8)
        )
        
        # Bind close button
        close_btn.bind(on_press=plot_popup.dismiss)
        content.add_widget(close_btn)
        
        plot_popup.open()

    def plot_selected_test_data(self, test_label):
        """Plot selected test types for a specific test"""
        try:
            # Define standard frequencies
            standard_freqs = [63, 125, 250, 500, 1000, 2000, 4000]
            
            # Create figure with two subplots
            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10), height_ratios=[3, 2])
            plt.subplots_adjust(hspace=0.3)

            print("\n=== Plot Selected Test Data ===")
            print(f"Test Label: {test_label}")
            
            # Get selected test types
            selected_types = [
                test_type for test_type, checkbox in self.test_type_checkboxes[test_label].items()
                if checkbox.active
            ]
            print(f"Selected Types: {[t.value for t in selected_types]}")
            
            if not selected_types:
                warning = Popup(
                    title='Warning',
                    content=Label(text='Please select at least one test type to plot'),
                    size_hint=(0.6, 0.2)
                )
                warning.open()
                return
            
            # Get test data
            test_data = self.test_data_manager.test_data_collection[test_label]
            
            # Initialize frequency data dictionary for calculations
            self.freq_data = {
                'room_props': None,
                'source': None,
                'receive': None,
                'avg_pos': None,
                'positions': None,
                'background': None,
                'rt': None
            }
            
            # Plot data for each selected test type
            for test_type in selected_types:
                if test_type in test_data:
                    test_obj = test_data[test_type].get('test_data')
                    if test_obj:
                        print(f"\nPlotting {test_type.value} data:")
                        
                        # Store room properties for calculations
                        self.freq_data['room_props'] = test_obj.room_properties
                        
                        # Process and store data for calculations while plotting
                        if hasattr(test_obj, 'srs_data'):
                            print("- Processing source room data")
                            if hasattr(test_obj.srs_data, 'raw_data'):
                                # SLMData object with raw_data
                                df = test_obj.srs_data.raw_data
                                print("- Plotting source room SLMData.raw_data")
                            elif isinstance(test_obj.srs_data, pd.DataFrame):
                                # Direct DataFrame
                                df = test_obj.srs_data
                                print("- Plotting source room DataFrame directly")
                            else:
                                print(f"- Unsupported source data type: {type(test_obj.srs_data)}")
                                df = None
                                
                            if df is not None:
                                ax1.plot(
                                    df['Frequency (Hz)'],
                                    df['Overall 1/3 Spectra'],
                                    label=f'{test_type.value} - Source Room'
                                )
                                # Store for calculations - using static method call
                                formatted_df = TestDataManager.format_slm_data(df)
                                self.freq_data['source'] = formatted_df['Overall 1/3 Spectra'].values[0:17]
                        
                        # Receive room data
                        if hasattr(test_obj, 'recive_data'):
                            print("- Processing receive room data")
                            if hasattr(test_obj.recive_data, 'raw_data'):
                                # SLMData object with raw_data
                                df = test_obj.recive_data.raw_data
                                print("- Plotting receive room SLMData.raw_data")
                            elif isinstance(test_obj.recive_data, pd.DataFrame):
                                # Direct DataFrame
                                df = test_obj.recive_data
                                print("- Plotting receive room DataFrame directly")
                            else:
                                print(f"- Unsupported receive data type: {type(test_obj.recive_data)}")
                                df = None
                                
                            if df is not None:
                                ax1.plot(
                                    df['Frequency (Hz)'],
                                    df['Overall 1/3 Spectra'],
                                    label=f'{test_type.value} - Receive Room'
                                )
                                # Store for calculations
                                formatted_df = TestDataManager.format_slm_data(df)
                                self.freq_data['receive'] = formatted_df['Overall 1/3 Spectra'].values[0:17]
                        
                        # Background data
                        if hasattr(test_obj, 'bkgrnd_data'):
                            print("- Processing background data")
                            if hasattr(test_obj.bkgrnd_data, 'raw_data'):
                                # SLMData object with raw_data
                                df = test_obj.bkgrnd_data.raw_data
                                print("- Plotting background SLMData.raw_data")
                            elif isinstance(test_obj.bkgrnd_data, pd.DataFrame):
                                # Direct DataFrame
                                df = test_obj.bkgrnd_data
                                print("- Plotting background DataFrame directly")
                            else:
                                print(f"- Unsupported background data type: {type(test_obj.bkgrnd_data)}")
                                df = None
                                
                            if df is not None:
                                ax1.plot(
                                    df['Frequency (Hz)'],
                                    df['Overall 1/3 Spectra'],
                                    label=f'{test_type.value} - Background'
                                )
                                # Store for calculations
                                formatted_df = TestDataManager.format_slm_data(df)
                                self.freq_data['background'] = formatted_df['Overall 1/3 Spectra'].values[0:17]
                        
                        # Store RT data for calculations
                        if hasattr(test_obj, 'rt'):
                            try:
                                # First try accessing rt_thirty if it exists
                                if hasattr(test_obj.rt, 'rt_thirty'):
                                    rt_data = test_obj.rt.rt_thirty[:17]
                                    print("Using rt_thirty property")
                                # Then try pandas DataFrame access if it's a DataFrame
                                elif isinstance(test_obj.rt, pd.DataFrame):
                                    if 'Unnamed: 10' in test_obj.rt.columns:
                                        rt_data = test_obj.rt['Unnamed: 10'][24:41]/1000
                                        print("Using DataFrame column access")
                                    else:
                                        # Try to find numeric columns in case column name changed
                                        numeric_cols = test_obj.rt.select_dtypes(include=[np.number]).columns
                                        if len(numeric_cols) > 0:
                                            rt_data = test_obj.rt[numeric_cols[0]][24:41]/1000
                                            print(f"Using alternative column: {numeric_cols[0]}")
                                        else:
                                            print("No suitable numeric columns found in RT data")
                                            rt_data = None
                                # If it's already a numpy array or list
                                elif isinstance(test_obj.rt, (np.ndarray, list)):
                                    rt_data = test_obj.rt[:17] if len(test_obj.rt) >= 17 else test_obj.rt
                                    print("Using array/list access")
                                else:
                                    print(f"Unsupported RT data type: {type(test_obj.rt)}")
                                    rt_data = None
                                    
                                # Store the data if we successfully retrieved it
                                if rt_data is not None:
                                    self.freq_data['rt'] = np.array(rt_data, dtype=np.float64).round(3)
                                    print(f"Stored RT data with shape: {self.freq_data['rt'].shape}")
                            except Exception as e:
                                print(f"Error accessing RT data: {str(e)}")
                                traceback.print_exc()
                                # Set to None so calculations know to handle missing RT data
                                self.freq_data['rt'] = None

                        # Process test-specific data
                        if test_type == TestType.AIIC:
                            positions, avg_pos = self._process_aiic_positions(test_obj)
                            if positions is not None:
                                self.freq_data['positions'] = positions
                                self.freq_data['avg_pos'] = avg_pos
                            self._process_aiic_plot(ax1, ax2, test_obj)
                        elif test_type == TestType.ASTC:
                            self._process_astc_plot(ax1, ax2, test_obj)
                        elif test_type == TestType.NIC:
                            self._process_nic_plot(ax1, ax2, test_obj)

            # Configure plots after all data is added
            ax1.set_title(f'Test {test_label} - Raw Data')
            ax1.set_xlabel('Frequency (Hz)')
            ax1.set_ylabel('Sound Pressure Level (dB)')
            ax1.grid(True)
            ax1.set_xscale('log')
            ax1.set_xticks(standard_freqs)
            ax1.set_xticklabels([str(f) for f in standard_freqs])
            ax1.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize='small')

            ax2.set_title('Noise Reduction Analysis')
            ax2.set_xlabel('Frequency (Hz)')
            ax2.set_ylabel('Noise Reduction (dB)')
            ax2.grid(True)
            ax2.set_xscale('log')
            ax2.set_xticks(standard_freqs)
            ax2.set_xticklabels([str(f) for f in standard_freqs])
            ax2.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize='small')

            # Adjust layout to prevent legend cutoff
            plt.tight_layout()
            
            # Save plot to a temporary file
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                plt.savefig(temp_file.name, dpi=300, bbox_inches='tight')
                
                # Create scrollable image view
                scroll = ScrollView(size_hint=(1, 0.9))  # Reduced height to make room for button
                image = KivyImage(
                    source=temp_file.name,
                    size_hint=(1, None)
                )
                image.height = image.texture_size[1]
                scroll.add_widget(image)
                
                # Create button layout at bottom
                button_layout = BoxLayout(
                    size_hint_y=None, 
                    height='48dp',
                    spacing='10dp',
                    padding='10dp'
                )
                
                # Add store values button
                store_button = Button(
                    text='Store Calculated Values',
                    size_hint_x=None,
                    width='200dp'
                )
                store_button.bind(on_press=lambda x: self._store_calculated_values(test_label))
                
                button_layout.add_widget(store_button)
                
                # Add both to main layout
                layout = BoxLayout(orientation='vertical')
                layout.add_widget(scroll)
                layout.add_widget(button_layout)
                
                # Show in popup
                plot_popup = Popup(
                    title=f'Test Data Plot - {test_label}',
                    content=layout,
                    size_hint=(0.9, 0.9)
                )
                plot_popup.open()
            
        except Exception as e:
            print(f"Error plotting test data: {str(e)}")
            traceback.print_exc()

    def _store_calculated_values(self, test_label):
        """Store calculated values for all processed test types"""
        try:
            print(f"\n=== Storing Calculated Values for Test {test_label} ===")
            
            # Get test data
            test_data = self.test_data_manager.test_data_collection[test_label]
            print(f"Available test types: {list(test_data.keys())}")
            
            # Debug: Print test_data_manager state
            print("\nTest Data Manager State:")
            print(f"Type: {type(self.test_data_manager)}")
            print(f"Has test_data_collection: {hasattr(self.test_data_manager, 'test_data_collection')}")
            print(f"test_data_collection length: {len(self.test_data_manager.test_data_collection) if hasattr(self.test_data_manager, 'test_data_collection') else 'N/A'}")
            print(f"Has test_plan: {hasattr(self.test_data_manager, 'test_plan')}")
            print(f"test_plan is None: {self.test_data_manager.test_plan is None if hasattr(self.test_data_manager, 'test_plan') else 'N/A'}")
            
            # Debug: Print collection structure
            print("\nTest Collection Structure:")
            for test_type, data in test_data.items():
                print(f"\n{test_type}:")
                print(f"  Keys in data: {list(data.keys())}")
                if 'test_data' in data:
                    print(f"  Test data attributes: {[attr for attr in dir(data['test_data']) if not attr.startswith('_')]}")
                    if hasattr(data['test_data'], 'calculated_values'):
                        print(f"  Calculated values: {data['test_data'].calculated_values.keys()}")
            
            # Add debug: Print state before storage
            print("\nBefore storage:")
            for test_type, data in test_data.items():
                if hasattr(data['test_data'], 'calculated_values'):
                    print(f"{test_type}: {data['test_data'].calculated_values.keys()}")
                else:
                    print(f"{test_type}: No calculated values")
            
            # Ensure TestProcessor has access to the collection
            self.test_data_manager.test_processor.set_test_collection(
                self.test_data_manager.test_data_collection
            )
            
            for test_type in test_data.keys():
                print(f"\nProcessing {test_type} test:")
                
                test_obj = test_data[test_type]['test_data']
                
                if test_type == TestType.AIIC:
                    # First process the raw data like we do for plotting
                    raw_data = self._get_aiic_raw_data(test_obj)
                    if not raw_data:
                        continue
                        
                    # Process frequency data
                    freq_data = self._process_aiic_frequencies(raw_data)
                    if not freq_data:
                        continue
                    
                    # Now calculate values using properly processed data
                    calculated_values = self._calculate_aiic_values(test_obj, freq_data)
                    
                    # Add debug output to see what's being stored
                    if calculated_values:
                        print("\nAIIC Calculated Values being stored:")
                        for key, value in calculated_values.items():
                            if hasattr(value, 'shape'):
                                print(f"  {key}: {type(value)} with shape {value.shape}")
                                if key == 'sabines':
                                    print(f"    values: {value}")
                            else:
                                print(f"  {key}: {type(value)}")
                
                elif test_type == TestType.ASTC:
                    print("\nASTC Calculation Input:")
                    # Process raw data first
                    raw_data = self._get_astc_raw_data(test_obj)
                    if not raw_data:
                        continue
                    
                    # Process frequency data
                    freq_data = self._process_astc_frequencies(raw_data)
                    if not freq_data:
                        continue
                    
                    # Calculate values using processed data
                    calculated_values = self._calculate_astc_values(test_obj, freq_data)
                    print("\nASTC Calculated Values:")
                    if calculated_values:
                        print(f"- Keys: {calculated_values.keys()}")
                        for key, value in calculated_values.items():
                            if hasattr(value, 'shape'):
                                print(f"- {key} shape: {value.shape}")
                            elif isinstance(value, (list, np.ndarray)):
                                print(f"- {key} length: {len(value)}")
                
                elif test_type == TestType.NIC:
                    print("\nNIC Calculation Input:")
                    # Process raw data first
                    raw_data = self._get_nic_raw_data(test_obj)
                    if not raw_data:
                        continue
                    
                    # Process frequency data
                    freq_data = self._process_nic_frequencies(raw_data)
                    if not freq_data:
                        continue
                    
                    # Calculate values using processed data
                    calculated_values = self._calculate_nic_values(test_obj, freq_data)
                    print("\nNIC Calculated Values:")
                    if calculated_values:
                        print(f"- Keys: {calculated_values.keys()}")
                        for key, value in calculated_values.items():
                            if hasattr(value, 'shape'):
                                print(f"- {key} shape: {value.shape}")
                            elif isinstance(value, (list, np.ndarray)):
                                print(f"- {key} length: {len(value)}")
                
                if calculated_values:
                    print(f"\nStoring values for {test_type}:")
                    self.test_data_manager.test_processor.store_calculated_values(
                        test_label,
                        test_type,
                        calculated_values
                    )
            
            # Add debug: Print state after storage
            print("\nAfter storage:")
            for test_type, data in test_data.items():
                if hasattr(data['test_data'], 'calculated_values'):
                    print(f"{test_type}: {data['test_data'].calculated_values.keys()}")
                else:
                    print(f"{test_type}: No calculated values")
            
            # Show success message
            success_popup = Popup(
                title='Success',
                content=Label(text='Calculated values stored successfully'),
                size_hint=(None, None),
                size=(300, 150)
            )
            success_popup.open()
            
            # Debug: Print final storage state
            print("\nFinal Storage State:")
            for test_type, data in test_data.items():
                if hasattr(data['test_data'], 'calculated_values'):
                    print(f"\n{test_type}:")
                    print(f"  Stored values: {data['test_data'].calculated_values.keys()}")
                    for key, value in data['test_data'].calculated_values.items():
                        if hasattr(value, 'shape'):
                            print(f"    {key} shape: {value.shape}")
                        elif isinstance(value, (list, np.ndarray)):
                            print(f"    {key} length: {len(value)}")
                        else:
                            print(f"    {key} type: {type(value)}")
            
            # Refresh the results dashboard after storing values
            self.refresh_results_dashboard()
            
        except Exception as e:
            error_msg = f"Error storing calculated values: {str(e)}"
            print(f"\nError details:")
            print(f"- Error type: {type(e)}")
            print(f"- Error message: {str(e)}")
            traceback.print_exc()
            self._show_error(error_msg)

    def _process_single_position(self, raw_data, pos_name):
        """Process a single AIIC position's data"""
        try:
            print(f"Processing {pos_name} data")
            # Use TestDataManager's format_slm_data method
            formatted_df = TestDataManager.format_slm_data(raw_data)
            # Get the Overall 1/3 Spectra values for the first 17 frequencies
            return formatted_df['Overall 1/3 Spectra'].values[0:17]
        except Exception as e:
            print(f"Error processing {pos_name}: {str(e)}")
            return None

    def _process_aiic_positions(self, test_obj):
        """Process AIIC position data and return averaged results"""
        print("\nProcessing AIIC position data:")
        average_pos = []

        # Process each position
        for i in range(1, 5):
            pos_attr = f'pos{i}'
            print(f"\n- Processing {pos_attr}")
            if hasattr(test_obj, pos_attr):
                pos_data = getattr(test_obj, pos_attr)
                if hasattr(pos_data, 'raw_data'):
                    processed_data = self._process_single_position(pos_data.raw_data, pos_attr)
                    if processed_data is not None:
                        average_pos.append(processed_data)

        # Calculate average of positions
        if average_pos:
            onethird_rec_Total = np.mean(average_pos, axis=0)
            print(f"Successfully processed {len(average_pos)} positions")
            return average_pos, onethird_rec_Total
        return None, None

    def _find_overall_spectra_row(self, df, start_idx):
        """Find the Overall 1/3 Spectra row in the dataframe"""
        for j in range(start_idx + 2, min(start_idx + 7, len(df))):
            if 'Overall 1/3 Spectra' in str(df.iloc[j].iloc[0]):
                return df.iloc[j]
        print("  Warning: Could not find Overall 1/3 Spectra row")
        return None

    def _process_frequency_data(self, freq_row, data_row):
        """Process frequency and SPL data"""
        # Convert and clean frequency values
        freq_values = pd.to_numeric(freq_row.iloc[1:], errors='coerce')
        spl_values = pd.to_numeric(data_row.iloc[1:], errors='coerce')

        # Remove NaN values
        mask = ~(freq_values.isna() | spl_values.isna())
        freq_values = freq_values[mask]
        spl_values = spl_values[mask]

        # Filter for frequencies between 63 and 4000 Hz
        mask = (freq_values >= 63) & (freq_values <= 4000)
        freq_values = freq_values[mask]
        spl_values = spl_values[mask]

        if len(spl_values) >= 16:
            return spl_values
        print(f"  Warning: Insufficient frequency bands: {len(spl_values)}")
        return None

    def _process_aiic_plot(self, ax1, ax2, test_obj):
        """Create AIIC-specific plots for both raw data and analysis"""
        try:
            print("\n=== START: AIIC PROCESSING ===")
            
            # Get raw data
            print("\n--- Getting Raw Data ---")
            raw_data = self._get_aiic_raw_data(test_obj)
            if not raw_data:
                return False
            
            print("\nRaw Data Validation:")
            print(f"Source range: {np.min(raw_data['source'])} to {np.max(raw_data['source'])}")
            print(f"Background range: {np.min(raw_data['background'])} to {np.max(raw_data['background'])}")
            print(f"RT range: {np.min(raw_data['rt'])} to {np.max(raw_data['rt'])}")
            for i, pos in enumerate(raw_data['positions']):
                print(f"Position {i+1} range: {np.min(pos)} to {np.max(pos)}")
            
            # Process frequency data
            print("\n--- Processing Frequencies ---")
            freq_data = self._process_aiic_frequencies(raw_data)
            if not freq_data:
                return False
            
            print("\nFrequency Data Validation:")
            print(f"Average position range: {np.min(freq_data['avg_pos'])} to {np.max(freq_data['avg_pos'])}")
            print("Position by frequency:")
            for freq, avg in zip(freq_data['target_freqs'], freq_data['avg_pos']):
                print(f"{freq} Hz: {avg:.1f} dB")
            
            # Calculate AIIC
            print("\n--- Calculating AIIC Values ---")
            aiic_results = self._calculate_aiic_values(test_obj, freq_data)
            if not aiic_results:
                return False
            
            print("\nAIIC Results Validation:")
            if aiic_results['AIIC_Normalized_recieve'] is not None:
                print("Normalized Receive by frequency:")
                for freq, val in zip(freq_data['target_freqs'], aiic_results['AIIC_Normalized_recieve']):
                    print(f"{freq} Hz: {val:.1f} dB")
            
            # Create plots
            print("\n--- Creating Plots ---")
            self._create_aiic_analysis_plot(ax1, ax2, freq_data, aiic_results)
            print("\n=== END: AIIC PROCESSING ===")
            return True
            
        except Exception as e:
            print(f"Error in AIIC processing: {str(e)}")
            traceback.print_exc()
            return False

    def _get_aiic_raw_data(self, test_obj):
        """Extract raw data for AIIC calculations"""
        try:
            print("\n=== AIIC Raw Data Processing ===")
            
            # Debug test object structure and file paths
            print("\nTest Object Information:")
            print(f"Test Name: {getattr(test_obj, 'test_name', 'Unknown')}")
            print(f"Test Date: {getattr(test_obj, 'test_date', 'Unknown')}")
            print(f"Test Directory: {getattr(test_obj, 'test_dir', 'Unknown')}")
            
            # 1/3 octave bands from 100 to 3150 Hz
            freq_indices = slice(12, 28)
            
            # Initialize raw data structure
            raw_data = {
                'freq': None,
                'source': None,
                'background': None,
                'rt': None,
                'room_props': test_obj.room_properties,
                'positions': []
            }
            
            # Source room data
            print("\nSource Room Data:")
            if hasattr(test_obj, 'srs_data'):
                if hasattr(test_obj.srs_data, 'raw_data'):
                    # Using SLMData with raw_data
                    source_df = test_obj.srs_data.raw_data
                    raw_data['freq'] = source_df['Frequency (Hz)'].values
                    raw_data['source'] = source_df['Overall 1/3 Spectra'].values[freq_indices]
                    print("Using source room SLMData.raw_data")
                    if hasattr(test_obj.srs_data, 'file_path'):
                        print(f"Source File: {os.path.basename(test_obj.srs_data.file_path)}")
                elif isinstance(test_obj.srs_data, pd.DataFrame):
                    # Direct DataFrame
                    source_df = test_obj.srs_data
                    raw_data['freq'] = source_df['Frequency (Hz)'].values
                    raw_data['source'] = source_df['Overall 1/3 Spectra'].values[freq_indices]
                    print("Using source room DataFrame directly")
                else:
                    print(f"Unsupported source data type: {type(test_obj.srs_data)}")
            
            # Background data
            print("\nBackground Data:")
            if hasattr(test_obj, 'bkgrnd_data'):
                if hasattr(test_obj.bkgrnd_data, 'raw_data'):
                    # Using SLMData with raw_data
                    raw_data['background'] = test_obj.bkgrnd_data.raw_data['Overall 1/3 Spectra'].values[freq_indices]
                    print("Using background SLMData.raw_data")
                    if hasattr(test_obj.bkgrnd_data, 'file_path'):
                        print(f"Background File: {os.path.basename(test_obj.bkgrnd_data.file_path)}")
                elif isinstance(test_obj.bkgrnd_data, pd.DataFrame):
                    # Direct DataFrame
                    raw_data['background'] = test_obj.bkgrnd_data['Overall 1/3 Spectra'].values[freq_indices]
                    print("Using background DataFrame directly")
                else:
                    print(f"Unsupported background data type: {type(test_obj.bkgrnd_data)}")
            
            # RT data
            print("\nRT Data:")
            if hasattr(test_obj, 'rt'):
                try:
                    # First try accessing rt_thirty if it exists
                    if hasattr(test_obj.rt, 'rt_thirty'):
                        raw_data['rt'] = test_obj.rt.rt_thirty[:-1]  # Remove last value
                        print("Using rt_thirty property")
                        if hasattr(test_obj.rt, 'file_path'):
                            print(f"RT File: {os.path.basename(test_obj.rt.file_path)}")
                    # Then try pandas DataFrame access if it's a DataFrame
                    elif isinstance(test_obj.rt, pd.DataFrame):
                        if 'Unnamed: 10' in test_obj.rt.columns:
                            raw_data['rt'] = test_obj.rt['Unnamed: 10'][24:40]/1000  # Slice to match rt_thirty[:-1]
                            print("Using DataFrame column access for RT")
                        else:
                            # Try to find numeric columns in case column name changed
                            numeric_cols = test_obj.rt.select_dtypes(include=[np.number]).columns
                            if len(numeric_cols) > 0:
                                raw_data['rt'] = test_obj.rt[numeric_cols[0]][24:40]/1000
                                print(f"Using alternative column for RT: {numeric_cols[0]}")
                            else:
                                print("No suitable numeric columns found in RT data")
                    # If it's already a numpy array or list
                    elif isinstance(test_obj.rt, (np.ndarray, list)):
                        raw_data['rt'] = test_obj.rt[:-1] if len(test_obj.rt) > 16 else test_obj.rt
                        print("Using array/list access for RT")
                    else:
                        print(f"Unsupported RT data type: {type(test_obj.rt)}")
                except Exception as e:
                    print(f"Error accessing RT data: {str(e)}")
                    traceback.print_exc()
            
            # Process tapping positions
            print("\nTapping Position Files:")
            positions = {
                1: 'pos1',
                2: 'pos2',
                3: 'pos3',
                4: 'pos4'
            }
            
            for pos_num, pos_attr in positions.items():
                if hasattr(test_obj, pos_attr):
                    pos = getattr(test_obj, pos_attr)
                    print(f"\nPosition {pos_num}:")
                    
                    # Check if it's an SLMData object with raw_data
                    if hasattr(pos, 'raw_data'):
                        pos_df = pos.raw_data
                        if hasattr(pos, 'file_path'):
                            print(f"Position File: {os.path.basename(pos.file_path)}")
                        print("Using position SLMData.raw_data")
                    # Check if it's directly a DataFrame
                    elif isinstance(pos, pd.DataFrame):
                        pos_df = pos
                        print("Using position DataFrame directly")
                    else:
                        print(f"Unsupported position data type: {type(pos)}")
                        continue
                        
                    try:
                        # Process position data
                        pos_data = pos_df.loc[pos_df['1/1 Octave'] == 'Overall 1/3 Spectra']
                        if not pos_data.empty:
                            pos_values = pos_data.iloc[0, 1:].values[freq_indices]
                            pos_values = np.array(pos_values, dtype=np.float64).round(1)
                            if len(pos_values) == 16:
                                raw_data['positions'].append(pos_values)
                                print(f"Successfully processed")
                                print(f"First few values: {pos_values[:5]}...")
                            else:
                                print(f"Warning: Wrong number of values: {len(pos_values)}")
                        else:
                            print("Warning: No '1/3 Spectra' row found")
                            print("Available rows in file:")
                            print(pos_df['1/1 Octave'].unique())
                    except Exception as e:
                        print(f"Error processing position: {str(e)}")
                else:
                    print(f"\nPosition {pos_num}: Not provided")
            
            # Room properties
            print("\nRoom Properties:")
            if hasattr(test_obj.room_properties, 'file_path'):
                print(f"Properties File: {os.path.basename(test_obj.room_properties.file_path)}")
            print(f"Receive Volume: {test_obj.room_properties.receive_vol}")
            
            # Verify all required data is present
            missing = [k for k, v in raw_data.items() if v is None and k != 'positions']
            if missing:
                print(f"Missing AIIC data: {missing}")
                return None
                
            # Final data validation
            print("\nFinal Data Summary:")
            if raw_data['freq'] is not None:
                print(f"Number of frequencies: {len(raw_data['freq'])}")
            print(f"Source data points: {len(raw_data['source'])}")
            print(f"Background data points: {len(raw_data['background'])}")
            print(f"RT data points: {len(raw_data['rt'])}")
            print(f"Number of valid positions: {len(raw_data['positions'])}")
            
            if not raw_data['positions']:
                raise ValueError("No valid position data could be processed")
                
            return raw_data
            
        except Exception as e:
            print(f"\nError getting AIIC raw data: {str(e)}")
            traceback.print_exc()
            return None

    def _process_aiic_frequencies(self, raw_data):
        """Process frequency data for AIIC calculations"""
        try:
            # Calculate average position using onethird Logavg
            onethird_rec_Total = calculate_onethird_Logavg(raw_data['positions'])
            print(f"onethird_rec_Total: {onethird_rec_Total}")
            processed_data = {
                'target_freqs': [100, 125, 160, 200, 250, 315, 400, 500, 
                               630, 800, 1000, 1250, 1600, 2000, 2500, 3150],
                'source': raw_data['source'],
                'background': raw_data['background'],
                'rt': raw_data['rt'],
                'room_props': raw_data['room_props'],
                'positions': raw_data['positions'],
                'avg_pos': onethird_rec_Total
            }
            
            print("\nProcessed data shapes:")
            print(f"Target frequencies: {len(processed_data['target_freqs'])} points")
            print(f"Source: {processed_data['source'].shape}")
            print(f"Background: {processed_data['background'].shape}")
            print(f"RT: {processed_data['rt'].shape}")
            print(f"RT: {processed_data['rt']}")
            print(f"Average position: {processed_data['avg_pos'].shape}")
            print(f"Number of processed positions: {len(processed_data['positions'])}")
            
            return processed_data
            
        except Exception as e:
            print(f"Error processing AIIC frequencies: {str(e)}")
            traceback.print_exc()
            return None

    def _calculate_aiic_values(self, test_obj, freq_data):
        """Calculate AIIC values"""
        try:
            print("\nCalculating AIIC values:")
            print("-=-=-=-=-=-=-=-=-=-raw data:-=-=-=--=-=-=-=-=-=")
            print(f"Source: {freq_data['source']}")
            print(f"Avg Pos: {freq_data['avg_pos']}")
            print(f"Background: {freq_data['background']}")
            print(f"Room Props: {freq_data['room_props']}")
            print(f"RT: {freq_data['rt']}")
            print("-=-=-=-=-=-=-=-=-=-raw data:-=-=-=--=-=-=-=-=-=")
            # Calculate NR using the averaged position data
            NR_val, _, sabines, AIIC_recieve_corr, ASTC_recieve_corr, AIIC_Normalized_recieve = calc_NR_new(
                freq_data['source'],
                freq_data['avg_pos'],
                None,  # No receive data for AIIC
                freq_data['background'],
                float(freq_data['room_props'].receive_vol),
                freq_data['rt']
            )
            
            print(f"AIIC_recieve_corr: {AIIC_recieve_corr}")
            print(f"AIIC_Normalized_recieve: {AIIC_Normalized_recieve}")
            print(f"NR_val: {NR_val}")
            print(f"Calculated sabines value: {sabines}")  # Debug print

            if AIIC_Normalized_recieve is None:
                print("Error: AIIC Normalized receive calculation failed")
                return None

            AIIC_contour_val, AIIC_contour_result = calc_AIIC_val_claude(AIIC_Normalized_recieve)
            print(f"AIIC_contour_val: {AIIC_contour_val}")
            print(f"AIIC_contour_result: {AIIC_contour_result}")
            
            # Process exceptions
            AIIC_Exceptions = []
            AIIC_exceptions_backcheck = []
            rec_roomvol = float(freq_data['room_props'].receive_vol)

            # Calculate exceptions using stored values
            for val in sabines:
                val_float = float(val)
                AIIC_Exceptions.append('1' if val_float > 2 * rec_roomvol**(2/3) else '0')

            # Background check exceptions
            print(f"AIIC_Normalized_recieve: {AIIC_Normalized_recieve}")
            print(f"onethird_bkgrd: {freq_data['background']}")
            # Calculate background difference using correct column (overall level)
            background_diff = AIIC_Normalized_recieve - freq_data['background']
            
            # Create exceptions list based on difference
            AIIC_exceptions_backcheck = ['0' if diff > 5 else '1' for diff in background_diff]
            print(f"AIIC_exceptions_backcheck: {AIIC_exceptions_backcheck}")
            #########################################################

            # Create return dictionary
            calculated_values = {
                'NR_val': NR_val,
                'sabines': sabines.copy() if hasattr(sabines, 'copy') else sabines,  # Ensure we store a copy
                'AIIC_recieve_corr': AIIC_recieve_corr,
                'AIIC_Normalized_recieve': AIIC_Normalized_recieve,
                'AIIC_Exceptions': AIIC_Exceptions,
                'AIIC_exceptions_backcheck': AIIC_exceptions_backcheck,
                'positions': freq_data['positions'],
                'AIIC_contour_val': AIIC_contour_val,
                'AIIC_contour_result': AIIC_contour_result,
                'room_vol': float(freq_data['room_props'].receive_vol)
            }
            
            # Verify the dictionary contains sabines
            print("\nVerifying calculated values before return:")
            for key, value in calculated_values.items():
                print(f"  {key}: {type(value)}")
                if hasattr(value, 'shape'):
                    print(f"    shape: {value.shape}")
                    if key == 'sabines':
                        print(f"    values: {value}")
            
            return calculated_values
            
        except Exception as e:
            print(f"Error calculating AIIC values: {str(e)}")
            traceback.print_exc()
            return None

    def _create_aiic_analysis_plot(self, ax1, ax2, freq_data, results):
        """Create the AIIC analysis plots"""
        try:
            # Plot position data on ax1 (raw data plot)
            for i, pos_data in enumerate(results['positions'], 1):
                ax1.plot(
                    freq_data['target_freqs'],
                    pos_data,
                    linestyle='--',
                    label=f'AIIC - Tapping Position {i}'
                )
            IIC_curve = [2,2,2,2,2,2,1,0,-1,-2,-3,-6,-9,-12,-15,-18]
            IIC_contour_final = list()
            # initial application of the IIC curve to the first AIIC start value 
            for vals in IIC_curve:
                IIC_contour_final.append(vals+(110-results['AIIC_contour_val']))
            #########################################################


            # Plot normalized receive levels on ax2
            ax2.plot(freq_data['target_freqs'], results['AIIC_Normalized_recieve'], 
                    label='Normalized Impact Sound Level', 
                    color='blue', marker='o')
            
            # Plot AIIC contour
            ax2.plot(freq_data['target_freqs'], IIC_contour_final, 
                    label=f'AIIC {results["AIIC_contour_val"]} Contour', 
                    color='red', linestyle='--')
            # Plot background check exceptions
            for i, (freq, level, exception) in enumerate(zip(freq_data['target_freqs'], 
                                                           results['AIIC_Normalized_recieve'],
                                                           results['AIIC_exceptions_backcheck'])):
                if exception == '1':
                    ax2.axvline(x=freq, color='orange', alpha=0.6, linestyle='--', label='Background Check Exception')
            
            # Update axis labels for AIIC
            ax2.set_title('AIIC Analysis')
            ax2.set_xlabel('Frequency (Hz)')
            ax2.set_ylabel('Sound Pressure Level (dB)')
            
            # Add results text box
            textstr = f'AIIC: {results["AIIC_contour_val"]}'
            props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
            ax2.text(0.05, 0.95, textstr, transform=ax2.transAxes, 
                    verticalalignment='top', bbox=props)
            
            print("Successfully created AIIC analysis plots")
            return True
            
        except Exception as e:
            print(f"Error creating AIIC analysis plots: {str(e)}")
            traceback.print_exc()
            return False

    def _process_astc_plot(self, ax1, ax2, test_obj):
        """Create ASTC-specific plots for both raw data and analysis"""
        try:
            print("\nProcessing ASTC data:")
            
            # Get raw data
            raw_data = self._get_astc_raw_data(test_obj)
            if not raw_data:
                return False
            
            # Process frequency data
            freq_data = self._process_astc_frequencies(raw_data)
            if not freq_data:
                return False
            
            # Calculate ASTC
            astc_results = self._calculate_astc_values(test_obj, freq_data)
            if not astc_results:
                return False
            
            # Create plot
            self._create_astc_analysis_plot(ax2, freq_data['target_freqs'], astc_results)
            print("Successfully processed ASTC data")
            return True
            
        except Exception as e:
            print(f"Error in ASTC processing: {str(e)}")
            traceback.print_exc()
            return False

    def _calculate_astc_values(self, test_obj, freq_data):
        """Calculate ASTC values"""
        try:
            print("\nCalculating ASTC values:")
            print("Input data shapes:")
            print(f"Source data: {freq_data['source'].shape if hasattr(freq_data['source'], 'shape') else len(freq_data['source'])}")
            print(f"Receive data: {freq_data['receive'].shape if hasattr(freq_data['receive'], 'shape') else len(freq_data['receive'])}")
            print(f"Background data: {freq_data['background'].shape if hasattr(freq_data['background'], 'shape') else len(freq_data['background'])}")
            print(f"RT data: {freq_data['rt'].shape if hasattr(freq_data['rt'], 'shape') else len(freq_data['rt'])}")
            
            # Ensure all data has correct length before calculation
            expected_length = 17  # We expect 17 frequency points
            if any(len(data) != expected_length for data in [freq_data['source'], freq_data['receive'], 
                                                        freq_data['background'], freq_data['rt']]):
                print("\nData length mismatch detected:")
                print(f"Source length: {len(freq_data['source'])}")
                print(f"Receive length: {len(freq_data['receive'])}")
                print(f"Background length: {len(freq_data['background'])}")
                print(f"RT length: {len(freq_data['rt'])}")
                raise ValueError(f"All frequency data must have length {expected_length}")
            NR_val, _, sabines, _, ASTC_recieve_corr, _ = calc_NR_new(
                freq_data['source'],
                None,
                freq_data['receive'],  # No receive data for AIIC
                freq_data['background'],
                float(freq_data['room_props'].receive_vol),
                freq_data['rt']
            )
            ATL_val, sabines = calc_atl_val(
                freq_data['source'],
                freq_data['receive'],
                freq_data['background'],
                freq_data['rt'],
                float(freq_data['room_props'].partition_area),
                float(freq_data['room_props'].receive_vol)
            )
            
            if ATL_val is None:
                print("Error: ATL calculation failed")
                return None
            
            print(f"ATL values shape: {ATL_val.shape if hasattr(ATL_val, 'shape') else len(ATL_val)}")
            print(f"ATL values: {ATL_val}")
            
            # Calculate ASTC and contour
            ASTC_final_val = calc_astc_val(ATL_val)
            STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4, 4]
            ASTC_contour_val = [val + ASTC_final_val for val in STCCurve]
            
            print(f"ASTC final value: {ASTC_final_val}")
            
            return {
                'ATL_val': ATL_val,
                'NR_val': NR_val,
                'ASTC_final_val': ASTC_final_val,
                'ASTC_contour_val': ASTC_contour_val,
                'ASTC_recieve_corr': ASTC_recieve_corr,
                'sabines': sabines,
                'room_vol': float(freq_data['room_props'].receive_vol)
            }
            
        except Exception as e:
            print(f"Error calculating ASTC values: {str(e)}")
            traceback.print_exc()
            return None

    def _create_astc_analysis_plot(self, ax, freq_values, results):
        """Create the ASTC analysis plot with all components"""
        try:
            # Plot ATL values
            ax.plot(freq_values, results['ATL_val'], 
                    label='Apparent Transmission Loss (ATL)', 
                    color='blue', marker='o')
            
            # Plot ASTC contour
            ax.plot(freq_values, results['ASTC_contour_val'], 
                    label=f'ASTC {results["ASTC_final_val"]} Contour', 
                    color='red', linestyle='--')
            
            # Update axis labels for ASTC
            ax.set_title('ASTC Analysis')
            ax.set_xlabel('Frequency (Hz)')
            ax.set_ylabel('Sound Level (dB)')
            
            # Add results text box
            textstr = (f'ASTC: {results["ASTC_final_val"]}\n'
                      f'Room Volume: {results["room_vol"]:.1f} m³')
            props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
            ax.text(0.05, 0.95, textstr, transform=ax.transAxes, 
                    verticalalignment='top', bbox=props)
            
            print("Successfully created ASTC analysis plot")
            return True
            
        except Exception as e:
            print(f"Error creating ASTC analysis plot: {str(e)}")
            traceback.print_exc()
            return False

    def _process_nic_plot(self, ax1, ax2, test_obj):
        """Create NIC-specific plots for both raw data and analysis"""
        try:
            print("\nProcessing NIC data:")
            
            # Get raw data
            raw_data = self._get_nic_raw_data(test_obj)
            if not raw_data:
                return False
            
            # Process frequency data
            freq_data = self._process_nic_frequencies(raw_data)
            if not freq_data:
                return False
            
            # Calculate NIC
            nic_results = self._calculate_nic_values(test_obj, freq_data)
            if not nic_results:
                return False
            
            # Create plot
            self._create_nic_analysis_plot(ax2, freq_data['target_freqs'], nic_results)
            print("Successfully processed NIC data")
            return True
            
        except Exception as e:
            print(f"Error in NIC processing: {str(e)}")
            traceback.print_exc()
            return False

    def _get_astc_raw_data(self, test_obj):
        """Get raw data for ASTC calculations"""
        try:
            print("Getting ASTC raw data")
            raw_data = {
                'source': None,
                'receive': None,
                'background': None,
                'rt': None,
                'room_props': test_obj.room_properties
            }
            
            # Handle source room data
            if hasattr(test_obj, 'srs_data'):
                if hasattr(test_obj.srs_data, 'raw_data'):
                    # SLMData object with raw_data
                    raw_data['source'] = test_obj.srs_data.raw_data
                    print("Using source room SLMData.raw_data")
                elif isinstance(test_obj.srs_data, pd.DataFrame):
                    # Direct DataFrame
                    raw_data['source'] = test_obj.srs_data
                    print("Using source room DataFrame directly")
                else:
                    print(f"Unsupported source data type: {type(test_obj.srs_data)}")
            
            # Handle receive room data
            if hasattr(test_obj, 'recive_data'):
                if hasattr(test_obj.recive_data, 'raw_data'):
                    # SLMData object with raw_data
                    raw_data['receive'] = test_obj.recive_data.raw_data
                    print("Using receive room SLMData.raw_data")
                elif isinstance(test_obj.recive_data, pd.DataFrame):
                    # Direct DataFrame
                    raw_data['receive'] = test_obj.recive_data
                    print("Using receive room DataFrame directly")
                else:
                    print(f"Unsupported receive data type: {type(test_obj.recive_data)}")
            
            # Handle background data
            if hasattr(test_obj, 'bkgrnd_data'):
                if hasattr(test_obj.bkgrnd_data, 'raw_data'):
                    # SLMData object with raw_data
                    raw_data['background'] = test_obj.bkgrnd_data.raw_data
                    print("Using background SLMData.raw_data")
                elif isinstance(test_obj.bkgrnd_data, pd.DataFrame):
                    # Direct DataFrame
                    raw_data['background'] = test_obj.bkgrnd_data
                    print("Using background DataFrame directly")
                else:
                    print(f"Unsupported background data type: {type(test_obj.bkgrnd_data)}")
            
            # Handle RT data
            if hasattr(test_obj, 'rt'):
                try:
                    # First try accessing rt_thirty if it exists
                    if hasattr(test_obj.rt, 'rt_thirty'):
                        raw_data['rt'] = test_obj.rt.rt_thirty[:17]
                        print("Using rt_thirty property")
                    # Then try pandas DataFrame access if it's a DataFrame
                    elif isinstance(test_obj.rt, pd.DataFrame):
                        if 'Unnamed: 10' in test_obj.rt.columns:
                            raw_data['rt'] = test_obj.rt['Unnamed: 10'][24:41]/1000
                            print("Using DataFrame column access for RT")
                        else:
                            # Try to find numeric columns in case column name changed
                            numeric_cols = test_obj.rt.select_dtypes(include=[np.number]).columns
                            if len(numeric_cols) > 0:
                                raw_data['rt'] = test_obj.rt[numeric_cols[0]][24:41]/1000
                                print(f"Using alternative column for RT: {numeric_cols[0]}")
                            else:
                                print("No suitable numeric columns found in RT data")
                    # If it's already a numpy array or list
                    elif isinstance(test_obj.rt, (np.ndarray, list)):
                        raw_data['rt'] = test_obj.rt[:17] if len(test_obj.rt) >= 17 else test_obj.rt
                        print("Using array/list access for RT")
                    else:
                        print(f"Unsupported RT data type: {type(test_obj.rt)}")
                        
                    # Convert to numpy array if needed
                    if raw_data['rt'] is not None:
                        raw_data['rt'] = np.array(raw_data['rt'], dtype=np.float64).round(3)
                except Exception as e:
                    print(f"Error accessing RT data: {str(e)}")
                    traceback.print_exc()
            
            # Debug the raw data we collected
            for key, value in raw_data.items():
                if key != 'room_props':
                    if value is not None:
                        print(f"{key} data type: {type(value)}")
                        print(f"{key} data shape/length: {value.shape if hasattr(value, 'shape') else len(value) if hasattr(value, '__len__') else 'unknown'}")
                    else:
                        print(f"{key} data: None")
            
            # Verify all required data is present
            missing = [k for k, v in raw_data.items() if v is None]
            if missing:
                print(f"Missing ASTC data: {missing}")
                return None
            
            return raw_data
            
        except Exception as e:
            print(f"Error getting ASTC raw data: {str(e)}")
            traceback.print_exc()
            return None

    def _process_astc_frequencies(self, raw_data):
        """Process frequency data for ASTC calculations"""
        try:
            print("\nProcessing ASTC frequency data:")
            
            # Define target frequencies (100-4000 Hz for ASTC)
            target_freqs = [100, 125, 160, 200, 250, 315, 400, 500, 
                           630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000]
            
            print("\nProcessing raw data:")
            print(f"Source data type: {type(raw_data['source'])}")
            print(f"Source data shape: {raw_data['source'].shape}")
            print(f"Source data columns: {raw_data['source'].columns.tolist()}")
            
            # Extract frequency data directly from raw DataFrame
            # Find indices for our target frequencies
            freq_indices = []
            for freq in target_freqs:
                idx = raw_data['source'][raw_data['source']['Frequency (Hz)'] == freq].index
                if len(idx) > 0:
                    freq_indices.append(idx[0])
            
            print(f"\nFound indices for target frequencies: {freq_indices}")
            
            # Extract data for target frequencies
            source_data = raw_data['source'].iloc[freq_indices]['Overall 1/3 Spectra'].values
            receive_data = raw_data['receive'].iloc[freq_indices]['Overall 1/3 Spectra'].values
            background_data = raw_data['background'].iloc[freq_indices]['Overall 1/3 Spectra'].values
            
            print("\nExtracted frequency data:")
            print(f"Source data shape: {source_data.shape}")
            print(f"Source values: {source_data}")
            print(f"Receive data shape: {receive_data.shape}")
            print(f"Background data shape: {background_data.shape}")
            
            # Create frequency data dictionary
            freq_data = {
                'target_freqs': target_freqs,
                'source': source_data,
                'receive': receive_data,
                'background': background_data,
                'rt': raw_data['rt'],
                'room_props': raw_data['room_props']
            }
            
            print("\nProcessed frequency data:")
            print(f"Target frequencies: {len(freq_data['target_freqs'])} points")
            for key in ['source', 'receive', 'background', 'rt']:
                print(f"{key} data: {len(freq_data[key])} points")
                print(f"{key} values: {freq_data[key]}")
            
            # Verify all arrays have the expected length
            expected_length = 17
            for key in ['source', 'receive', 'background', 'rt']:
                if len(freq_data[key]) != expected_length:
                    raise ValueError(f"{key} data length mismatch. Expected {expected_length}, got {len(freq_data[key])}")
            
            return freq_data
            
        except Exception as e:
            print(f"Error processing ASTC frequency data: {str(e)}")
            traceback.print_exc()
            return None

    def _get_nic_raw_data(self, test_obj):
        """Get raw data for NIC calculations"""
        try:
            print("Getting NIC raw data")
            raw_data = {
                'source': None,
                'receive': None,
                'background': None,
                'rt': None,
                'room_props': test_obj.room_properties
            }
            
            # Handle source room data
            if hasattr(test_obj, 'srs_data'):
                if hasattr(test_obj.srs_data, 'raw_data'):
                    # SLMData object with raw_data
                    raw_data['source'] = test_obj.srs_data.raw_data
                    print("Using source room SLMData.raw_data")
                elif isinstance(test_obj.srs_data, pd.DataFrame):
                    # Direct DataFrame
                    raw_data['source'] = test_obj.srs_data
                    print("Using source room DataFrame directly")
                else:
                    print(f"Unsupported source data type: {type(test_obj.srs_data)}")
            
            # Handle receive room data
            if hasattr(test_obj, 'recive_data'):
                if hasattr(test_obj.recive_data, 'raw_data'):
                    # SLMData object with raw_data
                    raw_data['receive'] = test_obj.recive_data.raw_data
                    print("Using receive room SLMData.raw_data")
                elif isinstance(test_obj.recive_data, pd.DataFrame):
                    # Direct DataFrame
                    raw_data['receive'] = test_obj.recive_data
                    print("Using receive room DataFrame directly")
                else:
                    print(f"Unsupported receive data type: {type(test_obj.recive_data)}")
            
            # Handle background data
            if hasattr(test_obj, 'bkgrnd_data'):
                if hasattr(test_obj.bkgrnd_data, 'raw_data'):
                    # SLMData object with raw_data
                    raw_data['background'] = test_obj.bkgrnd_data.raw_data
                    print("Using background SLMData.raw_data")
                elif isinstance(test_obj.bkgrnd_data, pd.DataFrame):
                    # Direct DataFrame
                    raw_data['background'] = test_obj.bkgrnd_data
                    print("Using background DataFrame directly")
                else:
                    print(f"Unsupported background data type: {type(test_obj.bkgrnd_data)}")
            
            # Handle RT data
            if hasattr(test_obj, 'rt'):
                try:
                    # First try accessing rt_thirty if it exists
                    if hasattr(test_obj.rt, 'rt_thirty'):
                        raw_data['rt'] = test_obj.rt.rt_thirty[:17]
                        print("Using rt_thirty property")
                    # Then try pandas DataFrame access if it's a DataFrame
                    elif isinstance(test_obj.rt, pd.DataFrame):
                        if 'Unnamed: 10' in test_obj.rt.columns:
                            raw_data['rt'] = test_obj.rt['Unnamed: 10'][24:41]/1000
                            print("Using DataFrame column access for RT")
                        else:
                            # Try to find numeric columns in case column name changed
                            numeric_cols = test_obj.rt.select_dtypes(include=[np.number]).columns
                            if len(numeric_cols) > 0:
                                raw_data['rt'] = test_obj.rt[numeric_cols[0]][24:41]/1000
                                print(f"Using alternative column for RT: {numeric_cols[0]}")
                            else:
                                print("No suitable numeric columns found in RT data")
                    # If it's already a numpy array or list
                    elif isinstance(test_obj.rt, (np.ndarray, list)):
                        raw_data['rt'] = test_obj.rt[:17] if len(test_obj.rt) >= 17 else test_obj.rt
                        print("Using array/list access for RT")
                    else:
                        print(f"Unsupported RT data type: {type(test_obj.rt)}")
                        
                    # Convert to numpy array if needed
                    if raw_data['rt'] is not None:
                        raw_data['rt'] = np.array(raw_data['rt'], dtype=np.float64).round(3)
                except Exception as e:
                    print(f"Error accessing RT data: {str(e)}")
                    traceback.print_exc()
            
            # Debug the raw data we collected
            for key, value in raw_data.items():
                if key != 'room_props':
                    if value is not None:
                        print(f"{key} data type: {type(value)}")
                        print(f"{key} data shape/length: {value.shape if hasattr(value, 'shape') else len(value) if hasattr(value, '__len__') else 'unknown'}")
                    else:
                        print(f"{key} data: None")
            
            # Verify all required data is present
            missing = [k for k, v in raw_data.items() if v is None]
            if missing:
                print(f"Missing NIC data: {missing}")
                return None
            
            return raw_data
            
        except Exception as e:
            print(f"Error getting NIC raw data: {str(e)}")
            traceback.print_exc()
            return None

    def _process_nic_frequencies(self, raw_data):
        """Process frequency data for NIC calculations"""
        try:
            print("\nProcessing NIC frequency data:")
            
            # Define target frequencies (100-4000 Hz for NIC)
            target_freqs = [100, 125, 160, 200, 250, 315, 400, 500, 
                           630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000]
            
            print("\nProcessing raw data:")
            print(f"Source data type: {type(raw_data['source'])}")
            print(f"Source data shape: {raw_data['source'].shape}")
            print(f"Source data columns: {raw_data['source'].columns.tolist()}")
            
            # Extract frequency data directly from raw DataFrame
            # Find indices for our target frequencies
            freq_indices = []
            for freq in target_freqs:
                idx = raw_data['source'][raw_data['source']['Frequency (Hz)'] == freq].index
                if len(idx) > 0:
                    freq_indices.append(idx[0])
            
            print(f"\nFound indices for target frequencies: {freq_indices}")
            
            # Extract data for target frequencies
            source_data = raw_data['source'].iloc[freq_indices]['Overall 1/3 Spectra'].values
            receive_data = raw_data['receive'].iloc[freq_indices]['Overall 1/3 Spectra'].values
            background_data = raw_data['background'].iloc[freq_indices]['Overall 1/3 Spectra'].values
            
            print("\nExtracted frequency data:")
            print(f"Source data shape: {source_data.shape}")
            print(f"Source values: {source_data}")
            print(f"Receive data shape: {receive_data.shape}")
            print(f"Background data shape: {background_data.shape}")
            
            # Create frequency data dictionary
            freq_data = {
                'target_freqs': target_freqs,
                'source': source_data,
                'receive': receive_data,
                'background': background_data,
                'rt': raw_data['rt'],
                'room_props': raw_data['room_props']
            }
            
            print("\nProcessed frequency data:")
            print(f"Target frequencies: {len(freq_data['target_freqs'])} points")
            for key in ['source', 'receive', 'background', 'rt']:
                print(f"{key} data: {len(freq_data[key])} points")
                print(f"{key} values: {freq_data[key]}")
            
            # Verify all arrays have the expected length
            expected_length = 17
            for key in ['source', 'receive', 'background', 'rt']:
                if len(freq_data[key]) != expected_length:
                    raise ValueError(f"{key} data length mismatch. Expected {expected_length}, got {len(freq_data[key])}")
            
            return freq_data
            
        except Exception as e:
            print(f"Error processing NIC frequency data: {str(e)}")
            traceback.print_exc()
            return None

    def _calculate_nic_values(self, test_obj, freq_data):
        """Calculate NIC values"""
        try:
            print("\nCalculating NIC values:")
            print("Input data validation:")
            for key in ['source', 'receive', 'background', 'rt']:
                print(f"{key} data length: {len(freq_data[key])}")
                print(f"{key} values: {freq_data[key]}")
            
            room_vol = test_obj.room_properties.receive_vol
            print(f"Room volume: {room_vol}")
            
            # Calculate NIC using the processed data
            NR_val, NIC_contour_val, sabines, _, NIC_recieve_corr, _ = calc_NR_new(
                freq_data['source'],
                None,  # No AIIC data
                freq_data['receive'],
                freq_data['background'],
                room_vol,
                freq_data['rt']
            )
            NIC_final_val = calc_astc_val(NR_val)
            print(f"NIC final value: {NIC_final_val}")
            if NR_val is not None:
                print(f"NR values shape: {NR_val.shape if hasattr(NR_val, 'shape') else len(NR_val)}")
                print(f"NIC contour value: {NIC_contour_val}")
                print(f"NIC final value: {NIC_final_val}")
                print(f"Sabines: {sabines}")
                
                return {
                    'NR_val': NR_val,
                    'NIC_contour_val': NIC_contour_val,
                    'NIC_final_val': NIC_final_val,
                    'NIC_recieve_corr': NIC_recieve_corr,
                    'sabines': sabines,
                    'room_vol': room_vol
                }
            
            return None
            
        except Exception as e:
            print(f"Error calculating NIC values: {str(e)}")
            traceback.print_exc()
            return None

    def _create_nic_analysis_plot(self, ax, freq_values, nic_results):
        """Create the NIC analysis plot with all components"""
        try:
            # Plot NR values
            ax.plot(freq_values, nic_results['NR_val'], 
                    label=f"Noise Reduction (NIC = {nic_results['NIC_contour_val']})", 
                    color='blue', marker='o')
            
            # Plot NIC reference curve
            ref_curve = np.array([-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4, 4]) + nic_results['NIC_contour_val']
            ax.plot(freq_values, ref_curve, 
                    label='NIC Reference Curve', 
                    color='red', linestyle='--')
            
            # Add results text box
            textstr = (f"NIC: {nic_results['NIC_contour_val']}\n"
                      f"Sabines: {np.mean(nic_results['sabines']):.1f}\n"
                      f"Room Volume: {nic_results['room_vol']:.1f} m³")
            props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
            ax.text(0.05, 0.95, textstr, transform=ax.transAxes, 
                    verticalalignment='top', bbox=props)
                
            print("Successfully created NIC analysis plot")
            return True
            
        except Exception as e:
            print(f"Error creating NIC analysis plot: {str(e)}")
            traceback.print_exc()
            return False

    def convert_to_test_data(self, test_collection_data):
        """Convert test collection data to TestData format"""
        try:
            # Debug: Print incoming data
            print("\nConverting test data:")
            if 'test_data' in test_collection_data:
                test_data = test_collection_data['test_data']
                if hasattr(test_data, 'calculated_values'):
                    print("Input calculated values:", test_data.calculated_values.keys())
                else:
                    print("No calculated values in input data")
            
            # Extract the test data from the collection format
            if not test_collection_data or 'test_data' not in test_collection_data:
                raise ValueError("Invalid test collection data format")
            
            test_data = test_collection_data['test_data']
            
            # Convert SLMData objects to pandas DataFrames if needed
            if hasattr(test_data, 'srs_data'):
                test_data.srs_data = test_data.srs_data.raw_data if hasattr(test_data.srs_data, 'raw_data') else test_data.srs_data
            
            if hasattr(test_data, 'recive_data'):
                test_data.recive_data = test_data.recive_data.raw_data if hasattr(test_data.recive_data, 'raw_data') else test_data.recive_data
            
            if hasattr(test_data, 'bkgrnd_data'):
                test_data.bkgrnd_data = test_data.bkgrnd_data.raw_data if hasattr(test_data.bkgrnd_data, 'raw_data') else test_data.bkgrnd_data
            
            # Special handling for RT data
            if hasattr(test_data, 'rt'):
                if hasattr(test_data.rt, 'raw_data'):
                    # If rt_thirty is already processed and stored
                    if hasattr(test_data.rt, 'rt_thirty'):
                        test_data.rt = test_data.rt.rt_thirty
                    else:
                        # Extract RT30 data from the raw data
                        # Assuming the RT30 data is in rows 24:41 of the 'Unnamed: 10' column
                        try:
                            rt_data = test_data.rt.raw_data
                            if 'Unnamed: 10' in rt_data.columns:
                                test_data.rt = rt_data['Unnamed: 10'][24:41].values / 1000
                            else:
                                # Try to find the correct column containing RT30 data
                                # This might need adjustment based on your data structure
                                numeric_columns = rt_data.select_dtypes(include=[np.number]).columns
                                if len(numeric_columns) > 0:
                                    test_data.rt = rt_data[numeric_columns[0]][24:41].values / 1000
                                else:
                                    raise ValueError("Could not find RT30 data in the raw data")
                        except Exception as e:
                            print(f"Error processing RT data: {str(e)}")
                            raise

            # For AIIC tests, handle position data
            if hasattr(test_data, 'pos1'):
                test_data.pos1 = test_data.pos1.raw_data if hasattr(test_data.pos1, 'raw_data') else test_data.pos1
            if hasattr(test_data, 'pos2'):
                test_data.pos2 = test_data.pos2.raw_data if hasattr(test_data.pos2, 'raw_data') else test_data.pos2
            if hasattr(test_data, 'pos3'):
                test_data.pos3 = test_data.pos3.raw_data if hasattr(test_data.pos3, 'raw_data') else test_data.pos3
            if hasattr(test_data, 'pos4'):
                test_data.pos4 = test_data.pos4.raw_data if hasattr(test_data.pos4, 'raw_data') else test_data.pos4

            print(f"Test data conversion complete. RT data type: {type(test_data.rt)}")
            if hasattr(test_data, 'rt'):
                print(f"RT data shape/length: {test_data.rt.shape if hasattr(test_data.rt, 'shape') else len(test_data.rt)}")
            return test_data

        except Exception as e:
            print(f"Error converting test data: {str(e)}")
            raise

    def _add_data_section(self, title, data_dict):
        """Helper method to add a section of data to the raw data view"""
        # Add section title
        self.raw_data_grid.add_widget(Label(
            text=title,
            bold=True,
            size_hint_y=None,
            height=40
        ))
        
        # Add data rows
        for key, value in data_dict.items():
            row = BoxLayout(
                orientation='horizontal',
                size_hint_y=None,
                height=30
            )
            row.add_widget(Label(
                text=str(key),
                size_hint_x=0.4
            ))
            row.add_widget(Label(
                text=str(value),
                size_hint_x=0.6
            ))
            self.raw_data_grid.add_widget(row)

    def _update_test_plan_display(self):
        """Update the test plan grid with current test plan data"""
        # Clear existing widgets
        self.test_plan_grid.clear_widgets()
        
        test_plan = self.test_data_manager.test_plan
        
        try:
            # Set a fixed height for the test plan grid (adjust this value as needed)
            self.test_plan_grid.height = 400  # This will show approximately 12-13 rows at once
            
            # Create a horizontal ScrollView container
            h_scroll = ScrollView(
                size_hint=(1, None),
                height=400,  # Match the grid height
                do_scroll_y=False,
                do_scroll_x=True,
                bar_width=10,
                scroll_type=['bars', 'content']
            )
            
            # Create main container for all content
            content_grid = GridLayout(
                cols=1,
                spacing=2,
                size_hint=(None, None),
                padding=5
            )
            content_grid.bind(minimum_height=content_grid.setter('height'))
            
            # Set content width based on number of columns
            column_width = 300
            content_width = column_width * len(test_plan.columns)
            content_grid.width = content_width
            
            # Create header row
            header = GridLayout(
                cols=len(test_plan.columns),
                size_hint=(None, None),
                height=40,
                width=content_width,
                spacing=2
            )
            
            # Add column headers
            for col in test_plan.columns:
                header_label = Label(
                    text=str(col),
                    bold=True,
                    size_hint=(None, None),
                    width=column_width,
                    height=60,
                    text_size=(column_width-10, 60),
                    halign='center',
                    valign='middle'
                )
                header.add_widget(header_label)
            content_grid.add_widget(header)
            
            # Create scrollable data container
            data_grid = GridLayout(
                cols=len(test_plan.columns),
                size_hint=(None, None),
                width=content_width,
                spacing=2
            )
            data_grid.bind(minimum_height=data_grid.setter('height'))
            
            # Add data rows
            for _, row in test_plan.iterrows():
                for value in row:
                    cell_label = Label(
                        text=str(value),
                        size_hint=(None, None),
                        width=column_width,
                        height=60,
                        text_size=(column_width-10, 60),
                        halign='center',
                        valign='middle'
                    )
                    data_grid.add_widget(cell_label)
            
            # Create vertical scroll view for data
            v_scroll = ScrollView(
                size_hint=(1, None),
                height=360,  # Grid height minus header height
                do_scroll_x=True,
                bar_width=10,
                scroll_type=['bars', 'content']
            )
            
            v_scroll.add_widget(data_grid)
            content_grid.add_widget(v_scroll)
            
            # Add the content grid to the horizontal scroll view
            h_scroll.add_widget(content_grid)
            
            # Add the horizontal scroll view to the main grid
            self.test_plan_grid.add_widget(h_scroll)
            
        except Exception as e:
            print(f"Error updating test plan display: {str(e)}")
            traceback.print_exc()
            self.test_plan_grid.add_widget(Label(
                text=f'Error displaying test plan: {str(e)}',
                size_hint_y=None,
                height=40
            ))

    def update_displays(self):
        """Central method to update all displays"""
        try:
            # Update test plan display
            self.test_plan_grid.clear_widgets()
            if hasattr(self.test_data_manager, 'test_plan') and self.test_data_manager.test_plan is not None:
                self._update_test_plan_display()
            else:
                self.test_plan_grid.add_widget(Label(
                    text='No test plan data loaded',
                    size_hint_y=None,
                    height=40
                ))
            
            # Update results dashboard
            if hasattr(self, 'results_dashboard'):
                self.results_dashboard.update_data(self.test_data_manager)
            
            # Clear raw data view
            self.raw_data_grid.clear_widgets()
            
        except Exception as e:
            print(f"Error in update_displays: {str(e)}")
            self.status_label.text = f'Status: Error updating displays - {str(e)}'

    def refresh_results_dashboard(self):
        """Refresh the results dashboard to show updated data"""
        try:
            if hasattr(self, 'results_dashboard'):
                print("Refreshing results dashboard...")
                
                # First update the dashboard's reference to the test data
                if hasattr(self, 'test_data_manager') and hasattr(self.test_data_manager, 'test_data_collection'):
                    print(f"Passing test_data_collection with {len(self.test_data_manager.test_data_collection)} items to dashboard")
                    
                    # Explicitly update the dashboard's reference to the manager
                    self.results_dashboard.test_data_manager = self.test_data_manager
                    
                    # Also provide direct access to the collection via a dedicated property
                    if not hasattr(self.results_dashboard, 'direct_test_collection'):
                        # Add the property if it doesn't exist
                        setattr(self.results_dashboard, 'direct_test_collection', {})
                    
                    # Update the direct collection reference
                    self.results_dashboard.direct_test_collection = self.test_data_manager.test_data_collection
                    print(f"Direct collection reference set with {len(self.test_data_manager.test_data_collection)} items")
                
                # Then refresh the display
                if hasattr(self.results_dashboard, 'refresh_results'):
                    self.results_dashboard.refresh_results()
                
                # Also select the results tab to show updated data
                if hasattr(self, 'results_tab') and hasattr(self, 'analysis_tabs'):
                    self.analysis_tabs.switch_to(self.results_tab)
                    print("Switched to Results Dashboard tab")
        except Exception as e:
            print(f"Error refreshing results dashboard: {str(e)}")
            traceback.print_exc()
            
    def _fix_excel_extension_issue(self, file_path):
        """Attempt to fix the Excel file extension issue by creating a CSV copy"""
        try:
            # Get current file name and path components
            file_dir = os.path.dirname(file_path)
            file_name = os.path.basename(file_path)
            file_base, file_ext = os.path.splitext(file_name)
            
            # Create new file path with CSV extension
            new_file_path = os.path.join(file_dir, f"{file_base}.csv")
            
            # Copy the file to the new path
            import shutil
            shutil.copy2(file_path, new_file_path)
            
            # Update the file path in the UI
            self.test_plan_path.text = new_file_path
            
            # Show success message
            self.status_label.text = f"Status: Created CSV copy at {new_file_path}"
            
            return new_file_path
        except Exception as e:
            self._show_error(f"Failed to fix file extension: {str(e)}")
            return None

    def _show_excel_extension_error(self, message):
        """Show error popup with option to fix Excel extension issue"""
        content_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        # Add scrollable text for error message
        scroll_view = ScrollView(size_hint=(1, 0.7))
        error_label = Label(
            text=message + "\n\nThis file appears to be a CSV file with an incorrect .xlsx extension.",
            size_hint_y=None,
            text_size=(380, None)
        )
        error_label.bind(texture_size=error_label.setter('size'))
        scroll_view.add_widget(error_label)
        content_layout.add_widget(scroll_view)
        
        # Add buttons layout
        button_layout = BoxLayout(orientation='horizontal', size_hint=(1, 0.3), spacing=10)
        
        # Add fix button
        fix_button = Button(text="Create CSV Copy", size_hint=(0.5, 1))
        
        # Add close button
        close_button = Button(text="Close", size_hint=(0.5, 1))
        
        button_layout.add_widget(fix_button)
        button_layout.add_widget(close_button)
        content_layout.add_widget(button_layout)
        
        popup = Popup(
            title='Excel File Extension Error',
            content=content_layout,
            size_hint=(None, None),
            size=(500, 350)
        )
        
        # Handle button clicks
        fix_button.bind(on_press=lambda x: self._handle_fix_extension(popup))
        close_button.bind(on_press=popup.dismiss)
        
        popup.open()

    def _handle_fix_extension(self, popup):
        """Handle the fix extension button click"""
        popup.dismiss()
        
        # Fix the extension issue
        new_file_path = self._fix_excel_extension_issue(self.test_plan_path.text)
        
        if new_file_path:
            # Try loading the data again
            self.load_data(None)

    def show_file_picker(self, target_input, filters):
        """Show a file picker dialog and update the target input with selected path
        
        Args:
            target_input: TextInput widget to update with selected path
            filters: List of tuples with file filters, e.g. [('Excel Files', '*.xlsx')]
        """
        content = BoxLayout(orientation='vertical')
        
        # Get the workspace directory path
        import os
        workspace_dir = os.path.abspath(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
        
        # Create file chooser with proper filters
        def custom_filter(folder, filename):
            if os.path.isdir(os.path.join(folder, filename)):
                return True
            return any(filename.endswith(ext[1:]) for _, ext in filters)
            
        file_chooser = FileChooserListView(
            path=workspace_dir,  # Start in workspace directory
            filters=[custom_filter],
            size_hint=(1, 0.9)
        )
        
        # Create buttons
        button_layout = BoxLayout(
            size_hint_y=None,
            height='40dp',
            spacing='5dp',
            padding='5dp'
        )
        
        # Create select and cancel buttons
        select_button = Button(
            text='Select',
            size_hint_x=0.5
        )
        cancel_button = Button(
            text='Cancel',
            size_hint_x=0.5
        )
        
        # Add buttons to layout
        button_layout.add_widget(select_button)
        button_layout.add_widget(cancel_button)
        
        # Add widgets to content
        content.add_widget(file_chooser)
        content.add_widget(button_layout)
        
        # Create popup
        popup = Popup(
            title='Choose File',
            content=content,
            size_hint=(0.9, 0.9)
        )
        
        # Bind button actions
        def on_select(instance):
            if file_chooser.selection:
                # Convert to relative path if within workspace
                selected_path = file_chooser.selection[0]
                try:
                    relative_path = os.path.relpath(selected_path, workspace_dir)
                    target_input.text = relative_path
                except ValueError:
                    # If not within workspace, use absolute path
                    target_input.text = selected_path
                popup.dismiss()
                
        def on_cancel(instance):
            popup.dismiss()
            
        select_button.bind(on_press=on_select)
        cancel_button.bind(on_press=on_cancel)
        
        # Show popup
        popup.open()

class MainApp(App):
    def build(self):
        return MainWindow()

if __name__ == '__main__':
    MainApp().run() 