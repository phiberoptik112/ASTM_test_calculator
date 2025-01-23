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

import logging
logging.getLogger('matplotlib.font_manager').disabled = True
logging.getLogger('matplotlib').setLevel(logging.WARNING)
from data_processor import (
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
from base_test_reporter import *

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
        
        self.test_plan_path = TextInput(
            multiline=False,
            hint_text='Path to test plan Excel file',
            size_hint_x=0.7
        )
        test_plan_layout.add_widget(self.test_plan_path)
        
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
            hint_text='Enter test number'
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
        
        self.report_button = Button(text='Generate Report')
        self.report_button.bind(on_press=self.generate_report)
        button_layout.add_widget(self.report_button)
        
        # Add Plot Data button
        self.plot_button = Button(text='Raw Loaded Data Plot')
        self.plot_button.bind(on_press=self.show_plot_selection)
        button_layout.add_widget(self.plot_button)
        
        test_controls.add_widget(button_layout)
        self.add_widget(test_controls)

    def _create_analysis_section(self):
        """Create analysis dashboard section"""
        self.analysis_tabs = TabbedPanel(size_hint_y=0.5)
        
        # Results Tab
        results_tab = TabbedPanelItem(text='Results')
        self.results_dashboard = ResultsAnalysisDashboard(self.test_data_manager)
        results_tab.add_widget(self.results_dashboard)
        self.analysis_tabs.add_widget(results_tab)
        
        # Raw Data Tab
        raw_data_tab = TabbedPanelItem(text='Raw Data')
        self.raw_data_view = ScrollView()
        raw_data_tab.add_widget(self.raw_data_view)
        self.analysis_tabs.add_widget(raw_data_tab)
        
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
        """Display error message in popup"""
        popup = Popup(
            title='Error',
            content=Label(text=message),
            size_hint=(None, None),
            size=(400, 200)
        )
        popup.open()

    def populate_test_inputs(self, instance):
        """Populate input fields with example data paths"""
        self.test_plan_path.text = "./Exampledata/TestPlan_ASTM_testingv2.xlsx"
        self.slm_data_1_path.text = "./Exampledata/RawData/A_Meter/"
        self.slm_data_2_path.text = "./Exampledata/RawData/E_Meter/"
        self.output_path.text = "./Exampledata/testeroutputs/"
        
        self.status_label.text = "Status: Test inputs populated with example data"

    def load_data(self, instance):
        """Load and process test data files"""
        try:
            # Update status
            self.status_label.text = 'Status: Loading Data...'
            
            # Get debug mode from checkbox
            debug_mode = self.debug_checkbox.active
            
            # Initialize TestDataManager with debug mode
            print(f"Debug mode: {debug_mode}")
            print(f"TestDataManager: {self.test_data_manager}")
            self.test_data_manager = TestDataManager(debug_mode=debug_mode)
            
            # Print debug info before path processing
            if debug_mode:
                print("\n=== Loading Data ===")
                print(f"Raw Test Plan Path: {self.test_plan_path.text}")
                print(f"Raw Meter 1 Path: {self.slm_data_1_path.text}")
                print(f"Raw Meter 2 Path: {self.slm_data_2_path.text}")
                print(f"Raw Output Path: {self.output_path.text}")
            
            # Convert relative paths to absolute and normalize them
            try:
                base_dir = os.path.dirname(os.path.abspath(__file__))  # Get the directory of main_window.py
                paths = {
                    'test_plan': os.path.abspath(os.path.join(base_dir, '..', '..', self.test_plan_path.text)),
                    'meter_1': os.path.abspath(os.path.join(base_dir, '..', '..', self.slm_data_1_path.text)),
                    'meter_2': os.path.abspath(os.path.join(base_dir, '..', '..', self.slm_data_2_path.text)),
                    'output': os.path.abspath(os.path.join(base_dir, '..', '..', self.output_path.text))
                }
                
                if debug_mode:
                    print("\n=== Resolved Paths ===")
                    for key, path in paths.items():
                        print(f"{key}: {path}")
                        print(f"Exists: {os.path.exists(path)}")
            
            except Exception as e:
                raise ValueError(f"Error resolving paths: {str(e)}")
            
            # Set paths in TestDataManager
            try:
                self.test_data_manager.set_data_paths(
                    test_plan_path=paths['test_plan'],
                    meter_d_path=paths['meter_1'],
                    meter_e_path=paths['meter_2'],
                    output_path=paths['output']
                )
            except Exception as e:
                raise ValueError(f"Error setting paths in TestDataManager: {str(e)}")
            
            # Load test plan
            try:
                if debug_mode:
                    print("\nLoading test plan...")
                self.test_data_manager.load_test_plan()
            except Exception as e:
                raise ValueError(f"Error loading test plan: {str(e)}")
            
            # Process test data
            try:
                if debug_mode:
                    print("\nProcessing test data...")
                self.test_data_manager.process_test_data()
                self.status_label.text = 'Status: All test data loaded successfully'
                
                # Show plot selection popup after successful load
                self.show_plot_selection(instance)
                
                return True
                
            except Exception as e:
                raise ValueError(f"Error processing test data: {str(e)}")
                
        except Exception as e:
            error_msg = f"Error loading data: {str(e)}"
            if debug_mode:
                print(f"\nERROR: {error_msg}")
                print("Stack trace:")
                import traceback
                traceback.print_exc()
            self._show_error(error_msg)
            self.status_label.text = f"Status: {error_msg}"
            return False

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
        """View raw data for current test"""
        try:
            test_number = self.test_number_input.text
            if not test_number:
                raise ValueError("Please enter a test number")
                
            # Implementation similar to ASTC_GUI_proto's view_test_data
            # but integrated with TestDataManager
            
        except Exception as e:
            self.status_label.text = f'Status: Error viewing test data - {str(e)}'
            if self.debug_checkbox.active:
                print(f"Error viewing test data: {str(e)}")

    def show_test_plan_input(self, instance):
        """Show test plan input window"""
        content = TestPlanInputWindow(callback_on_save=self.on_test_plan_save)
        self.test_plan_popup = Popup(
            title='Add New Test',
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

    def generate_report(self, test_data: TestData):
        """Generate PDF report for test data"""
        try:
            if not hasattr(self, 'test_data_manager') or not self.test_data_manager:
                raise ValueError("No test data loaded. Please load data first.")

            # Get test collection from manager
            test_collection = self.test_data_manager.get_test_collection()
            
            for test_label, test_data_dict in test_collection.items():
                if self.debug_checkbox.active:
                    print(f'Generating reports for test: {test_label}')
                
                for test_type, data in test_data_dict.items():
                    self.status_label.text = f'Status: Generating {test_type.value} report for {test_label}...'
                    
                    try:
                        # Get appropriate report class
                        report_class = {
                            TestType.ASTC: ASTCTestReport,
                            TestType.AIIC: AIICTestReport,
                            TestType.NIC: NICTestReport,
                            TestType.DTC: DTCTestReport
                        }.get(test_type)
                        
                        if report_class:
                            # Generate report
                            report = report_class.create_report(
                                test_data=data.test_data,
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
            
            # Plot data for each selected test type
            for test_type in selected_types:
                if test_type in test_data:
                    test_obj = test_data[test_type].get('test_data')
                    if test_obj:
                        print(f"\nPlotting {test_type.value} data:")
                        
                        # Plot base data (source, receive, background) for all test types
                        if hasattr(test_obj, 'srs_data') and hasattr(test_obj.srs_data, 'raw_data'):
                            print("- Plotting source room data")
                            df = test_obj.srs_data.raw_data
                            ax1.plot(
                                df['Frequency (Hz)'],
                                df['Overall 1/3 Spectra'],
                                label=f'{test_type.value} - Source Room'
                            )
                        
                        if hasattr(test_obj, 'recive_data') and hasattr(test_obj.recive_data, 'raw_data'):
                            print("- Plotting receive room data")
                            df = test_obj.recive_data.raw_data
                            ax1.plot(
                                df['Frequency (Hz)'],
                                df['Overall 1/3 Spectra'],
                                label=f'{test_type.value} - Receive Room'
                            )
                        
                        if hasattr(test_obj, 'bkgrnd_data') and hasattr(test_obj.bkgrnd_data, 'raw_data'):
                            print("- Plotting background data")
                            df = test_obj.bkgrnd_data.raw_data
                            ax1.plot(
                                df['Frequency (Hz)'],
                                df['Overall 1/3 Spectra'],
                                label=f'{test_type.value} - Background'
                            )

                        # Process test-specific data
                        if test_type == TestType.AIIC:
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
            
            # Show plot in popup
            self.show_plot_popup()
            
        except Exception as e:
            print(f"Error plotting selected test data: {str(e)}")
            traceback.print_exc()

    def show_plot_popup(self):
        """Show the plot in a popup window"""
        # Save plot to a temporary file
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
            plt.savefig(temp_file.name, dpi=300, bbox_inches='tight')
            
            # Create scrollable image view
            scroll = ScrollView(size_hint=(1, 1))
            image = KivyImage(
                source=temp_file.name,
                size_hint=(1, None)
            )
            image.height = image.texture_size[1]
            scroll.add_widget(image)
            
            # Create and show popup
            plot_popup = Popup(
                title='Test Data Plot',
                content=scroll,
                size_hint=(0.9, 0.9)
            )
            plot_popup.open()

    def _process_aiic_positions(self, test_obj):
        """Process AIIC position data and return averaged results"""
        print("\nProcessing AIIC position data:")
        average_pos = []

        # Debug raw data structure
        for attr in ['pos1', 'pos2', 'pos3', 'pos4', 'source', 'carpet']:
            if hasattr(test_obj, attr):
                pos_data = getattr(test_obj, attr)
                if hasattr(pos_data, 'raw_data'):
                    print(f"\nRaw data structure for {attr}:")
                    print(f"DataFrame shape: {pos_data.raw_data.shape}")
                    print(f"Columns: {pos_data.raw_data.columns.tolist()}")

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

    def _process_single_position(self, pos_data):
        """Process a single AIIC position's data"""
        try:
            print("\n=== Processing Single Position ===")
            print("Input Data Structure:")
            print(f"DataFrame shape: {pos_data.shape}")
            print(f"Columns: {pos_data.columns.tolist()}")
            print("\nFirst few rows of raw data:")
            print(pos_data.head())
            
            # Get frequency values from first row
            freq_values = pd.to_numeric(pos_data.iloc[0, 1:], errors='coerce')
            print("\nFrequency Values (first row):")
            print(f"Raw values: {freq_values.values}")
            print(f"Number of frequencies: {len(freq_values)}")
            print(f"Range: {freq_values.min()} to {freq_values.max()} Hz")
            
            # Get SPL values from second row
            spl_values = pd.to_numeric(pos_data.iloc[1, 1:], errors='coerce')
            print("\nSPL Values (second row):")
            print(f"Raw values: {spl_values.values}")
            print(f"Number of SPL values: {len(spl_values)}")
            print(f"Range: {spl_values.min()} to {spl_values.max()} dB")
            
            # Remove any NaN values
            mask = ~(freq_values.isna() | spl_values.isna())
            print("\nNaN Removal:")
            print(f"Number of NaN values in frequencies: {freq_values.isna().sum()}")
            print(f"Number of NaN values in SPL: {spl_values.isna().sum()}")
            print(f"Valid data points after NaN removal: {mask.sum()}")
            
            freq_values = freq_values[mask]
            spl_values = spl_values[mask]
            
            # Filter for frequencies between 100 and 3150 Hz
            freq_mask = (freq_values >= 100) & (freq_values <= 3150)
            print("\nFrequency Filtering (100-3150 Hz):")
            print(f"Number of values in range: {freq_mask.sum()}")
            print("Frequency-SPL pairs in range:")
            for f, s in zip(freq_values[freq_mask], spl_values[freq_mask]):
                print(f"{f:.1f} Hz: {s:.1f} dB")
            
            spl_values = spl_values[freq_mask]
            
            if len(spl_values) >= 16:  # Should have 16 points (100-3150 Hz)
                print("\nFinal SPL Values:")
                print(f"Number of values: {len(spl_values)}")
                print(f"Values: {spl_values.values}")
                return spl_values
                
            print(f"\nWarning: Insufficient frequency bands: {len(spl_values)}")
            print(f"Required: 16, Found: {len(spl_values)}")
            return None
            
        except Exception as e:
            print(f"\nError processing position data: {str(e)}")
            print("Detailed error information:")
            traceback.print_exc()
            return None

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
            # still a hardcoded value... is there a better way to do this?
            freq_indices = slice(12, 28)
            
            # Source room data with detailed file info
            print("\nSource Room Data:")
            if hasattr(test_obj.srs_data, 'file_path'):
                print(f"Source File: {os.path.basename(test_obj.srs_data.file_path)}")
                print(f"Full Path: {test_obj.srs_data.file_path}")
            print(f"Data Shape: {test_obj.srs_data.raw_data.shape}")
            
            # Background data with detailed file info
            print("\nBackground Data:")
            if hasattr(test_obj.bkgrnd_data, 'file_path'):
                print(f"Background File: {os.path.basename(test_obj.bkgrnd_data.file_path)}")
                print(f"Full Path: {test_obj.bkgrnd_data.file_path}")
            print(f"Data Shape: {test_obj.bkgrnd_data.raw_data.shape}")
            
            # RT data with detailed file info
            # still a hardcoded value... is there a better way to do this?
            print("\nRT Data:")
            if hasattr(test_obj.rt, 'file_path'):
                print(f"RT File: {os.path.basename(test_obj.rt.file_path)}")
                print(f"Full Path: {test_obj.rt.file_path}")
            print(f"RT30 Values Shape: {test_obj.rt.rt_thirty.shape}")
            
            raw_data = {
                'freq': test_obj.srs_data.raw_data['Frequency (Hz)'].values,
                'source': test_obj.srs_data.raw_data['Overall 1/3 Spectra'].values[freq_indices],
                'background': test_obj.bkgrnd_data.raw_data['Overall 1/3 Spectra'].values[freq_indices],
                'rt': test_obj.rt.rt_thirty[:-1], # this also is still hardcoded 
                'room_props': test_obj.room_properties,
                'positions': []
            }
            
            # Process tapping positions with detailed file info
            positions = {
                1: test_obj.pos1,
                2: test_obj.pos2,
                3: test_obj.pos3,
                4: test_obj.pos4
            }
            
            print("\nTapping Position Files:")
            for pos_num, pos in positions.items():
                if pos is not None:
                    print(f"\nPosition {pos_num}:")
                    if hasattr(pos, 'file_path'):
                        print(f"Position File: {os.path.basename(pos.file_path)}")
                        print(f"Full Path: {pos.file_path}")
                    print(f"Data Shape: {pos.raw_data.shape}")
                    try:
                        # Process position data
                        pos_data = pos.raw_data.loc[pos.raw_data['1/1 Octave'] == 'Overall 1/3 Spectra']
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
                            print(pos.raw_data['1/1 Octave'].unique())
                    except Exception as e:
                        print(f"Error processing position: {str(e)}")
                else:
                    print(f"\nPosition {pos_num}: Not provided")
            
            # Room properties with detailed file info
            print("\nRoom Properties:")
            if hasattr(test_obj.room_properties, 'file_path'):
                print(f"Properties File: {os.path.basename(test_obj.room_properties.file_path)}")
                print(f"Full Path: {test_obj.room_properties.file_path}")
            print(f"Receive Volume: {test_obj.room_properties.receive_vol}")
            
            # Final data validation
            print("\nFinal Data Summary:")
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

            if AIIC_Normalized_recieve is None:
                print("Error: AIIC Normalized receive calculation failed")
                return None

            AIIC_contour_val, AIIC_contour_result = calc_AIIC_val_claude(AIIC_Normalized_recieve)
            print(f"AIIC_contour_val: {AIIC_contour_val}")
            print(f"AIIC_contour_result: {AIIC_contour_result}")
            
            return {
                'NR_val': NR_val,
                'AIIC_recieve_corr': AIIC_recieve_corr,
                'AIIC_Normalized_recieve': AIIC_Normalized_recieve,
                'positions': freq_data['positions'],
                'AIIC_contour_val': AIIC_contour_val,
                'AIIC_contour_result': AIIC_contour_result,
                'room_vol': float(freq_data['room_props'].receive_vol)
            }
            
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
            # Plot normalized receive levels on ax2
            ax2.plot(freq_data['target_freqs'], results['AIIC_Normalized_recieve'], 
                    label='Normalized Impact Sound Level', 
                    color='blue', marker='o')
            
            # Plot AIIC contour
            ax2.plot(freq_data['target_freqs'], IIC_contour_final, 
                    label=f'AIIC {results["AIIC_contour_val"]} Contour', 
                    color='red', linestyle='--')
            
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
            # Calculate ATL using processed data
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
            
            print(f"ATL values shape: {ATL_val.shape}")
            print(f"ATL values: {ATL_val}")
            
            # Calculate ASTC and contour
            ASTC_final_val = calc_astc_val(ATL_val)
            STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4, 4]
            ASTC_contour_val = [val + ASTC_final_val for val in STCCurve]
            
            print(f"ASTC final value: {ASTC_final_val}")
            
            return {
                'ATL_val': ATL_val,
                'ASTC_final_val': ASTC_final_val,
                'ASTC_contour_val': ASTC_contour_val,
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

    def _get_nic_raw_data(self, test_obj):
        """Extract raw data for NIC calculations"""
        try:
            raw_data = {
                'freq': test_obj.srs_data.raw_data['Frequency (Hz)'].values,
                'source': test_obj.srs_data.raw_data['Overall 1/3 Spectra'].values,
                'receive': test_obj.recive_data.raw_data['Overall 1/3 Spectra'].values,
                'background': test_obj.bkgrnd_data.raw_data['Overall 1/3 Spectra'].values,
                'rt': test_obj.rt.rt_thirty  # Already in correct format (17 points)
            }
            
            print("\nRaw data shapes:")
            print(f"Frequencies: {raw_data['freq'].shape}")
            print(f"Source: {raw_data['source'].shape}")
            print(f"Receive: {raw_data['receive'].shape}")
            print(f"Background: {raw_data['background'].shape}")
            print(f"RT: {raw_data['rt'].shape}")
            
            return raw_data
            
        except Exception as e:
            print(f"Error getting NIC raw data: {str(e)}")
            traceback.print_exc()
            return None

    def _process_nic_frequencies(self, raw_data):
        """Process frequency data for NIC calculations"""
        try:
            # Define target frequencies
            target_freqs = [100, 125, 160, 200, 250, 315, 400, 500, 
                           630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000]
            
            # Create frequency map
            freq_indices = []
            for freq in target_freqs:
                idx = np.where(raw_data['freq'] == freq)[0]
                if len(idx) > 0:
                    freq_indices.append(idx[0])
                
            # Extract correct frequency points
            processed_data = {
                'target_freqs': target_freqs,
                'source': raw_data['source'][freq_indices],
                'receive': raw_data['receive'][freq_indices],
                'background': raw_data['background'][freq_indices],
                'rt': raw_data['rt']  # Use full RT data
            }
            
            print("\nProcessed data shapes:")
            print(f"Source: {processed_data['source'].shape}")
            print(f"Receive: {processed_data['receive'].shape}")
            print(f"Background: {processed_data['background'].shape}")
            print(f"RT: {processed_data['rt'].shape}")
            
            return processed_data
            
        except Exception as e:
            print(f"Error processing NIC frequencies: {str(e)}")
            traceback.print_exc()
            return None

    def _calculate_nic_values(self, test_obj, freq_data):
        """Calculate NIC values"""
        try:
            room_vol = test_obj.room_properties.receive_vol
            
            # Calculate NIC using the processed data
            NR_val, NIC_final_val, sabines, _, _, _ = calc_NR_new(
                freq_data['source'],
                None,  # No AIIC data
                freq_data['receive'],
                freq_data['background'],
                room_vol,
                freq_data['rt']
            )
            
            if NR_val is not None:
                print(f"NR values shape: {NR_val.shape}")
                return {
                    'NR_val': NR_val,
                    'NIC_final_val': NIC_final_val,
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
                    label=f"Noise Reduction (NIC = {nic_results['NIC_final_val']})", 
                    color='blue', marker='o')
            
            # Plot NIC reference curve
            ref_curve = np.array([-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4, 4]) + nic_results['NIC_final_val']
            ax.plot(freq_values, ref_curve, 
                    label='NIC Reference Curve', 
                    color='red', linestyle='--')
            
            # Add results text box
            textstr = (f"NIC: {nic_results['NIC_final_val']}\n"
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

    def _get_astc_raw_data(self, test_obj):
        """Extract raw data for ASTC calculations"""
        try:
            raw_data = {
                'freq': test_obj.srs_data.raw_data['Frequency (Hz)'].values,
                'source': test_obj.srs_data.raw_data['Overall 1/3 Spectra'].values,
                'receive': test_obj.recive_data.raw_data['Overall 1/3 Spectra'].values,
                'background': test_obj.bkgrnd_data.raw_data['Overall 1/3 Spectra'].values,
                'rt': test_obj.rt.rt_thirty,  # Full RT data including 4000 Hz
                'room_props': test_obj.room_properties
            }
            
            print("\nRaw data shapes:")
            print(f"Frequencies: {raw_data['freq'].shape}")
            print(f"Source: {raw_data['source'].shape}")
            print(f"Receive: {raw_data['receive'].shape}")
            print(f"Background: {raw_data['background'].shape}")
            print(f"RT: {raw_data['rt'].shape}")
            
            return raw_data
            
        except Exception as e:
            print(f"Error getting ASTC raw data: {str(e)}")
            traceback.print_exc()
            return None

    def _process_astc_frequencies(self, raw_data):
        """Process frequency data for ASTC calculations"""
        try:
            # Define target frequencies (100-4000 Hz for ASTC)
            target_freqs = [100, 125, 160, 200, 250, 315, 400, 500, 
                           630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000]
            
            # Create frequency map
            freq_indices = []
            for freq in target_freqs:
                idx = np.where(raw_data['freq'] == freq)[0]
                if len(idx) > 0:
                    freq_indices.append(idx[0])
                else:
                    print(f"Warning: Frequency {freq} Hz not found in raw data")
                
            if len(freq_indices) != len(target_freqs):
                print(f"Error: Not all target frequencies found. Expected {len(target_freqs)}, got {len(freq_indices)}")
                return None
                
            # Extract correct frequency points
            processed_data = {
                'target_freqs': target_freqs,
                'source': raw_data['source'][freq_indices],
                'receive': raw_data['receive'][freq_indices],
                'background': raw_data['background'][freq_indices],
                'rt': raw_data['rt'],  # Already in correct format (17 points)
                'room_props': raw_data['room_props']
            }
            
            # Validate data shapes
            print("\nProcessed data shapes:")
            print(f"Target frequencies: {len(processed_data['target_freqs'])} points")
            print(f"Source: {processed_data['source'].shape}")
            print(f"Receive: {processed_data['receive'].shape}")
            print(f"Background: {processed_data['background'].shape}")
            print(f"RT: {processed_data['rt'].shape}")
            
            # Verify all arrays have the correct length (17 points)
            expected_length = 17
            if not all(len(processed_data[key]) == expected_length 
                      for key in ['source', 'receive', 'background', 'target_freqs']):
                print("Error: Processed arrays don't match expected length of 17 points")
                return None
            
            return processed_data
            
        except Exception as e:
            print(f"Error processing ASTC frequencies: {str(e)}")
            traceback.print_exc()
            return None

class MainApp(App):
    def build(self):
        return MainWindow()

if __name__ == '__main__':
    MainApp().run() 