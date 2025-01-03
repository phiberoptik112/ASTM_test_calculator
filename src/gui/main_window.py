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
    sanitize_filepath
)
from src.core.test_processor import TestProcessor

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
        """Generate PDF report for test"""
        try:
            from src.reports.base_report import BaseReport
            
            report = BaseReport(
                test_data=test_data,
                output_folder=self.output_path.text,
                test_type=test_data.test_type
            )
            
            pdf_path = report.generate()
            self.preview_pdf(pdf_path)
            self.status_label.text = 'Status: Report generated successfully'
            
        except Exception as e:
            self._show_error(f'Error generating report: {str(e)}')

    def show_plot_selection(self, instance):
        """Show popup with buttons to plot data for each test"""
        # Create content layout
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)
        
        # Add title label
        content.add_widget(Label(
            text='Select a test to plot:',
            size_hint_y=None,
            height=40
        ))
        
        # Create scrollable button list
        scroll = ScrollView(size_hint=(1, 0.8))
        button_layout = GridLayout(
            cols=1,
            spacing=5,
            size_hint_y=None
        )
        button_layout.bind(minimum_height=button_layout.setter('height'))
        
        # Add button for each test
        for test_label in self.test_data_manager.test_data_collection:
            btn = Button(
                text=f'Plot Test {test_label}',
                size_hint_y=None,
                height=40
            )
            btn.bind(on_press=lambda x, label=test_label: self.plot_test_data(label))
            button_layout.add_widget(btn)
        
        scroll.add_widget(button_layout)
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

    def plot_test_data(self, test_label):
        """Create plot for a specific test's data"""
        try:
            # Get test data
            test_data = self.test_data_manager.test_data_collection[test_label]
            
            # Create figure
            plt.figure(figsize=(10, 6))
            
            # Plot each data series
            data_types = {
                'Source Room': 'srs_data',
                'Receive Room': 'recive_data',
                'Background': 'bkgrnd_data'
            }
            
            for test_type, test_info in test_data.items():
                if 'test_data' in test_info:
                    test_obj = test_info['test_data']
                    
                    for label, attr in data_types.items():
                        if hasattr(test_obj, attr):
                            data = getattr(test_obj, attr)
                            if hasattr(data, 'raw_data'):
                                df = data.raw_data
                                if 'Frequency (Hz)' in df.columns and 'Overall 1/3 Spectra' in df.columns:
                                    plt.semilogx(
                                        df['Frequency (Hz)'],
                                        df['Overall 1/3 Spectra'],
                                        label=f'{label} - {test_type.value}'
                                    )
            
            # Configure plot
            plt.title(f'Raw Data for Test {test_label}')
            plt.xlabel('Frequency (Hz)')
            plt.ylabel('Sound Pressure Level (dB)')
            plt.grid(True)
            plt.legend()
            
            # Add vertical gridlines at standard frequencies
            standard_freqs = [63, 125, 250, 500, 1000, 2000, 4000]
            plt.gca().set_xticks(standard_freqs)
            plt.gca().set_xticklabels([str(f) for f in standard_freqs])
            plt.grid(True, which='both', linestyle='--', alpha=0.7)
            
            # Show plot in a popup
            plot_popup = Popup(
                title=f'Test {test_label} Data Plot',
                size_hint=(0.9, 0.9)
            )
            
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
                
                # Add to popup
                plot_popup.content = scroll
            
            plot_popup.open()
            plt.close()
            
        except Exception as e:
            self._show_error(f'Error plotting test data: {str(e)}')

class MainApp(App):
    def build(self):
        return MainWindow()

if __name__ == '__main__':
    MainApp().run() 