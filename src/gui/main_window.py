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

from typing import Dict, List, Optional
import pandas as pd
import tempfile
import os
import fitz

from src.core.test_data_manager import TestDataManager
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
    calculate_onethird_Logavg
)
from src.gui.test_plan_input import TestPlanInputWindow

class MainWindow(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = 10
        self.spacing = 10
        
        # Initialize managers
        self.test_data_manager = TestDataManager()
        
        # Create UI Components
        self._create_file_inputs()
        self._create_test_controls()
        self._create_analysis_section()
        self._create_status_section()
        
    def _create_file_inputs(self):
        """Create file input section"""
        input_grid = GridLayout(cols=2, spacing=5, size_hint_y=0.4)
        
        # Test Plan Input
        input_grid.add_widget(Label(text='Test Plan:'))
        self.test_plan_path = TextInput(
            multiline=False,
            hint_text='Path to test plan Excel file'
        )
        input_grid.add_widget(self.test_plan_path)
        
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
        
        # Load Data Button
        self.load_button = Button(
            text='Load Data',
            size_hint_y=0.1
        )
        self.load_button.bind(on_press=self.load_data)
        self.add_widget(self.load_button)

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

    def load_data(self, instance):
        """Load and process test data files"""
        try:
            if not all([self.test_plan_path.text, self.slm_data_1_path.text, 
                       self.slm_data_2_path.text, self.output_path.text]):
                raise ValueError("All file paths must be specified")
                
            # Process test plan
            test_plan = pd.read_excel(self.test_plan_path.text)
            
            # Process SLM data
            slm_data_1 = extract_sound_levels(pd.read_excel(self.slm_data_1_path.text))
            slm_data_2 = extract_sound_levels(pd.read_excel(self.slm_data_2_path.text))
            
            # Update test data manager
            self.test_data_manager.process_test_data(
                test_plan, slm_data_1, slm_data_2, self.output_path.text
            )
            
            self.status_label.text = 'Status: Data loaded successfully'
            self.update_test_list()
            
        except Exception as e:
            self._show_error(f'Error loading data: {str(e)}')

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

    def open_test_plan_input(self, instance):
        """Open test plan input window"""
        content = TestPlanInputWindow(callback_on_save=self.on_test_plan_saved)
        popup = Popup(
            title='Test Plan Input',
            content=content,
            size_hint=(0.8, 0.9)
        )
        popup.open()

    def on_test_plan_saved(self, room_properties: RoomProperties, test_type: TestType):
        """Handle saved test plan data"""
        # Update test data manager
        self.test_data_manager.add_test(room_properties, test_type)
        # Update UI
        self.update_test_list()

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

class MainApp(App):
    def build(self):
        return MainWindow()

if __name__ == '__main__':
    MainApp().run() 