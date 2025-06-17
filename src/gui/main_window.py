"""
Refactored Main Window for ASTM Test Calculator

This is the refactored main window that uses the new service-oriented architecture.
The main window now focuses on UI coordination and delegates business logic to services.
"""

import traceback
import tempfile
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from kivy.uix.image import Image as KivyImage
from kivy.core.window import Window
from typing import Dict, Optional, Any
import os

# Import services
from src.core.test_data_manager import TestDataManager
from src.core.calculation_service import CalculationService
from src.core.validation_service import ValidationService, ValidationError
from src.gui.plotting_service import PlottingService

# Import UI components
from src.gui.components.file_input_panel import FileInputPanel
from src.gui.components.test_control_panel import TestControlPanel
from src.gui.components.status_panel import StatusPanel

# Import existing UI components
from src.gui.test_plan_input import TestPlanInputWindow
from src.gui.analysis_dashboard import ResultsAnalysisDashboard
from src.gui.test_plan_manager import TestPlanManagerWindow

# Import data processing components
from src.core.data_processor import TestType
from src.reports.test_data_exporter import export_test_results


class MainWindow(BoxLayout):
    """Refactored main window using service-oriented architecture"""
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.spacing = 10
        self.padding = 10
        
        # Initialize services
        self._initialize_services()
        
        # Initialize UI components
        self._initialize_ui_components()
        
        # Set up UI layout
        self._setup_layout()
        
        # Bind events
        self._bind_events()
    
    def _initialize_services(self):
        """Initialize all services"""
        # Core services
        self.test_data_manager = TestDataManager(debug_mode=False)
        self.calculation_service = CalculationService(debug_mode=False)
        self.validation_service = ValidationService(debug_mode=False)
        self.plotting_service = PlottingService(debug_mode=False)
        
        # State management
        self.temp_plot_files = []  # Track temporary plot files for cleanup
    
    def _initialize_ui_components(self):
        """Initialize UI components"""
        # File input panel
        self.file_input_panel = FileInputPanel(
            on_file_selected_callback=self._on_file_selected
        )
        
        # Test control panel
        self.test_control_panel = TestControlPanel()
        
        # Status panel
        self.status_panel = StatusPanel()
        
        # Analysis section
        self._create_analysis_section()
    
    def _setup_layout(self):
        """Set up the main layout"""
        # Add file input panel
        self.add_widget(self.file_input_panel)
        
        # Add test control panel
        self.add_widget(self.test_control_panel)
        
        # Add analysis section
        self.add_widget(self.analysis_tabs)
        
        # Add status panel
        self.add_widget(self.status_panel)
    
    def _bind_events(self):
        """Bind events to UI components"""
        # File input panel events
        self.file_input_panel.bind_load_button(self.load_data)
        self.file_input_panel.bind_populate_button(self.populate_test_inputs)
        
        # Test control panel events
        self.test_control_panel.bind_report_button(self.generate_reports)
        self.test_control_panel.bind_plot_button(self.show_plot_selection)
        self.test_control_panel.bind_test_plan_manager_button(self.show_test_plan_manager)
        self.test_control_panel.bind_test_plan_input_button(self.show_test_plan_input)
        self.test_control_panel.bind_refresh_button(self.refresh_results_dashboard)
        self.test_control_panel.bind_debug_checkbox(self._on_debug_mode_changed)
    
    def _create_analysis_section(self):
        """Create analysis dashboard section"""
        self.analysis_tabs = TabbedPanel(size_hint_y=0.5, do_default_tab=False)
        
        # Test Plan Tab
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
        
        # Results Tab
        self.results_tab = TabbedPanelItem(text='Results Dashboard')
        self.results_dashboard = ResultsAnalysisDashboard(self.test_data_manager)
        self.results_tab.add_widget(self.results_dashboard)
        self.analysis_tabs.add_widget(self.results_tab)
        
        # Raw Data Tab
        self.raw_data_tab = TabbedPanelItem(text='Raw Data')
        self.raw_data_scroll = ScrollView()
        self.raw_data_grid = GridLayout(
            cols=1,
            spacing=2,
            size_hint_y=None,
            padding=5
        )
        self.raw_data_grid.bind(minimum_height=self.raw_data_grid.setter('height'))
        
        self.raw_data_scroll.add_widget(self.raw_data_grid)
        self.raw_data_tab.add_widget(self.raw_data_scroll)
        self.analysis_tabs.add_widget(self.raw_data_tab)
    
    # Event handlers
    def _on_file_selected(self, text_input, file_path):
        """Handle file selection events"""
        if self._get_debug_mode():
            print(f"File selected: {file_path}")
    
    def _on_debug_mode_changed(self, checkbox, value):
        """Handle debug mode changes"""
        debug_mode = value
        
        # Update services with new debug mode
        self.calculation_service.debug_mode = debug_mode
        self.validation_service.debug_mode = debug_mode
        self.plotting_service.debug_mode = debug_mode
        self.test_data_manager.debug_mode = debug_mode
        
        if debug_mode:
            self.status_panel.set_status("Debug mode enabled", "info")
        else:
            self.status_panel.set_status("Debug mode disabled", "info")
    
    # Main application actions
    def load_data(self, instance):
        """Load and process all test data using validation service"""
        try:
            self.status_panel.set_status("Loading data...", "info")
            
            # Get paths from file input panel
            paths = self.file_input_panel.get_paths()
            
            # Validate all paths using validation service
            try:
                self.validation_service.validate_all_paths(
                    test_plan_path=paths['test_plan_path'],
                    slm_data_d_path=paths['slm_data_1_path'],
                    slm_data_e_path=paths['slm_data_2_path'],
                    report_output_path=paths['output_path']
                )
            except ValidationError as e:
                error_message = self.validation_service.suggest_error_fixes(str(e))
                self.status_panel.show_error(error_message)
                self.status_panel.set_status("Validation failed", "error")
                return False
            
            # Set paths in test data manager
            self.test_data_manager.set_data_paths(
                test_plan_path=paths['test_plan_path'],
                meter_d_path=paths['slm_data_1_path'],
                meter_e_path=paths['slm_data_2_path'],
                output_path=paths['output_path']
            )
            
            # Load test plan
            self.status_panel.set_status("Loading test plan...", "info")
            self.test_data_manager.load_test_plan()
            
            # Process test data
            self.status_panel.set_status("Processing test data...", "info")
            self.test_data_manager.process_test_data()
            
            # Update displays
            self.update_displays()
            
            # Update button states
            self.test_control_panel.update_button_states(data_loaded=True)
            
            self.status_panel.set_status("All test data loaded successfully", "success")
            
            # Show plot selection popup after successful load
            self.show_plot_selection(instance)
            
            return True
            
        except Exception as e:
            error_message = f"Error loading data: {str(e)}"
            if self._get_debug_mode():
                print(f"\nERROR: {error_message}")
                traceback.print_exc()
            
            self.status_panel.show_error(error_message)
            self.status_panel.set_status("Data loading failed", "error")
            return False
    
    def populate_test_inputs(self, instance):
        """Populate input fields with example data"""
        self.file_input_panel.populate_example_paths()
        self.status_panel.set_status("Example paths populated", "info")
    
    def generate_reports(self, instance):
        """Generate test reports using test data exporter"""
        try:
            self.status_panel.set_status("Generating reports...", "info")
            
            # Get output path
            paths = self.file_input_panel.get_paths()
            output_path = paths['output_path']
            
            # Generate reports using existing exporter
            success = export_test_results(self.test_data_manager, output_path)
            
            if success:
                self.status_panel.set_status("Reports generated successfully", "success")
                self.status_panel.show_success(f"Reports generated successfully in {output_path}")
            else:
                self.status_panel.set_status("Report generation failed", "error")
                self.status_panel.show_error("Failed to generate reports")
                
        except Exception as e:
            error_message = f"Error generating reports: {str(e)}"
            if self._get_debug_mode():
                print(f"\nERROR: {error_message}")
                traceback.print_exc()
            
            self.status_panel.show_error(error_message)
            self.status_panel.set_status("Report generation failed", "error")
    
    def show_plot_selection(self, instance):
        """Show plot selection dialog"""
        try:
            test_collection = self.test_data_manager.get_test_collection()
            
            if not test_collection:
                self.status_panel.show_warning("No test data loaded. Please load data first.")
                return
            
            # Create plot selection popup
            content = BoxLayout(orientation='vertical', padding=10, spacing=10)
            
            # Title
            content.add_widget(Label(
                text='Select tests to plot:',
                size_hint_y=None,
                height=30
            ))
            
            # Create scrollable list of tests
            scroll = ScrollView()
            test_grid = GridLayout(cols=1, spacing=5, size_hint_y=None)
            test_grid.bind(minimum_height=test_grid.setter('height'))
            
            checkboxes = {}
            for test_label, test_data in test_collection.items():
                for test_type in test_data.keys():
                    test_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=30)
                    
                    checkbox = CheckBox(size_hint_x=None, width=30)
                    test_layout.add_widget(checkbox)
                    test_layout.add_widget(Label(text=f"{test_label} - {test_type.name}"))
                    
                    test_grid.add_widget(test_layout)
                    checkboxes[(test_label, test_type)] = checkbox
            
            scroll.add_widget(test_grid)
            content.add_widget(scroll)
            
            # Buttons
            button_layout = BoxLayout(size_hint_y=None, height=50, spacing=10)
            
            plot_button = Button(text='Plot Selected')
            def plot_selected(instance):
                selected_tests = [(label, test_type) for (label, test_type), cb in checkboxes.items() if cb.active]
                popup.dismiss()
                if selected_tests:
                    self._plot_selected_tests(selected_tests)
                else:
                    self.status_panel.show_warning("No tests selected for plotting")
            
            plot_button.bind(on_press=plot_selected)
            button_layout.add_widget(plot_button)
            
            cancel_button = Button(text='Cancel')
            cancel_button.bind(on_press=lambda x: popup.dismiss())
            button_layout.add_widget(cancel_button)
            
            content.add_widget(button_layout)
            
            popup = Popup(
                title='Select Tests to Plot',
                content=content,
                size_hint=(0.8, 0.8)
            )
            popup.open()
            
        except Exception as e:
            error_message = f"Error showing plot selection: {str(e)}"
            if self._get_debug_mode():
                print(f"\nERROR: {error_message}")
                traceback.print_exc()
            self.status_panel.show_error(error_message)
    
    def _plot_selected_tests(self, selected_tests):
        """Plot the selected tests using plotting service"""
        try:
            for test_label, test_type in selected_tests:
                self.status_panel.set_status(f"Plotting {test_label} - {test_type.name}...", "info")
                success = self._plot_single_test(test_label, test_type)
                
                if not success:
                    self.status_panel.show_error(f"Failed to plot {test_label} - {test_type.name}")
                    
        except Exception as e:
            error_message = f"Error plotting tests: {str(e)}"
            if self._get_debug_mode():
                print(f"\nERROR: {error_message}")
                traceback.print_exc()
            self.status_panel.show_error(error_message)
    
    def _plot_single_test(self, test_label: str, test_type: TestType) -> bool:
        """Plot a single test using the plotting service"""
        try:
            # Get test data
            test_collection = self.test_data_manager.get_test_collection()
            test_data = test_collection[test_label][test_type]
            test_obj = test_data['test_data']
            
            # Get raw data and process frequencies (this logic would need to be extracted from original main_window.py)
            # For now, we'll use a simplified approach
            freq_data = self._get_frequency_data(test_obj, test_type)
            if not freq_data:
                return False
            
            # Calculate values using calculation service
            calculated_values = self.calculation_service.calculate_test_values(test_type, test_obj, freq_data)
            if not calculated_values:
                return False
            
            # Create plot using plotting service
            plot_path = self.plotting_service.create_test_plot(test_type, test_obj, freq_data, calculated_values)
            if not plot_path:
                return False
            
            # Track temp file for cleanup
            self.temp_plot_files.append(plot_path)
            
            # Show plot in popup
            self._show_plot_popup(plot_path, f"{test_label} - {test_type.name}", test_label)
            
            return True
            
        except Exception as e:
            if self._get_debug_mode():
                print(f"Error plotting {test_label} - {test_type.name}: {str(e)}")
                traceback.print_exc()
            return False
    
    def _get_frequency_data(self, test_obj, test_type: TestType) -> Optional[Dict[str, Any]]:
        """Get frequency data for the test (simplified version)"""
        # This is a simplified version - the full implementation would need to extract
        # the frequency processing logic from the original main_window.py
        try:
            # For now, return a basic structure
            # In full implementation, this would process the SLM data
            return {
                'source': [],
                'background': [],
                'rt': [],
                'room_props': test_obj.room_properties
            }
        except Exception as e:
            if self._get_debug_mode():
                print(f"Error getting frequency data: {str(e)}")
            return None
    
    def _show_plot_popup(self, plot_path: str, title: str, test_label: str):
        """Show plot in popup with store button"""
        try:
            # Create image from plot file
            image = KivyImage(
                source=plot_path,
                allow_stretch=True,
                keep_ratio=True,
                size_hint=(1, 1)
            )
            
            # Create button layout
            button_layout = BoxLayout(
                size_hint_y=None,
                height='48dp',
                spacing='10dp',
                padding='10dp'
            )
            
            store_button = Button(
                text='Store Calculated Values',
                size_hint_x=None,
                width='200dp'
            )
            store_button.bind(on_press=lambda x: self._store_calculated_values(test_label))
            button_layout.add_widget(store_button)
            
            # Main layout
            layout = BoxLayout(orientation='vertical')
            layout.add_widget(image)
            layout.add_widget(button_layout)
            
            # Show popup
            plot_popup = Popup(
                title=f'Test Data Plot - {title}',
                content=layout,
                size_hint=(0.9, 0.9)
            )
            plot_popup.open()
            
        except Exception as e:
            if self._get_debug_mode():
                print(f"Error showing plot popup: {str(e)}")
                traceback.print_exc()
    
    def _store_calculated_values(self, test_label: str):
        """Store calculated values for a test"""
        try:
            self.status_panel.set_status(f"Storing calculated values for {test_label}...", "info")
            
            # Get test data
            test_data = self.test_data_manager.test_data_collection[test_label]
            
            for test_type in test_data.keys():
                test_obj = test_data[test_type]['test_data']
                
                # Get frequency data and calculate values
                freq_data = self._get_frequency_data(test_obj, test_type)
                if freq_data:
                    calculated_values = self.calculation_service.calculate_test_values(test_type, test_obj, freq_data)
                    if calculated_values:
                        # Store values using test processor
                        self.test_data_manager.test_processor.store_calculated_values(
                            test_label, test_type, calculated_values
                        )
            
            # Update results dashboard
            self.refresh_results_dashboard(None)
            
            self.status_panel.set_status("Calculated values stored successfully", "success")
            self.status_panel.show_success("Calculated values stored successfully")
            
        except Exception as e:
            error_message = f"Error storing calculated values: {str(e)}"
            if self._get_debug_mode():
                print(f"\nERROR: {error_message}")
                traceback.print_exc()
            self.status_panel.show_error(error_message)
    
    # UI helper methods
    def show_test_plan_input(self, instance):
        """Show test plan input window"""
        try:
            test_plan_window = TestPlanInputWindow()
            test_plan_window.show()
        except Exception as e:
            self.status_panel.show_error(f"Error opening test plan input: {str(e)}")
    
    def show_test_plan_manager(self, instance):
        """Show test plan manager window"""
        try:
            test_plan_manager = TestPlanManagerWindow()
            test_plan_manager.show()
        except Exception as e:
            self.status_panel.show_error(f"Error opening test plan manager: {str(e)}")
    
    def update_displays(self):
        """Update all displays with current data"""
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
            if self._get_debug_mode():
                print(f"Error in update_displays: {str(e)}")
            self.status_panel.set_status(f'Error updating displays - {str(e)}', "error")
    
    def _update_test_plan_display(self):
        """Update test plan display"""
        try:
            test_plan = self.test_data_manager.test_plan
            if test_plan is not None and not test_plan.empty:
                for idx, row in test_plan.iterrows():
                    row_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=30)
                    for col in test_plan.columns:
                        row_layout.add_widget(Label(text=str(row[col]), size_hint_x=None, width=100))
                    self.test_plan_grid.add_widget(row_layout)
        except Exception as e:
            if self._get_debug_mode():
                print(f"Error updating test plan display: {str(e)}")
    
    def refresh_results_dashboard(self, instance):
        """Refresh the results dashboard"""
        try:
            if hasattr(self, 'results_dashboard'):
                self.results_dashboard.test_data_manager = self.test_data_manager
                if hasattr(self.results_dashboard, 'refresh_results'):
                    self.results_dashboard.refresh_results()
                
                # Switch to results tab
                if hasattr(self, 'results_tab') and hasattr(self, 'analysis_tabs'):
                    self.analysis_tabs.switch_to(self.results_tab)
                    
            self.status_panel.set_status("Results dashboard refreshed", "info")
                    
        except Exception as e:
            if self._get_debug_mode():
                print(f"Error refreshing results dashboard: {str(e)}")
                traceback.print_exc()
            self.status_panel.show_error(f"Error refreshing results dashboard: {str(e)}")
    
    def _get_debug_mode(self) -> bool:
        """Get current debug mode state"""
        return self.test_control_panel.get_debug_mode()
    
    def cleanup(self):
        """Clean up resources"""
        # Clean up temporary plot files
        if self.temp_plot_files:
            self.plotting_service.cleanup_temp_plots(self.temp_plot_files)
            self.temp_plot_files.clear()


class MainApp(App):
    """Main application class"""
    
    def build(self):
        Window.size = (1200, 800)
        return MainWindow()
    
    def on_stop(self):
        """Clean up when app stops"""
        if hasattr(self.root, 'cleanup'):
            self.root.cleanup()


if __name__ == '__main__':
    MainApp().run()