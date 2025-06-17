"""
Test Control Panel Component

This component handles test control buttons and actions for the ASTM Test Calculator.
Extracted from main_window.py to separate test control UI logic.
"""

from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.label import Label
from typing import Callable, Optional


class TestControlPanel(BoxLayout):
    """Panel for handling test control actions"""
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.spacing = 5
        self.size_hint_y = 0.3
        
        # Initialize control widgets
        self.report_button = None
        self.plot_button = None
        self.debug_checkbox = None
        
        self._create_control_widgets()
    
    def _create_control_widgets(self):
        """Create the test control widgets"""
        # Debug checkbox
        debug_layout = BoxLayout(
            orientation='horizontal',
            size_hint_y=None,
            height='32dp',
            spacing=5
        )
        
        self.debug_checkbox = CheckBox(
            active=False,
            size_hint_x=None,
            width='32dp'
        )
        debug_layout.add_widget(self.debug_checkbox)
        debug_layout.add_widget(Label(
            text='Debug Mode',
            size_hint_x=None,
            width='100dp'
        ))
        
        # Add spacer to push debug controls to the left
        debug_layout.add_widget(Label(text=''))  # Spacer
        
        self.add_widget(debug_layout)
        
        # Main control buttons
        button_layout = BoxLayout(
            spacing=5,
            size_hint_y=None,
            height='48dp'
        )
        
        # Generate Reports Button
        self.report_button = Button(
            text='Generate Reports',
            size_hint_x=0.5
        )
        button_layout.add_widget(self.report_button)
        
        # Raw Loaded Data Plot Button
        self.plot_button = Button(
            text='Raw Loaded Data Plot',
            size_hint_x=0.5
        )
        button_layout.add_widget(self.plot_button)
        
        self.add_widget(button_layout)
        
        # Additional action buttons
        action_layout = BoxLayout(
            spacing=5,
            size_hint_y=None,
            height='48dp'
        )
        
        # Test Plan Manager Button
        self.test_plan_manager_button = Button(
            text='Test Plan Manager',
            size_hint_x=0.33
        )
        action_layout.add_widget(self.test_plan_manager_button)
        
        # Test Plan Input Button
        self.test_plan_input_button = Button(
            text='Open Test Plan GUI',
            size_hint_x=0.33
        )
        action_layout.add_widget(self.test_plan_input_button)
        
        # Refresh Results Button
        self.refresh_button = Button(
            text='Refresh Results',
            size_hint_x=0.34
        )
        action_layout.add_widget(self.refresh_button)
        
        self.add_widget(action_layout)
    
    def bind_report_button(self, callback: Callable):
        """Bind callback to generate reports button"""
        if self.report_button:
            self.report_button.bind(on_press=callback)
    
    def bind_plot_button(self, callback: Callable):
        """Bind callback to plot data button"""
        if self.plot_button:
            self.plot_button.bind(on_press=callback)
    
    def bind_test_plan_manager_button(self, callback: Callable):
        """Bind callback to test plan manager button"""
        if self.test_plan_manager_button:
            self.test_plan_manager_button.bind(on_press=callback)
    
    def bind_test_plan_input_button(self, callback: Callable):
        """Bind callback to test plan input button"""
        if self.test_plan_input_button:
            self.test_plan_input_button.bind(on_press=callback)
    
    def bind_refresh_button(self, callback: Callable):
        """Bind callback to refresh results button"""
        if self.refresh_button:
            self.refresh_button.bind(on_press=callback)
    
    def bind_debug_checkbox(self, callback: Callable):
        """Bind callback to debug checkbox"""
        if self.debug_checkbox:
            self.debug_checkbox.bind(active=callback)
    
    def set_button_enabled(self, button_name: str, enabled: bool):
        """Enable or disable specific buttons"""
        button_map = {
            'report': self.report_button,
            'plot': self.plot_button,
            'test_plan_manager': self.test_plan_manager_button,
            'test_plan_input': self.test_plan_input_button,
            'refresh': self.refresh_button
        }
        
        button = button_map.get(button_name)
        if button:
            button.disabled = not enabled
    
    def get_debug_mode(self) -> bool:
        """Get the current debug mode state"""
        return self.debug_checkbox.active if self.debug_checkbox else False
    
    def set_debug_mode(self, debug: bool):
        """Set the debug mode state"""
        if self.debug_checkbox:
            self.debug_checkbox.active = debug
    
    def update_button_states(self, data_loaded: bool = False, results_available: bool = False):
        """Update button states based on application state"""
        # Enable plot button only if data is loaded
        self.set_button_enabled('plot', data_loaded)
        
        # Enable report button only if results are available
        self.set_button_enabled('report', results_available)
        
        # Test plan buttons are always available
        self.set_button_enabled('test_plan_manager', True)
        self.set_button_enabled('test_plan_input', True)
        
        # Refresh button is available when data is loaded
        self.set_button_enabled('refresh', data_loaded)