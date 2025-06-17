"""
File Input Panel Component

This component handles all file path inputs and file selection for the ASTM Test Calculator.
Extracted from main_window.py to separate file input UI logic.
"""

import os
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.filechooser import FileChooserListView
from typing import Callable, List, Tuple, Optional


class FileInputPanel(BoxLayout):
    """Panel for handling file path inputs and file selection"""
    
    def __init__(self, on_file_selected_callback: Optional[Callable] = None, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.spacing = 10
        self.size_hint_y = 0.4
        
        # Callback for when files are selected
        self.on_file_selected_callback = on_file_selected_callback
        
        # Initialize input widgets
        self.test_plan_path = None
        self.slm_data_1_path = None
        self.slm_data_2_path = None
        self.output_path = None
        
        self._create_input_widgets()
        self._create_control_buttons()
    
    def _create_input_widgets(self):
        """Create the file input widgets"""
        input_grid = GridLayout(
            cols=2,
            spacing=5
        )
        
        # Test Plan Input
        input_grid.add_widget(Label(text='Test Plan:'))
        test_plan_layout = self._create_path_input_layout(
            hint_text='Path to test plan Excel file',
            file_filters=[('Excel Files', '*.xlsx'), ('CSV Files', '*.csv'), ('All Files', '*.*')],
            is_directory=False
        )
        self.test_plan_path = test_plan_layout['text_input']
        input_grid.add_widget(test_plan_layout['layout'])
        
        # SLM Data Folder 1
        input_grid.add_widget(Label(text='SLM Raw Data Folder 1:'))
        slm1_layout = self._create_path_input_layout(
            hint_text='Path to first SLM data folder',
            file_filters=[('All Files', '*.*')],
            is_directory=True
        )
        self.slm_data_1_path = slm1_layout['text_input']
        input_grid.add_widget(slm1_layout['layout'])
        
        # SLM Data Folder 2
        input_grid.add_widget(Label(text='SLM Raw Data Folder 2:'))
        slm2_layout = self._create_path_input_layout(
            hint_text='Path to second SLM data folder',
            file_filters=[('All Files', '*.*')],
            is_directory=True
        )
        self.slm_data_2_path = slm2_layout['text_input']
        input_grid.add_widget(slm2_layout['layout'])
        
        # Output Folder
        input_grid.add_widget(Label(text='Output Folder:'))
        output_layout = self._create_path_input_layout(
            hint_text='Path for output reports',
            file_filters=[('All Files', '*.*')],
            is_directory=True
        )
        self.output_path = output_layout['text_input']
        input_grid.add_widget(output_layout['layout'])
        
        self.add_widget(input_grid)
    
    def _create_path_input_layout(self, hint_text: str, file_filters: List[Tuple[str, str]], 
                                 is_directory: bool = False) -> dict:
        """Create a path input layout with text input and file picker button"""
        layout = BoxLayout(orientation='horizontal', spacing=5)
        
        # Text input
        text_input = TextInput(
            multiline=False,
            hint_text=hint_text,
            size_hint_x=0.85
        )
        layout.add_widget(text_input)
        
        # File picker button
        file_picker_btn = Button(
            text='...',
            size_hint_x=0.15
        )
        file_picker_btn.bind(on_press=lambda x: self.show_file_picker(
            text_input, file_filters, is_directory))
        layout.add_widget(file_picker_btn)
        
        return {'layout': layout, 'text_input': text_input}
    
    def _create_control_buttons(self):
        """Create control buttons for the file input panel"""
        button_layout = BoxLayout(
            orientation='horizontal',
            spacing=5,
            size_hint_y=None,
            height='48dp'
        )
        
        # Load Data Button
        self.load_button = Button(
            text='Load Data',
            size_hint_x=0.5
        )
        button_layout.add_widget(self.load_button)
        
        # Populate Test Inputs Button
        self.populate_button = Button(
            text='Populate Test Inputs',
            size_hint_x=0.5
        )
        button_layout.add_widget(self.populate_button)
        
        self.add_widget(button_layout)
    
    def show_file_picker(self, target_input: TextInput, filters: List[Tuple[str, str]], 
                        is_directory: bool = False):
        """Show a file picker dialog and update the target input with selected path"""
        content = BoxLayout(orientation='vertical')
        
        # Get the workspace directory path
        workspace_dir = os.path.abspath(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))))
        
        # Create file chooser with proper filters
        def custom_filter(folder, filename):
            if os.path.isdir(os.path.join(folder, filename)):
                return True
            return any(filename.endswith(ext[1:]) for _, ext in filters)
            
        file_chooser = FileChooserListView(
            path=workspace_dir,
            filters=[custom_filter],
            size_hint=(1, 0.9),
            dirselect=is_directory
        )
        content.add_widget(file_chooser)
        
        # Buttons
        button_layout = BoxLayout(size_hint_y=None, height='48dp', spacing='10dp')
        
        # Select button
        select_btn = Button(text='Select')
        def on_select(instance):
            if file_chooser.selection:
                selected_path = file_chooser.selection[0]
                target_input.text = selected_path
                if self.on_file_selected_callback:
                    self.on_file_selected_callback(target_input, selected_path)
            popup.dismiss()
        select_btn.bind(on_press=on_select)
        button_layout.add_widget(select_btn)
        
        # Cancel button
        cancel_btn = Button(text='Cancel')
        cancel_btn.bind(on_press=lambda x: popup.dismiss())
        button_layout.add_widget(cancel_btn)
        
        content.add_widget(button_layout)
        
        # Show popup
        popup = Popup(
            title='Select File' if not is_directory else 'Select Directory',
            content=content,
            size_hint=(0.8, 0.8)
        )
        popup.open()
    
    def populate_example_paths(self):
        """Populate input fields with example data paths"""
        self.test_plan_path.text = "./Exampledata/Example_TestPlan.csv"
        self.slm_data_1_path.text = "./Exampledata/RawData/A_Meter/"
        self.slm_data_2_path.text = "./Exampledata/RawData/E_Meter/"
        self.output_path.text = "./Exampledata/testeroutputs/"
    
    def get_paths(self) -> dict:
        """Get all current path values"""
        return {
            'test_plan_path': self.test_plan_path.text,
            'slm_data_1_path': self.slm_data_1_path.text,
            'slm_data_2_path': self.slm_data_2_path.text,
            'output_path': self.output_path.text
        }
    
    def set_paths(self, paths: dict):
        """Set path values from dictionary"""
        if 'test_plan_path' in paths:
            self.test_plan_path.text = paths['test_plan_path']
        if 'slm_data_1_path' in paths:
            self.slm_data_1_path.text = paths['slm_data_1_path']
        if 'slm_data_2_path' in paths:
            self.slm_data_2_path.text = paths['slm_data_2_path']
        if 'output_path' in paths:
            self.output_path.text = paths['output_path']
    
    def clear_paths(self):
        """Clear all path inputs"""
        self.test_plan_path.text = ""
        self.slm_data_1_path.text = ""
        self.slm_data_2_path.text = ""
        self.output_path.text = ""
    
    def bind_load_button(self, callback: Callable):
        """Bind callback to load button"""
        self.load_button.bind(on_press=callback)
    
    def bind_populate_button(self, callback: Callable):
        """Bind callback to populate button"""
        self.populate_button.bind(on_press=callback)