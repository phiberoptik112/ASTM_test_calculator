from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.checkbox import CheckBox
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView
from data_processor import RoomProperties, TestType
import csv
import os
from datetime import datetime
from kivy.uix.popup import Popup

class TestPlanInputWindow(BoxLayout):
    def __init__(self, callback_on_save=None, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.spacing = 10
        self.padding = 10
        self.callback_on_save = callback_on_save
        
        # Create a ScrollView to contain all fields
        scroll = ScrollView(size_hint=(1, 0.9))
        
        # Main form layout
        self.form = GridLayout(
            cols=2,
            spacing=5,
            padding=5,
            size_hint_y=None
        )
        # Make sure the height is properly set for scrolling
        self.form.bind(minimum_height=self.form.setter('height'))
        
        # Initialize input fields dictionary
        self.input_fields = {}
        
        # Create all test plan fields
        self._create_test_plan_fields()
        
        # Add form to scroll view
        scroll.add_widget(self.form)
        self.add_widget(scroll)
        
        # Create buttons
        self._create_buttons()

    def _create_test_plan_fields(self):
        """Create all fields matching test plan structure"""
        # Define all fields with their display names - exact match to test plan columns
        fields = [
            ('Test_Label', 'Test Label'),
            ('Client_Name', 'Client Name'),
            ('Site_Name', 'Site Name'),
            ('Source Room', 'Source Room'),
            ('Receiving Room', 'Receiving Room'),
            ('Test Date', 'Test Date'),
            ('Report Date', 'Report Date'),
            ('Project Name', 'Project Name'),
            ('source room vol', 'Source Room Volume'),
            ('receive room vol', 'Receive Room Volume'),
            ('partition area', 'Partition Area'),
            ('partition dim', 'Partition Dimensions'),
            ('source room finish', 'Source Room Finish'),
            ('receive room finish', 'Receive Room Finish'),
            ('srs floor descrip.', 'Source Room Floor'),
            ('srs walls descrip.', 'Source Room Walls'),
            ('srs ceiling descrip.', 'Source Room Ceiling'),
            ('rec floor descrip.', 'Receive Room Floor'),
            ('rec walls descrip.', 'Receive Room Walls'),
            ('rec ceiling descrip.', 'Receive Room Ceiling'),
            ('tested assembly', 'Tested Assembly'),
            ('expected performance', 'Expected Performance'),
            ('Test assembly Type', 'Test Assembly Type'),
            ('Annex 2 used?', 'Annex 2 Used'),
            ('AIIC', 'AIIC Test'),
            ('ASTC', 'ASTC Test'),
            ('NIC', 'NIC Test'),
            ('DTC', 'DTC Test')
        ]
        
        # Add fields to form
        for field_id, display_name in fields:
            # Add label
            self.form.add_widget(Label(
                text=display_name,
                size_hint_y=None,
                height=40
            ))
            
            # Add input widget based on field type
            if field_id in ['Annex 2 used?', 'AIIC', 'ASTC', 'NIC', 'DTC']:
                input_widget = CheckBox(
                    size_hint_y=None,
                    height=40,
                    active=False
                )
            else:
                input_widget = TextInput(
                    multiline=False,
                    size_hint_y=None,
                    height=40
                )
            self.input_fields[field_id] = input_widget
            self.form.add_widget(input_widget)

    def _create_buttons(self):
        """Create save and cancel buttons"""
        button_layout = BoxLayout(
            orientation='horizontal',
            size_hint_y=0.1,
            spacing=10,
            padding=5
        )
        
        # Save button
        save_button = Button(text='Save')
        save_button.bind(on_press=self.save_data)
        button_layout.add_widget(save_button)
        
        # Cancel button
        cancel_button = Button(text='Cancel')
        cancel_button.bind(on_press=self.cancel)
        button_layout.add_widget(cancel_button)
        
        self.add_widget(button_layout)

    def save_data(self, instance):
        """Collect and save the input data"""
        try:
            # Collect all field values
            test_data = {}
            for field_id, input_widget in self.input_fields.items():
                if isinstance(input_widget, CheckBox):
                    test_data[field_id] = 1 if input_widget.active else 0
                else:
                    test_data[field_id] = input_widget.text.strip()
            
            # Validate required fields
            required_fields = ['Test_Label', 'Client_Name', 'Site_Name']
            missing_fields = [f for f in required_fields if not test_data.get(f)]
            if missing_fields:
                raise ValueError(f"Required fields missing: {', '.join(missing_fields)}")
            
            # Validate at least one test type is selected
            test_types = ['AIIC', 'ASTC', 'NIC', 'DTC']
            if not any(test_data.get(tt, 0) == 1 for tt in test_types):
                raise ValueError("Please select at least one test type")
            
            # Save backup
            backup_dir = os.path.join(os.path.dirname(__file__), '..', '..', 'backups')
            os.makedirs(backup_dir, exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_file = os.path.join(backup_dir, f'test_input_backup_{timestamp}.csv')
            
            try:
                with open(backup_file, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(['field', 'value'])
                    for field, value in test_data.items():
                        writer.writerow([field, value])
            except Exception as e:
                print(f"Warning: Could not save backup file: {str(e)}")
            
            # Call callback with the data
            if self.callback_on_save:
                self.callback_on_save(test_data)
            
            # Close the window
            parent = self.parent
            while parent is not None:
                if isinstance(parent, Popup):
                    parent.dismiss()
                    break
                parent = parent.parent
            
        except Exception as e:
            self.show_error(str(e))

    def cancel(self, instance):
        """Close the window without saving"""
        parent = self.parent
        while parent is not None:
            if isinstance(parent, Popup):
                parent.dismiss()
                break
            parent = parent.parent

    def show_error(self, message):
        """Show error popup"""
        popup = Popup(
            title='Error',
            content=Label(text=message),
            size_hint=(None, None),
            size=(400, 200)
        )
        popup.open()