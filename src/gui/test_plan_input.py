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
        
        # Create room property fields
        self._create_room_property_fields()
        
        # Create test type checkboxes
        self._create_test_type_section()
        
        # Add form to scroll view
        scroll.add_widget(self.form)
        self.add_widget(scroll)
        
        # Create buttons
        self._create_buttons()

    def _create_room_property_fields(self):
        """Create input fields for all RoomProperties"""
        # Define fields with their display names
        fields = [
            ('site_name', 'Site Name'),
            ('client_name', 'Client Name'),
            ('source_room', 'Source Room'),
            ('receive_room', 'Receive Room'),
            ('test_date', 'Test Date'),
            ('report_date', 'Report Date'),
            ('project_name', 'Project Name'),
            ('test_label', 'Test Label'),
            ('source_vol', 'Source Room Volume'),
            ('receive_vol', 'Receive Room Volume'),
            ('partition_area', 'Partition Area'),
            ('partition_dim', 'Partition Dimensions'),
            ('source_room_finish', 'Source Room Finish'),
            ('source_room_name', 'Source Room Name'),
            ('receive_room_finish', 'Receive Room Finish'),
            ('receive_room_name', 'Receive Room Name'),
            ('srs_floor', 'Source Room Floor'),
            ('srs_walls', 'Source Room Walls'),
            ('srs_ceiling', 'Source Room Ceiling'),
            ('rec_floor', 'Receive Room Floor'),
            ('rec_walls', 'Receive Room Walls'),
            ('rec_ceiling', 'Receive Room Ceiling'),
            ('tested_assembly', 'Tested Assembly'),
            ('expected_performance', 'Expected Performance'),
            ('test_assembly_type', 'Test Assembly Type')
        ]
        
        # Add fields to form
        for field_id, display_name in fields:
            # Add label
            self.form.add_widget(Label(
                text=display_name,
                size_hint_y=None,
                height=40
            ))
            
            # Add input
            input_widget = TextInput(
                multiline=False,
                size_hint_y=None,
                height=40
            )
            self.input_fields[field_id] = input_widget
            self.form.add_widget(input_widget)
        
        # Add Annex 2 checkbox separately
        self.form.add_widget(Label(
            text='Annex 2 Used',
            size_hint_y=None,
            height=40
        ))
        annex_checkbox = CheckBox(size_hint_y=None, height=40)
        self.input_fields['annex_2_used'] = annex_checkbox
        self.form.add_widget(annex_checkbox)

    def _create_test_type_section(self):
        """Create test type checkbox section"""
        # Add a header for test types
        self.form.add_widget(Label(
            text='Test Types',
            size_hint_y=None,
            height=40,
            bold=True
        ))
        
        # Create layout for checkboxes
        test_type_layout = GridLayout(
            cols=2,
            size_hint_y=None,
            height=len(TestType) * 20
        )
        
        # Add checkbox for each test type
        self.test_type_checkboxes = {}
        for test_type in TestType:
            checkbox = CheckBox(size_hint_y=None, height=20)
            label = Label(
                text=test_type.value,
                size_hint_y=None,
                height=20
            )
            self.test_type_checkboxes[test_type] = checkbox
            test_type_layout.add_widget(checkbox)
            test_type_layout.add_widget(label)
            
        self.form.add_widget(test_type_layout)

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
            # Collect room properties
            room_data = {}
            for field_id, input_widget in self.input_fields.items():
                if isinstance(input_widget, CheckBox):
                    room_data[field_id] = input_widget.active
                else:
                    room_data[field_id] = input_widget.text
            
            # Create RoomProperties instance
            room_properties = RoomProperties.from_dict(room_data)
            
            # Get selected test types
            selected_test_types = [
                test_type for test_type, checkbox in self.test_type_checkboxes.items()
                if checkbox.active
            ]
            # Save input data to CSV as backup


            # Create backups directory if it doesn't exist
            backup_dir = os.path.join(os.path.dirname(__file__), '..', '..', 'backups')
            os.makedirs(backup_dir, exist_ok=True)

            # Generate filename with timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_file = os.path.join(backup_dir, f'test_plan_backup_{timestamp}.csv')

            # Prepare data for CSV
            csv_data = []
            
            # Add room properties
            for field_id, value in room_data.items():
                csv_data.append(['room_property', field_id, str(value)])
                
            # Add selected test types
            for test_type in selected_test_types:
                csv_data.append(['test_type', test_type.value, 'selected'])

            # Write to CSV file
            try:
                with open(backup_file, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(['data_type', 'field', 'value'])  # Header
                    writer.writerows(csv_data)
                print(f"Backup saved to {backup_file}")
            except Exception as e:
                print(f"Warning: Could not save backup file: {str(e)}")
            if not selected_test_types:
                raise ValueError("Please select at least one test type")
            
            # Call callback with the data
            if self.callback_on_save:
                for test_type in selected_test_types:
                    self.callback_on_save(room_properties, test_type)
            
            # Close the window
            self.parent.parent.dismiss()
            
        except Exception as e:
            # Show error popup
            self.show_error(str(e))

    def cancel(self, instance):
        """Close the window without saving"""
        self.parent.parent.dismiss()

    def show_error(self, message):
        """Show error popup"""
        from kivy.uix.popup import Popup
        popup = Popup(
            title='Error',
            content=Label(text=message),
            size_hint=(None, None),
            size=(400, 200)
        )
        popup.open()