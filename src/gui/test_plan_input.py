from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.spinner import Spinner
from kivy.uix.checkbox import CheckBox
from kivy.uix.popup import Popup

from data_processor import RoomProperties, TestType

class TestPlanInputWindow(BoxLayout):
    def __init__(self, callback_on_save=None, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = 10
        self.spacing = 5
        self.callback_on_save = callback_on_save
        
        # Create scrollable form
        self.form = GridLayout(cols=2, spacing=5, size_hint_y=None)
        self.form.bind(minimum_height=self.form.setter('height'))
        
        # Initialize input fields
        self.input_fields = {}
        self._create_form_fields()
        
        # Add buttons
        self._create_buttons()
        
    def _create_form_fields(self):
        """Create input fields based on RoomProperties"""
        fields = [
            ('site_name', 'Site Name', TextInput),
            ('client_name', 'Client Name', TextInput),
            ('source_room', 'Source Room', TextInput),
            ('receive_room', 'Receive Room', TextInput),
            ('test_date', 'Test Date', TextInput),
            ('report_date', 'Report Date', TextInput),
            ('project_name', 'Project Name', TextInput),
            ('test_label', 'Test Label', TextInput),
            ('source_vol', 'Source Volume (cu.ft)', TextInput),
            ('receive_vol', 'Receive Volume (cu.ft)', TextInput),
            ('partition_area', 'Partition Area (sq.ft)', TextInput),
            ('partition_dim', 'Partition Dimensions', TextInput),
            ('source_room_finish', 'Source Room Finish', TextInput),
            ('source_room_name', 'Source Room Name', TextInput),
            ('receive_room_finish', 'Receive Room Finish', TextInput),
            ('receive_room_name', 'Receive Room Name', TextInput),
            ('srs_floor', 'Source Floor Description', TextInput),
            ('srs_walls', 'Source Walls Description', TextInput),
            ('srs_ceiling', 'Source Ceiling Description', TextInput),
            ('rec_floor', 'Receive Floor Description', TextInput),
            ('rec_walls', 'Receive Walls Description', TextInput),
            ('rec_ceiling', 'Receive Ceiling Description', TextInput),
            ('tested_assembly', 'Tested Assembly', TextInput),
            ('expected_performance', 'Expected Performance', TextInput),
            ('test_assembly_type', 'Test Assembly Type', TextInput),
            ('annex_2_used', 'Annex 2 Used', CheckBox)
        ]
        
        for field_name, label_text, input_type in fields:
            # Add label
            self.form.add_widget(Label(text=label_text))
            
            # Create and add input widget
            if input_type == CheckBox:
                input_widget = input_type()
            else:
                input_widget = input_type(multiline=False)
            
            self.input_fields[field_name] = input_widget
            self.form.add_widget(input_widget)
            
        # Add test type selector
        self.form.add_widget(Label(text='Test Type'))
        self.test_type_spinner = Spinner(
            text='AIIC',
            values=[t.value for t in TestType],
            size_hint=(None, None),
            size=(100, 44),
            pos_hint={'center_x': .5, 'center_y': .5})
        self.form.add_widget(self.test_type_spinner)
        
        self.add_widget(self.form)

    def _create_buttons(self):
        """Create save and cancel buttons"""
        button_layout = BoxLayout(
            orientation='horizontal',
            size_hint_y=None,
            height=50,
            spacing=10,
            padding=10
        )
        
        save_btn = Button(text='Save', size_hint_x=0.5)
        save_btn.bind(on_press=self.save_data)
        
        cancel_btn = Button(text='Cancel', size_hint_x=0.5)
        cancel_btn.bind(on_press=self.cancel)
        
        button_layout.add_widget(save_btn)
        button_layout.add_widget(cancel_btn)
        self.add_widget(button_layout)

    def save_data(self, instance):
        """Save form data to RoomProperties"""
        try:
            # Collect data from form
            data = {}
            for field_name, input_widget in self.input_fields.items():
                if isinstance(input_widget, CheckBox):
                    data[field_name] = input_widget.active
                elif isinstance(input_widget, TextInput):
                    # Convert numeric fields
                    if field_name in ['source_vol', 'receive_vol', 'partition_area']:
                        try:
                            data[field_name] = float(input_widget.text)
                        except ValueError:
                            raise ValueError(f"Invalid numeric value for {field_name}")
                    else:
                        data[field_name] = input_widget.text

            # Create RoomProperties instance
            room_props = RoomProperties.from_dict(data)
            
            # Get selected test type
            test_type = TestType(self.test_type_spinner.text)
            
            # Call callback with data if provided
            if self.callback_on_save:
                self.callback_on_save(room_props, test_type)
            
            # Close window
            self.parent.parent.dismiss()
            
        except Exception as e:
            self._show_error(f"Error saving data: {str(e)}")

    def cancel(self, instance):
        """Cancel and close window"""
        self.parent.parent.dismiss()

    def _show_error(self, message):
        """Show error popup"""
        popup = Popup(
            title='Error',
            content=Label(text=message),
            size_hint=(None, None),
            size=(400, 200)
        )
        popup.open()

    def load_data(self, room_properties: RoomProperties, test_type: TestType):
        """Load existing data into form"""
        data = vars(room_properties)
        for field_name, value in data.items():
            if field_name in self.input_fields:
                input_widget = self.input_fields[field_name]
                if isinstance(input_widget, CheckBox):
                    input_widget.active = value
                else:
                    input_widget.text = str(value)
        
        self.test_type_spinner.text = test_type.value 