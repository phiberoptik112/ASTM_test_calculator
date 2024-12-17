class StandardsEditor(BoxLayout):
    def __init__(self, config_manager, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.config_manager = config_manager
        
        # Standards Selection
        standards_layout = BoxLayout(size_hint_y=0.1)
        standards_layout.add_widget(Label(text='Test Standard:'))
        self.standard_spinner = Spinner(
            text='Select Standard',
            values=('ASTM E336', 'ASTM E90', 'ISO 717-1')
        )
        standards_layout.add_widget(self.standard_spinner)
        
        # Standard Text Editor
        self.text_input = TextInput(
            multiline=True,
            hint_text='Enter standard text...'
        )
        
        self.add_widget(standards_layout)
        self.add_widget(self.text_input) 