class EquipmentEditor(GridLayout):
    def __init__(self, config_manager, **kwargs):
        super().__init__(**kwargs)
        self.cols = 2
        self.spacing = 10
        self.padding = 10
        self.config_manager = config_manager
        
        # Equipment Categories
        self.categories = {
            'Sound Level Meters': [],
            'Microphones': [],
            'Calibrators': [],
            'Speakers': [],
            'Other Equipment': []
        }
        
        self.load_equipment()
        
    def load_equipment(self):
        equipment_data = self.config_manager.get_equipment_data()
        
        for category in self.categories:
            # Category Header
            self.add_widget(Label(text=category, bold=True))
            
            # Add Equipment Button
            add_btn = Button(
                text=f'Add {category}',
                size_hint_x=0.2,
                on_press=lambda x, cat=category: self.add_equipment_item(cat)
            )
            self.add_widget(add_btn)
            
            # Equipment Items
            for item in equipment_data.get(category, []):
                self.add_equipment_row(category, item)
    
    def add_equipment_row(self, category, item_data):
        # Equipment details grid
        details_grid = GridLayout(cols=4)
        
        # Model
        details_grid.add_widget(TextInput(
            text=item_data.get('model', ''),
            hint_text='Model'
        ))
        
        # Serial Number
        details_grid.add_widget(TextInput(
            text=item_data.get('serial', ''),
            hint_text='Serial Number'
        ))
        
        # Calibration Date
        details_grid.add_widget(TextInput(
            text=item_data.get('cal_date', ''),
            hint_text='Cal Date (YYYY-MM-DD)'
        ))
        
        # Delete Button
        del_btn = Button(
            text='Delete',
            size_hint_x=0.1,
            on_press=lambda x: self.remove_equipment_row(details_grid)
        )
        details_grid.add_widget(del_btn)
        
        self.add_widget(details_grid) 