import json
from datetime import datetime

class ConfigManager:
    def __init__(self, config_file='config.json'):
        self.config_file = config_file
        self.config = self.load_config()
        
    def load_config(self):
        try:
            with open(self.config_file, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            return self.create_default_config()
    
    def create_default_config(self):
        return {
            'equipment': {
                'Sound Level Meters': [],
                'Microphones': [],
                'Calibrators': [],
                'Speakers': [],
                'Other Equipment': []
            },
            'standards': {
                'ASTM E336': '',
                'ASTM E90': '',
                'ISO 717-1': ''
            },
            'report_text': {
                'header_text': '',
                'footer_text': '',
                'disclaimer': ''
            }
        }
    
    def save_equipment(self, category, equipment_data):
        """Save equipment data with validation"""
        if not self._validate_equipment_data(equipment_data):
            raise ValueError("Invalid equipment data")
            
        if category not in self.config['equipment']:
            self.config['equipment'][category] = []
            
        equipment_data['last_updated'] = datetime.now().isoformat()
        self.config['equipment'][category].append(equipment_data)
        self.save_config()
    
    def _validate_equipment_data(self, data):
        required_fields = ['model', 'serial']
        return all(field in data for field in required_fields)
    
    def save_config(self):
        with open(self.config_file, 'w') as f:
            json.dump(self.config, f, indent=4) 