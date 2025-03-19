from kivy.app import App
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner

class ConfigEditorWindow(BoxLayout):
    def __init__(self, config_manager, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.config_manager = config_manager
        
        # Create tabbed panel for different config sections
        self.tabs = TabbedPanel()
        
        # Equipment Tab
        equipment_tab = TabbedPanelItem(text='Equipment')
        equipment_tab.add_widget(EquipmentEditor(self.config_manager))
        
        # Standards Tab
        standards_tab = TabbedPanelItem(text='Test Standards')
        standards_tab.add_widget(StandardsEditor(self.config_manager))
        
        # Report Text Tab
        report_tab = TabbedPanelItem(text='Report Text')
        report_tab.add_widget(ReportTextEditor(self.config_manager))
        
        # Add tabs
        self.tabs.add_widget(equipment_tab)
        self.tabs.add_widget(standards_tab)
        self.tabs.add_widget(report_tab)
        
        # Add save/cancel buttons
        button_layout = BoxLayout(size_hint_y=0.1)
        save_btn = Button(text='Save Changes', on_press=self.save_changes)
        cancel_btn = Button(text='Cancel', on_press=self.cancel_changes)
        button_layout.add_widget(save_btn)
        button_layout.add_widget(cancel_btn)
        
        self.add_widget(self.tabs)
        self.add_widget(button_layout) 