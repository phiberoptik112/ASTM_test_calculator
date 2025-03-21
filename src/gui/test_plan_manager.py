from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.checkbox import CheckBox
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from data_processor import RoomProperties, TestType
from src.gui.test_plan_input import TestPlanInputWindow
import pandas as pd
import os
from datetime import datetime
import csv

class TestPlanManagerWindow(BoxLayout):
    def __init__(self, test_data_manager=None, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.spacing = 10
        self.padding = 10
        self.test_data_manager = test_data_manager
        
        # Create main sections
        self._create_header()
        self._create_content()
        self._create_footer()
        
        # Initialize data
        self.current_test_plan = None
        self.edited_rows = set()
        
    def _create_header(self):
        """Create header with action buttons"""
        header = BoxLayout(
            orientation='horizontal',
            size_hint_y=0.1,
            spacing=10,
            padding=5
        )
        
        # Create New Test Plan button
        new_btn = Button(
            text='Create New Test Plan',
            size_hint_x=0.3
        )
        new_btn.bind(on_press=self._create_new_test_plan)
        header.add_widget(new_btn)
        
        # Modify Current Test Plan button
        modify_btn = Button(
            text='Modify Current Test Plan',
            size_hint_x=0.3
        )
        modify_btn.bind(on_press=self._modify_current_test_plan)
        header.add_widget(modify_btn)
        
        # Save button
        save_btn = Button(
            text='Save Changes',
            size_hint_x=0.3
        )
        save_btn.bind(on_press=self._save_changes)
        header.add_widget(save_btn)
        
        self.add_widget(header)
        
    def _create_content(self):
        """Create main content area with tabs"""
        self.content_tabs = TabbedPanel(size_hint_y=0.8)
        
        # Test Plan Tab
        self.test_plan_tab = TabbedPanelItem(text='Test Plan')
        self.test_plan_tab.do_default_tab = True
        
        # Create scrollable grid for test plan
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
        self.content_tabs.add_widget(self.test_plan_tab)
        
        # Add Test Row Tab
        self.add_test_tab = TabbedPanelItem(text='Add/Edit Test')
        self.add_test_form = TestPlanInputWindow(callback_on_save=self._on_test_save)
        self.add_test_tab.add_widget(self.add_test_form)
        self.content_tabs.add_widget(self.add_test_tab)
        
        self.add_widget(self.content_tabs)
        
    def _create_footer(self):
        """Create footer with status information"""
        footer = BoxLayout(
            orientation='horizontal',
            size_hint_y=0.1,
            spacing=10,
            padding=5
        )
        
        self.status_label = Label(
            text='Status: Ready',
            size_hint_x=0.7
        )
        footer.add_widget(self.status_label)
        
        # Close button
        close_btn = Button(
            text='Close',
            size_hint_x=0.3
        )
        close_btn.bind(on_press=self._close_window)
        footer.add_widget(close_btn)
        
        self.add_widget(footer)
        
    def _create_new_test_plan(self, instance):
        """Initialize a new test plan"""
        try:
            # Create empty DataFrame with required columns
            columns = [
                'Test_Label', 'Client_Name', 'Site_Name', 'Source Room',
                'Receiving Room', 'Test Date', 'Report Date', 'Project Name',
                'source room vol', 'receive room vol', 'partition area',
                'partition dim', 'source room finish', 'receive room finish',
                'srs floor descrip.', 'srs walls descrip.', 'srs ceiling descrip.',
                'rec floor descrip.', 'rec walls descrip.', 'rec ceiling descrip.',
                'tested assembly', 'expected performance', 'Test assembly Type',
                'Annex 2 used?', 'AIIC', 'ASTC', 'NIC', 'DTC'
            ]
            
            self.current_test_plan = pd.DataFrame(columns=columns)
            self._update_test_plan_display()
            self.status_label.text = 'Status: New test plan created'
            
        except Exception as e:
            self._show_error(f"Error creating new test plan: {str(e)}")
            
    def _modify_current_test_plan(self, instance):
        """Load and display current test plan"""
        try:
            if self.test_data_manager and self.test_data_manager.test_plan is not None:
                self.current_test_plan = self.test_data_manager.test_plan.copy()
                self._update_test_plan_display()
                self.status_label.text = 'Status: Current test plan loaded'
            else:
                self._show_error("No current test plan available")
                
        except Exception as e:
            self._show_error(f"Error loading current test plan: {str(e)}")
            
    def _update_test_plan_display(self):
        """Update the test plan grid with current data"""
        self.test_plan_grid.clear_widgets()
        
        if self.current_test_plan is None or self.current_test_plan.empty:
            self.test_plan_grid.add_widget(Label(
                text='No test plan data available',
                size_hint_y=None,
                height=40
            ))
            return
            
        try:
            # Create header row
            header = GridLayout(
                cols=len(self.current_test_plan.columns),
                size_hint_y=None,
                height=40,
                spacing=2
            )
            
            # Add column headers
            for col in self.current_test_plan.columns:
                header.add_widget(Label(
                    text=str(col),
                    bold=True,
                    size_hint_y=None,
                    height=40
                ))
            
            self.test_plan_grid.add_widget(header)
            
            # Add data rows
            for idx, row in self.current_test_plan.iterrows():
                row_layout = GridLayout(
                    cols=len(row),
                    size_hint_y=None,
                    height=40,
                    spacing=2
                )
                
                for value in row:
                    cell = TextInput(
                        text=str(value),
                        multiline=False,
                        size_hint_y=None,
                        height=40
                    )
                    cell.bind(text=lambda instance, value, idx=idx: self._on_cell_edit(idx, value))
                    row_layout.add_widget(cell)
                
                self.test_plan_grid.add_widget(row_layout)
                
        except Exception as e:
            self._show_error(f"Error updating test plan display: {str(e)}")
            
    def _on_cell_edit(self, row_idx, value):
        """Handle cell edit events"""
        self.edited_rows.add(row_idx)
        self.status_label.text = 'Status: Changes pending save'
        
    def _on_test_save(self, room_properties: RoomProperties, test_type: TestType):
        """Handle saving a new or edited test"""
        try:
            # Convert RoomProperties to dictionary
            test_data = vars(room_properties)
            
            # Add test type flags
            test_data.update({
                'AIIC': 1 if test_type == TestType.AIIC else 0,
                'ASTC': 1 if test_type == TestType.ASTC else 0,
                'NIC': 1 if test_type == TestType.NIC else 0,
                'DTC': 1 if test_type == TestType.DTC else 0
            })
            
            # Create or update DataFrame
            if self.current_test_plan is None:
                self.current_test_plan = pd.DataFrame([test_data])
            else:
                # Check if test label exists
                existing_idx = self.current_test_plan[self.current_test_plan['Test_Label'] == test_data['test_label']].index
                if len(existing_idx) > 0:
                    # Update existing row
                    self.current_test_plan.loc[existing_idx[0]] = test_data
                else:
                    # Add new row
                    self.current_test_plan = pd.concat([self.current_test_plan, pd.DataFrame([test_data])], ignore_index=True)
            
            self._update_test_plan_display()
            self.status_label.text = 'Status: Test saved successfully'
            
        except Exception as e:
            self._show_error(f"Error saving test: {str(e)}")
            
    def _save_changes(self, instance):
        """Save changes to test plan"""
        try:
            if self.current_test_plan is None or self.current_test_plan.empty:
                self._show_error("No test plan data to save")
                return
                
            # Create backups directory if it doesn't exist
            backup_dir = os.path.join(os.path.dirname(__file__), '..', '..', 'backups')
            os.makedirs(backup_dir, exist_ok=True)
            
            # Generate filename with timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_file = os.path.join(backup_dir, f'test_plan_backup_{timestamp}.csv')
            
            # Save to CSV
            self.current_test_plan.to_csv(backup_file, index=False)
            
            # Update test_data_manager if available
            if self.test_data_manager:
                self.test_data_manager.test_plan = self.current_test_plan
                self.test_data_manager.test_plan_path = backup_file
            
            self.edited_rows.clear()
            self.status_label.text = 'Status: Changes saved successfully'
            
        except Exception as e:
            self._show_error(f"Error saving changes: {str(e)}")
            
    def _close_window(self, instance):
        """Close the window"""
        if self.edited_rows:
            # Show confirmation popup
            content = BoxLayout(orientation='vertical', spacing=10, padding=10)
            content.add_widget(Label(
                text='You have unsaved changes.\nDo you want to save before closing?',
                size_hint_y=None,
                height=60
            ))
            
            buttons = BoxLayout(
                orientation='horizontal',
                spacing=10,
                size_hint_y=None,
                height=40
            )
            
            save_btn = Button(text='Save')
            save_btn.bind(on_press=lambda x: self._save_and_close())
            buttons.add_widget(save_btn)
            
            close_btn = Button(text='Close Without Saving')
            close_btn.bind(on_press=lambda x: self._force_close())
            buttons.add_widget(close_btn)
            
            content.add_widget(buttons)
            
            popup = Popup(
                title='Unsaved Changes',
                content=content,
                size_hint=(None, None),
                size=(400, 200)
            )
            popup.open()
        else:
            self._force_close()
            
    def _save_and_close(self):
        """Save changes and close window"""
        self._save_changes(None)
        self._force_close()
        
    def _force_close(self):
        """Close window without saving"""
        # Find the popup window by traversing up the widget tree
        parent = self.parent
        while parent is not None:
            if isinstance(parent, Popup):
                parent.dismiss()
                break
            parent = parent.parent
        
    def _show_error(self, message):
        """Show error popup"""
        popup = Popup(
            title='Error',
            content=Label(text=message),
            size_hint=(None, None),
            size=(400, 200)
        )
        popup.open() 