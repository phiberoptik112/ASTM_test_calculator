from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from kivy.uix.widget import Widget
import traceback
import numpy as np

from data_processor import (
    TestType,
    TestData
)
from src.core.test_data_manager import TestDataManager

class ResultsAnalysisDashboard(BoxLayout):
    def __init__(self, test_data_manager, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.test_data_manager = test_data_manager
        
        # Create results grid - increase columns to show more data
        self.results_grid = GridLayout(
            cols=5,  # Increased to accommodate more data
            size_hint_y=None,
            spacing=[5, 5],
            padding=[5, 5]
        )
        self.results_grid.bind(minimum_height=self.results_grid.setter('height'))
        
        # Add headers
        headers = ['Test Label', 'Test Type', 'Primary Result', 'Details', 'Status']
        for header in headers:
            self.results_grid.add_widget(Label(
                text=header,
                size_hint_y=None,
                height=40,
                bold=True
            ))
        
        # Add results in a scroll view
        scroll = ScrollView(size_hint=(1, 0.8))
        scroll.add_widget(self.results_grid)
        self.add_widget(scroll)
        
        # Add refresh button
        refresh_btn = Button(
            text='Refresh Results',
            size_hint_y=0.1
        )
        refresh_btn.bind(on_press=self.refresh_results)
        self.add_widget(refresh_btn)
        
        # Initial load
        self.refresh_results()
    
    def refresh_results(self, *args):
        """Refresh the results display"""
        # First debug the test data manager state
        self._debug_test_data_manager()
        
        # Clear existing results (except headers)
        while len(self.results_grid.children) > 5:
            self.results_grid.remove_widget(self.results_grid.children[0])
        
        # Try multiple methods to get test data
        test_collection = None
        
        # Method 0: Try direct_test_collection property first (set by MainWindow)
        if hasattr(self, 'direct_test_collection') and self.direct_test_collection:
            print(f"Using direct_test_collection property with {len(self.direct_test_collection)} items")
            test_collection = self.direct_test_collection
            
        # If that fails, try other methods
        if not test_collection:
            # Method 1: Try get_test_collection() first
            try:
                print("Trying standard method: get_test_collection()...")
                test_collection = self.test_data_manager.get_test_collection()
                print(f"Test collection contains {len(test_collection)} tests")
            except ValueError as e:
                print(f"Standard method failed: {str(e)}")
                
                # Method 2: Try accessing test_data_collection directly 
                try:
                    print("Trying direct dictionary access...")
                    if hasattr(self.test_data_manager, 'test_data_collection'):
                        test_collection = self.test_data_manager.test_data_collection
                        if test_collection:
                            print(f"Found test data via direct access: {len(test_collection)} tests")
                        else:
                            print("test_data_collection exists but is empty")
                except Exception as e2:
                    print(f"Direct access failed: {str(e2)}")
                    
                # Method 3: Try via __dict__
                if not test_collection:
                    try:
                        print("Trying access via __dict__...")
                        if hasattr(self.test_data_manager, '__dict__'):
                            tdm_dict = self.test_data_manager.__dict__
                            if 'test_data_collection' in tdm_dict:
                                test_collection = tdm_dict['test_data_collection']
                                print(f"Found test data via __dict__: {len(test_collection)} tests")
                    except Exception as e3:
                        print(f"__dict__ access failed: {str(e3)}")
                        
                # Method 4: Try accessing via getattr
                if not test_collection:
                    try:
                        print("Trying access via getattr...")
                        test_collection = getattr(self.test_data_manager, 'test_data_collection', None)
                        if test_collection:
                            print(f"Found test data via getattr: {len(test_collection)} tests")
                    except Exception as e4:
                        print(f"getattr access failed: {str(e4)}")
                        
                # Method 5: Directly get data from MainWindow
                if not test_collection and hasattr(self, 'get_parent_window'):
                    try:
                        print("Trying access via parent window...")
                        parent = self.get_parent_window()
                        if parent and hasattr(parent, 'test_data_manager'):
                            parent_tdm = parent.test_data_manager
                            if hasattr(parent_tdm, 'test_data_collection'):
                                test_collection = parent_tdm.test_data_collection
                                print(f"Found test data via parent: {len(test_collection)} tests")
                    except Exception as e5:
                        print(f"Parent access failed: {str(e5)}")
        
        # If still no data, show message and return
        if not test_collection:
            print("No test data available after all attempts")
            for _ in range(5):  # Add empty row
                self.results_grid.add_widget(Label(
                    text="No test data loaded yet",
                    size_hint_y=None,
                    height=40
                ))
            return
        
        # If test_collection is empty, show message
        if len(test_collection) == 0:
            print("Test collection is empty")
            for _ in range(5):  # Add empty row
                self.results_grid.add_widget(Label(
                    text="No test data available",
                    size_hint_y=None,
                    height=40
                ))
            return
            
        # We found data! Log what we found
        print(f"SUCCESS: Found test collection with {len(test_collection)} tests")
        for label, data in test_collection.items():
            print(f"  Test '{label}' has types: {list(data.keys())}")
        
        # Process each test
        print(f"Processing {len(test_collection)} tests:")
        for test_label, test_data in test_collection.items():
            print(f"\nProcessing test: {test_label}")
            print(f"Test types: {list(test_data.keys())}")
            
            for test_type, data in test_data.items():
                print(f"Processing {test_type} data")
                
                # Verify test data structure
                if 'test_data' not in data:
                    print(f"Warning: 'test_data' missing for {test_label} - {test_type}")
                    continue
                
                test_obj = data['test_data']
                has_calculated_values = hasattr(test_obj, 'calculated_values')
                
                if has_calculated_values:
                    print(f"Calculated values keys: {test_obj.calculated_values.keys()}")
                else:
                    print(f"No calculated values for {test_label} - {test_type}")
                
                # Add test label
                self.results_grid.add_widget(Label(
                    text=test_label,
                    size_hint_y=None,
                    height=40
                ))
                
                # Add test type
                self.results_grid.add_widget(Label(
                    text=test_type.value,
                    size_hint_y=None,
                    height=40
                ))
                
                # Add primary result value
                primary_result = self._get_primary_result(test_type, test_obj)
                self.results_grid.add_widget(Label(
                    text=primary_result,
                    size_hint_y=None,
                    height=40
                ))
                
                # Add view details button
                details_btn = Button(
                    text="View Details",
                    size_hint_y=None,
                    height=40
                )
                details_btn.bind(on_press=lambda btn, label=test_label, t_type=test_type: 
                                self._show_details(label, t_type))
                self.results_grid.add_widget(details_btn)
                
                # Add status
                status_text = 'Complete' if has_calculated_values else 'Incomplete'
                self.results_grid.add_widget(Label(
                    text=status_text,
                    size_hint_y=None,
                    height=40
                ))
    
    def _get_primary_result(self, test_type, test_data):
        """Get the primary result value for a test"""
        try:
            if not hasattr(test_data, 'calculated_values'):
                return "N/A"
                
            calculated_values = test_data.calculated_values
            print(f"Getting primary result for {test_type}")
            
            if test_type == TestType.AIIC:
                # Look for AIIC values
                if 'AIIC_contour_result' in calculated_values:
                    print(f"Found AIIC result: {calculated_values['AIIC_contour_result']}")
                    return f"AIIC {calculated_values['AIIC_contour_result']}"
                elif 'AIIC_contour_val' in calculated_values:
                    print(f"Found AIIC val: {calculated_values['AIIC_contour_val']}")
                    return f"AIIC {calculated_values['AIIC_contour_val']}"
            
            elif test_type == TestType.ASTC:
                # Look for ASTC values
                if 'ASTC_final_val' in calculated_values:
                    print(f"Found ASTC result: {calculated_values['ASTC_final_val']}")
                    return f"ASTC {calculated_values['ASTC_final_val']}"
            
            elif test_type == TestType.NIC:
                # Look for NIC values
                if 'NIC_final_val' in calculated_values:
                    print(f"Found NIC result: {calculated_values['NIC_final_val']}")
                    return f"NIC {calculated_values['NIC_final_val']}"
            
            elif test_type == TestType.DTC:
                # Look for DTC values
                if 'DTC_value' in calculated_values:
                    print(f"Found DTC result: {calculated_values['DTC_value']}")
                    return f"DTC {calculated_values['DTC_value']}"
            
            # If no specific value found, look for common patterns
            for key in calculated_values.keys():
                if 'final_val' in key.lower() or 'value' in key.lower() or 'result' in key.lower():
                    print(f"Found alternative result in {key}: {calculated_values[key]}")
                    return f"{key}: {calculated_values[key]}"
            
            # If still no result, just indicate data exists
            print("No primary result found, showing 'Calculated'")
            return "Calculated"
        except Exception as e:
            print(f"Error getting primary result: {str(e)}")
            traceback.print_exc()
            return "Error"
    
    def _show_details(self, test_label, test_type):
        """Show detailed test results in a popup"""
        try:
            print(f"Showing details for {test_label} - {test_type}")
            
            # Try direct_test_collection first (set by MainWindow)
            test_data = None
            if hasattr(self, 'direct_test_collection') and self.direct_test_collection:
                if test_label in self.direct_test_collection:
                    if test_type in self.direct_test_collection[test_label]:
                        test_data = self.direct_test_collection[test_label][test_type]
                        print(f"Retrieved test data via direct_test_collection")
            
            # If direct access fails, try standard methods
            if not test_data:
                # Try to get test data using standard method
                try:
                    test_data = self.test_data_manager.get_test_data(test_label, test_type)
                except Exception as e:
                    print(f"Standard get_test_data failed: {str(e)}")
                
                # If that fails, try direct access
                if not test_data and hasattr(self.test_data_manager, 'test_data_collection'):
                    try:
                        if test_label in self.test_data_manager.test_data_collection:
                            if test_type in self.test_data_manager.test_data_collection[test_label]:
                                test_data = self.test_data_manager.test_data_collection[test_label][test_type]
                                print(f"Retrieved test data via direct access")
                    except Exception as e:
                        print(f"Direct access also failed: {str(e)}")
            
            # Check if we have valid test data
            if not test_data or 'test_data' not in test_data:
                print(f"No test data found for {test_label} - {test_type}")
                return
            
            test_obj = test_data['test_data']
            if not hasattr(test_obj, 'calculated_values'):
                print(f"No calculated values for {test_label} - {test_type}")
                popup = Popup(
                    title=f"{test_label} - {test_type.value}",
                    content=Label(text="No calculated values available"),
                    size_hint=(0.8, 0.4)
                )
                popup.open()
                return
            
            print(f"Displaying {len(test_obj.calculated_values)} calculated values")
            # Create content layout
            content = BoxLayout(orientation='vertical', padding=10, spacing=5)
            
            # Add header
            content.add_widget(Label(
                text=f"{test_label} - {test_type.value} Results",
                size_hint_y=None,
                height=40,
                font_size=16,
                bold=True
            ))
            
            # Create scrollable grid for values
            scroll = ScrollView(size_hint=(1, 0.85))
            values_grid = GridLayout(
                cols=2, 
                spacing=5,
                size_hint_y=None
            )
            values_grid.bind(minimum_height=values_grid.setter('height'))
            
            # First add the primary result at the top if it exists
            primary_key = self._get_primary_key_for_test_type(test_type)
            if primary_key and primary_key in test_obj.calculated_values:
                # Create section header
                values_grid.add_widget(Label(
                    text="Primary Result:",
                    size_hint_y=None,
                    height=40,
                    bold=True,
                    halign='center',
                    valign='middle',
                    text_size=(None, 40)
                ))
                values_grid.add_widget(Label(
                    text="",
                    size_hint_y=None,
                    height=40
                ))
                
                # Add primary value
                values_grid.add_widget(Label(
                    text=f"{primary_key}:",
                    size_hint_y=None,
                    height=40,
                    halign='right',
                    valign='middle',
                    text_size=(None, 40),
                    bold=True
                ))
                
                primary_value = test_obj.calculated_values[primary_key]
                values_grid.add_widget(Label(
                    text=self._format_value_for_display(primary_value),
                    size_hint_y=None,
                    height=40,
                    halign='left',
                    valign='middle',
                    text_size=(None, 40),
                    bold=True
                ))
                
                # Add separator
                values_grid.add_widget(Widget(
                    size_hint_y=None,
                    height=2
                ))
                values_grid.add_widget(Widget(
                    size_hint_y=None,
                    height=2
                ))
            
            # Add section header for other values
            values_grid.add_widget(Label(
                text="All Calculated Values:",
                size_hint_y=None,
                height=40,
                bold=True,
                halign='center',
                valign='middle',
                text_size=(None, 40)
            ))
            values_grid.add_widget(Label(
                text="",
                size_hint_y=None,
                height=40
            ))
            
            # Add all calculated values sorted alphabetically
            for key in sorted(test_obj.calculated_values.keys()):
                # Skip primary key as we already displayed it
                if primary_key and key == primary_key:
                    continue
                
                # Add key
                values_grid.add_widget(Label(
                    text=f"{key}:",
                    size_hint_y=None,
                    height=30,
                    halign='right',
                    valign='middle',
                    text_size=(None, 30)
                ))
                
                # Format and add value
                value = test_obj.calculated_values[key]
                formatted_value = self._format_value_for_display(value)
                values_grid.add_widget(Label(
                    text=formatted_value,
                    size_hint_y=None,
                    height=30 * max(1, formatted_value.count('\n') + 1),
                    halign='left',
                    valign='middle',
                    text_size=(None, None)
                ))
            
            scroll.add_widget(values_grid)
            content.add_widget(scroll)
            
            # Add close button
            close_btn = Button(
                text="Close",
                size_hint_y=0.1
            )
            content.add_widget(close_btn)
            
            # Create and show popup
            popup = Popup(
                title=f"{test_label} - {test_type.value} Results",
                content=content,
                size_hint=(0.8, 0.8)
            )
            close_btn.bind(on_press=popup.dismiss)
            popup.open()
            
        except Exception as e:
            print(f"Error showing details: {str(e)}")
            traceback.print_exc()
    
    def _get_primary_key_for_test_type(self, test_type):
        """Get the primary result key for a given test type"""
        if test_type == TestType.AIIC:
            return 'AIIC_contour_result'
        elif test_type == TestType.ASTC:
            return 'ASTC_final_val'
        elif test_type == TestType.NIC:
            return 'NIC_final_val'
        elif test_type == TestType.DTC:
            return 'DTC_value'
        return None
    
    def _format_value_for_display(self, value):
        """Format a value for display in the details view"""
        if isinstance(value, np.ndarray):
            if value.size <= 5:
                return f"Array[{value.size}]: {value}"
            else:
                return f"Array[{value.size}]:\n{value[:3]}...\n{value[-2:]}"
        elif isinstance(value, list):
            if len(value) <= 5:
                return f"List[{len(value)}]: {value}"
            else:
                return f"List[{len(value)}]:\n{value[:3]}...\n{value[-2:]}"
        elif isinstance(value, float):
            return f"{value:.2f}"
        else:
            return str(value)
            
    def _debug_test_data_manager(self):
        """Debug method to inspect the test data manager state"""
        print("\n=== DEBUG TEST DATA MANAGER ===")
        
        # Check basic attributes
        print(f"test_data_manager type: {type(self.test_data_manager)}")
        print(f"Available attributes: {[attr for attr in dir(self.test_data_manager) if not attr.startswith('_')]}")
        
        # Check test_data_collection more thoroughly
        has_attr = hasattr(self.test_data_manager, 'test_data_collection')
        print(f"test_data_collection exists (hasattr): {has_attr}")
        
        # Try different ways to access the collection
        try:
            collection_dir = getattr(self.test_data_manager, 'test_data_collection', None)
            print(f"test_data_collection via getattr: {type(collection_dir)}, empty: {collection_dir == {}}, value: {collection_dir}")
        except Exception as e:
            print(f"Error accessing via getattr: {str(e)}")
            
        try:
            # Try direct dictionary access using dir()
            collection_direct = self.test_data_manager.test_data_collection
            print(f"test_data_collection direct access: {type(collection_direct)}, empty: {collection_direct == {}}")
            
            if collection_direct:
                print(f"test_data_collection has {len(collection_direct)} items")
                for label, data in collection_direct.items():
                    print(f"  Test '{label}' has types: {list(data.keys())}")
                    for test_type, test_data in data.items():
                        has_calculated = 'test_data' in test_data and hasattr(test_data['test_data'], 'calculated_values')
                        print(f"    {test_type}: has test_data={('test_data' in test_data)}, has calculated values={has_calculated}")
                        if has_calculated:
                            print(f"      Calculated values: {test_data['test_data'].calculated_values.keys()}")
            else:
                print("test_data_collection is empty")
                
        except Exception as e:
            print(f"Error accessing directly: {str(e)}")
            traceback.print_exc()
            
        # Try accessing using __dict__
        try:
            if hasattr(self.test_data_manager, '__dict__'):
                dict_attrs = self.test_data_manager.__dict__
                print(f"__dict__ keys: {list(dict_attrs.keys())}")
                if 'test_data_collection' in dict_attrs:
                    dict_collection = dict_attrs['test_data_collection']
                    print(f"test_data_collection via __dict__: {type(dict_collection)}, empty: {dict_collection == {}}")
                    if dict_collection:
                        print(f"Found {len(dict_collection)} items via __dict__")
        except Exception as e:
            print(f"Error accessing via __dict__: {str(e)}")
        
        # Check test_plan
        if hasattr(self.test_data_manager, 'test_plan'):
            test_plan = self.test_data_manager.test_plan
            print(f"test_plan exists: {test_plan is not None}")
            if test_plan is not None:
                print(f"test_plan type: {type(test_plan)}")
                print(f"test_plan shape: {test_plan.shape if hasattr(test_plan, 'shape') else 'N/A'}")
        else:
            print("test_plan does not exist")
        
        print("=== END DEBUG ===\n")