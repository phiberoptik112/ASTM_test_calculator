from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView

class ResultsAnalysisDashboard(BoxLayout):
    def __init__(self, test_data_manager, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.test_data_manager = test_data_manager
        
        # Create results grid
        self.results_grid = GridLayout(
            cols=4,  # Adjust based on your needs
            size_hint_y=None
        )
        self.results_grid.bind(minimum_height=self.results_grid.setter('height'))
        
        # Add headers
        headers = ['Test Label', 'Test Type', 'Result', 'Status']
        for header in headers:
            self.results_grid.add_widget(Label(
                text=header,
                size_hint_y=None,
                height=40
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
        # Clear existing results (except headers)
        while len(self.results_grid.children) > 4:
            self.results_grid.remove_widget(self.results_grid.children[0])
        
        # Add test results
        for test in self.test_data_manager.get_sorted_tests():
            # Add test label
            self.results_grid.add_widget(Label(
                text=test.room_properties.test_label,
                size_hint_y=None,
                height=40
            ))
            
            # Add test type
            self.results_grid.add_widget(Label(
                text=str(test.test_type),
                size_hint_y=None,
                height=40
            ))
            
            # Add result value
            result = self._get_test_result(test)
            self.results_grid.add_widget(Label(
                text=str(result),
                size_hint_y=None,
                height=40
            ))
            
            # Add status
            self.results_grid.add_widget(Label(
                text='Complete',
                size_hint_y=None,
                height=40
            ))
    
    def _get_test_result(self, test):
        """Get the primary result value for a test"""
        try:
            if hasattr(test, 'NIC_final_val'):
                return f"NIC {test.NIC_final_val}"
            elif hasattr(test, 'ASTC_val'):
                return f"ASTC {test.ASTC_val}"
            elif hasattr(test, 'AIIC_val'):
                return f"AIIC {test.AIIC_val}"
            else:
                return "N/A"
        except Exception:
            return "Error"