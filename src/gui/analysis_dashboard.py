class ResultsAnalysisDashboard(BoxLayout):
    def __init__(self, test_data_manager, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = 10
        self.spacing = 10
        self.test_data_manager = test_data_manager
        
        # Controls Layout at top
        self.add_widget(self._create_controls())
        
        # Results Display
        self.results_display = BoxLayout(size_hint_y=0.1)
        self.single_number_label = Label(text="")
        self.results_display.add_widget(self.single_number_label)
        self.add_widget(self.results_display)
        
        # Tabs for different analysis views
        self.tabs = TabbedPanel()
        
        # Frequency Response Tab
        self.freq_tab = TabbedPanelItem(text='Frequency Response')
        self.freq_plot_container = BoxLayout(orientation='vertical')
        self.freq_tab.add_widget(self.freq_plot_container)
        
        # Statistical Analysis Tab
        self.stats_tab = TabbedPanelItem(text='Statistical Analysis')
        self.stats_container = BoxLayout(orientation='vertical')
        self.stats_tab.add_widget(self.stats_container)
        
        # Comparison Tab
        self.compare_tab = TabbedPanelItem(text='Test Comparisons')
        self.compare_container = BoxLayout(orientation='vertical')
        self.compare_tab.add_widget(self.compare_container)
        
        # Add tabs
        self.tabs.add_widget(self.freq_tab)
        self.tabs.add_widget(self.stats_tab)
        self.tabs.add_widget(self.compare_tab)
        
        self.add_widget(self.tabs)
        
        # Initialize empty plots
        self.create_empty_plots()

    def _create_controls(self):
        """Create control panel with test type and test selection"""
        controls = BoxLayout(size_hint_y=0.1)
        
        # Test type selector
        self.test_type_spinner = Spinner(
            text='Select Test Type',
            values=[t.value for t in TestType],
            size_hint_x=0.2
        )
        self.test_type_spinner.bind(text=self.on_test_type_selected)
        
        # Test number selector
        self.test_spinner = Spinner(
            text='Select Test',
            values=[],
            size_hint_x=0.2
        )
        self.test_spinner.bind(text=self.on_test_selected)
        
        controls.add_widget(Label(text='Test Type:', size_hint_x=0.1))
        controls.add_widget(self.test_type_spinner)
        controls.add_widget(Label(text='Test:', size_hint_x=0.1))
        controls.add_widget(self.test_spinner)
        
        return controls

    def create_empty_plots(self):
        """Initialize empty plots for all tabs"""
        # Frequency Response Plot
        self.freq_fig, self.freq_ax = plt.subplots(figsize=(10, 6))
        self.freq_canvas = FigureCanvasKivyAgg(self.freq_fig)
        self.freq_plot_container.clear_widgets()
        self.freq_plot_container.add_widget(self.freq_canvas)
        
        # Statistics Plot
        self.stats_fig, self.stats_ax = plt.subplots(figsize=(10, 6))
        self.stats_canvas = FigureCanvasKivyAgg(self.stats_fig)
        self.stats_container.clear_widgets()
        self.stats_container.add_widget(self.stats_canvas)
        
        # Comparison Plot
        self.compare_fig, self.compare_ax = plt.subplots(figsize=(10, 6))
        self.compare_canvas = FigureCanvasKivyAgg(self.compare_fig)
        self.compare_container.clear_widgets()
        self.compare_container.add_widget(self.compare_canvas)

    def on_test_selected(self, instance, value):
        """Update all tabs when a test is selected"""
        if value != 'Select Test':
            self.current_test = value
            self.current_test_data = self.test_data_manager.get_test_data(
                self.test_type_spinner.text,
                value
            )
            self.update_single_number_result()
            self.update_frequency_plot()
            self.update_stats_view()
            self.update_comparison_view()

    def update_frequency_plot(self):
        """Update frequency response plot with all measurement data"""
        self.freq_ax.clear()
        
        if hasattr(self, 'current_test_data'):
            data = self.current_test_data
            
            # Plot all measurements
            self.freq_ax.plot(data.frequencies, data.source_levels, 
                            label='Source Room', color='blue', marker='o')
            self.freq_ax.plot(data.frequencies, data.receive_levels, 
                            label='Receive Room', color='green', marker='s')
            self.freq_ax.plot(data.frequencies, data.background_levels, 
                            label='Background', color='red', marker='^')
            self.freq_ax.plot(data.frequencies, data.final_results, 
                            label='Final Results', color='purple', marker='D')
            
            self._format_freq_plot()
            self.freq_canvas.draw()

    def update_stats_view(self):
        """Update statistical analysis view"""
        self.stats_ax.clear()
        
        if hasattr(self, 'current_test_data'):
            # Add statistical analysis plots (boxplots, distributions, etc.)
            data = self.current_test_data
            # ... statistical visualization code ...
            self.stats_canvas.draw()

    def update_comparison_view(self):
        """Update test comparison view"""
        self.compare_ax.clear()
        
        if hasattr(self, 'current_test_data'):
            # Add comparison with other tests of same type
            test_type = self.current_test_data.test_type
            all_tests = self.test_data_manager.get_tests_by_type(test_type)
            # ... comparison visualization code ...
            self.compare_canvas.draw() 