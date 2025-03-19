import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt
from typing import Dict, List
from data_processor import TestType, SLMData

class TestOverviewWindow(tk.Toplevel):
    def __init__(self, parent, test_data_manager):
        super().__init__(parent)
        self.title("Test Overview")
        self.test_data_manager = test_data_manager
        
        # Configure window
        self.geometry("800x600")
        
        # Create main frame
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create table
        self.create_test_table()
        
        # Create plot button
        self.plot_button = ttk.Button(
            self.main_frame,
            text="Plot Raw Data",
            command=self.plot_selected_data
        )
        self.plot_button.pack(pady=10)

    def create_test_table(self):
        """Create table showing tests and their types"""
        # Create frame for table
        table_frame = ttk.Frame(self.main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create treeview
        columns = [t.value for t in TestType]
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        
        # Configure scrollbars
        y_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        
        # Configure grid
        self.tree.grid(row=0, column=0, sticky='nsew')
        y_scroll.grid(row=0, column=1, sticky='ns')
        x_scroll.grid(row=1, column=0, sticky='ew')
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        
        # Set up columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='center')
        
        # Add data
        self.populate_table()

    def populate_table(self):
        """Populate table with test data"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Get test data
        test_collection = self.test_data_manager.test_data_collection
        
        # Add each test
        for test_label in sorted(test_collection.keys()):
            test_types = test_collection[test_label].keys()
            
            # Create values list with checkmarks for present test types
            values = []
            for test_type in TestType:
                values.append('✓' if test_type in test_types else '')
            
            # Insert row
            self.tree.insert('', 'end', values=values, text=test_label)

    def plot_selected_data(self):
        """Plot raw data for selected test"""
        # Get selected item
        selection = self.tree.selection()
        if not selection:
            return
        
        test_label = self.tree.item(selection[0])['text']
        test_data = self.test_data_manager.test_data_collection[test_label]
        
        # Create figure
        plt.figure(figsize=(12, 6))
        
        # Plot data for each test type
        for test_type, data in test_data.items():
            if 'test_data' in data:
                self.plot_test_data(data['test_data'], test_type)
        
        # Configure plot
        plt.title(f'Raw Data for Test {test_label}')
        plt.xlabel('Frequency (Hz)')
        plt.ylabel('Level (dB)')
        plt.grid(True)
        plt.legend()
        plt.xscale('log')
        
        # Show plot
        plt.show()

    def plot_test_data(self, test_data: Dict, test_type: TestType):
        """Plot data for a specific test"""
        # Extract SLMData objects from test data
        slm_data_objects = [
            data for data in test_data.values() 
            if isinstance(data, SLMData)
        ]
        
        # Plot each SLMData object
        for slm_data in slm_data_objects:
            if hasattr(slm_data, 'frequency_bands') and hasattr(slm_data, 'overall_levels'):
                plt.plot(
                    slm_data.frequency_bands,
                    slm_data.overall_levels,
                    label=f'{test_type.value}'
                ) 