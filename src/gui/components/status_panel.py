"""
Status Panel Component

This component handles status reporting and error display for the ASTM Test Calculator.
Extracted from main_window.py to separate status UI logic.
"""

from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from typing import Optional


class StatusPanel(BoxLayout):
    """Panel for displaying application status and handling error display"""
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'horizontal'
        self.spacing = 5
        self.size_hint_y = None
        self.height = '40dp'
        
        # Initialize status widgets
        self.status_label = None
        
        self._create_status_widgets()
    
    def _create_status_widgets(self):
        """Create the status display widgets"""
        # Main status label
        self.status_label = Label(
            text='Status: Ready',
            size_hint_x=1.0,
            text_size=(None, None),
            halign='left',
            valign='middle'
        )
        self.status_label.bind(size=self.status_label.setter('text_size'))
        
        self.add_widget(self.status_label)
    
    def set_status(self, message: str, status_type: str = 'info'):
        """
        Set the status message
        
        Args:
            message: The status message to display
            status_type: Type of status ('info', 'success', 'error', 'warning')
        """
        if status_type == 'error':
            self.status_label.text = f'Status: ERROR - {message}'
        elif status_type == 'success':
            self.status_label.text = f'Status: SUCCESS - {message}'
        elif status_type == 'warning':
            self.status_label.text = f'Status: WARNING - {message}'
        else:
            self.status_label.text = f'Status: {message}'
    
    def show_error(self, message: str, title: str = 'Error'):
        """
        Display error message in popup with scrollable text
        
        Args:
            message: The error message to display
            title: The popup title
        """
        content_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        # Add scrollable text for long error messages
        scroll_view = ScrollView(size_hint=(1, 0.8))
        error_label = Label(
            text=message, 
            size_hint_y=None,
            text_size=(380, None),
            halign='left',
            valign='top'
        )
        error_label.bind(texture_size=error_label.setter('size'))
        scroll_view.add_widget(error_label)
        content_layout.add_widget(scroll_view)
        
        # Add close button
        close_button = Button(text="Close", size_hint=(1, 0.2))
        content_layout.add_widget(close_button)
        
        popup = Popup(
            title=title,
            content=content_layout,
            size_hint=(None, None),
            size=(400, 300)
        )
        
        close_button.bind(on_press=popup.dismiss)
        popup.open()
    
    def show_success(self, message: str, title: str = 'Success'):
        """
        Display success message in popup
        
        Args:
            message: The success message to display
            title: The popup title
        """
        content_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        # Success message
        success_label = Label(
            text=message,
            size_hint=(1, 0.8),
            text_size=(380, None),
            halign='center',
            valign='middle'
        )
        content_layout.add_widget(success_label)
        
        # Add close button
        close_button = Button(text="OK", size_hint=(1, 0.2))
        content_layout.add_widget(close_button)
        
        popup = Popup(
            title=title,
            content=content_layout,
            size_hint=(None, None),
            size=(300, 200)
        )
        
        close_button.bind(on_press=popup.dismiss)
        popup.open()
    
    def show_warning(self, message: str, title: str = 'Warning'):
        """
        Display warning message in popup
        
        Args:
            message: The warning message to display
            title: The popup title
        """
        content_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        # Warning message
        warning_label = Label(
            text=message,
            size_hint=(1, 0.8),
            text_size=(380, None),
            halign='center',
            valign='middle'
        )
        content_layout.add_widget(warning_label)
        
        # Add close button
        close_button = Button(text="OK", size_hint=(1, 0.2))
        content_layout.add_widget(close_button)
        
        popup = Popup(
            title=title,
            content=content_layout,
            size_hint=(None, None),
            size=(300, 200)
        )
        
        close_button.bind(on_press=popup.dismiss)
        popup.open()
    
    def show_confirmation(self, message: str, on_confirm: callable, 
                         on_cancel: callable = None, title: str = 'Confirm'):
        """
        Display confirmation dialog
        
        Args:
            message: The confirmation message
            on_confirm: Callback for confirm action
            on_cancel: Callback for cancel action (optional)
            title: The popup title
        """
        content_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        # Confirmation message
        confirm_label = Label(
            text=message,
            size_hint=(1, 0.6),
            text_size=(380, None),
            halign='center',
            valign='middle'
        )
        content_layout.add_widget(confirm_label)
        
        # Button layout
        button_layout = BoxLayout(size_hint=(1, 0.4), spacing=10)
        
        # Confirm button
        confirm_button = Button(text="Yes")
        def on_confirm_click(instance):
            popup.dismiss()
            if on_confirm:
                on_confirm()
        confirm_button.bind(on_press=on_confirm_click)
        button_layout.add_widget(confirm_button)
        
        # Cancel button
        cancel_button = Button(text="No")
        def on_cancel_click(instance):
            popup.dismiss()
            if on_cancel:
                on_cancel()
        cancel_button.bind(on_press=on_cancel_click)
        button_layout.add_widget(cancel_button)
        
        content_layout.add_widget(button_layout)
        
        popup = Popup(
            title=title,
            content=content_layout,
            size_hint=(None, None),
            size=(300, 200)
        )
        
        popup.open()
    
    def clear_status(self):
        """Clear the status message"""
        self.status_label.text = 'Status: Ready'
    
    def get_status_text(self) -> str:
        """Get the current status text"""
        return self.status_label.text