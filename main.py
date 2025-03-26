from datetime import datetime
import os
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.popup import Popup
from kivy.uix.slider import Slider
from kivy.core.window import Window
from kivy.uix.widget import Widget

from do import createMonthReportFile

# Set the window size
Window.size = (600, 400)  # Width, Height

# Function to find the first .xlsx file in the current folder
def find_first_xlsx_file():
    for file in os.listdir("."):
        if file.endswith(".xlsx"):
            return os.path.abspath(file)
    return None

# Placeholder for your processFile function
def processFile(file_path, year):
    # Simulate creating reports for all months in the year
    print(f"Processing file: {file_path}, Year: {year}")
    for month in range(1, 13):
        createMonthReportFile(file_path, year, month)
        
    return f"Reports for all months in {year} created successfully."

class ModernExcelApp(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(orientation="vertical", **kwargs)
        self.spacing = 30
        self.padding = 20

        # Find the first .xlsx file automatically
        self.selected_file = find_first_xlsx_file()

        # Title
        self.add_widget(Label(
            text="Excel Report Generator",
            font_size="24sp",
            halign="center",
            size_hint=(1, 0.2),
            bold=True
        ))

        # File Label
        file_text = f"File: {os.path.basename(self.selected_file)}" if self.selected_file else "File: No Excel file found!"
        self.file_label = Label(
            text=file_text,
            font_size="16sp",
            halign="center",
            size_hint=(1, 0.1)
        )
        self.add_widget(self.file_label)

        # Year Slider Section
        self.add_widget(Label(
            text="Select Year:",
            font_size="16sp",
            halign="center",
            size_hint=(1, 0.1)
        ))
        self.year_slider = Slider(
            min=datetime.now().year - 10,
            max=datetime.now().year + 10,
            value=datetime.now().year,
            step=1,
            size_hint=(1, 0.2),
        )
        self.add_widget(self.year_slider)

        # Year Slider Value Display
        self.year_label = Label(
            text=f"Year: {int(self.year_slider.value)}",
            font_size="16sp",
            halign="center",
            size_hint=(1, 0.1)
        )
        self.add_widget(self.year_label)

        # Link the slider to dynamically update the year label
        self.year_slider.bind(value=self.update_year_label)

        # Generate Reports Button
        self.generate_button = Button(
            text="Generate Reports",
            size_hint=(1, 0.2),
            background_color=(0.2, 0.6, 0.86, 1),
            font_size="18sp"
        )
        self.generate_button.bind(on_press=self.generate_reports)
        self.add_widget(self.generate_button)

        # Output Widget
        self.output_widget = BoxLayout(
            orientation="horizontal",
            size_hint=(1, 0.1),
            padding=10
        )
        self.output_label = Label(
            text="",
            font_size="16sp",
            halign="center",
        )
        self.add_widget(self.output_widget)
        self.output_widget.add_widget(self.output_label)

    def update_year_label(self, instance, value):
        self.year_label.text = f"Year: {int(value)}"

    def generate_reports(self, instance):
        # Check if a file is selected
        if not self.selected_file:
            self.display_message("Error: No Excel file found in the current folder.", success=False)
            return

        # Get the selected year
        year = int(self.year_slider.value)

        # Process the file and generate reports
        try:
            result = processFile(self.selected_file, year)
            self.display_message(result, success=True)
        except Exception as e:
            self.display_message(f"Error: {str(e)}", success=False)

    def display_message(self, message, success=True):
        # Update the message style and text dynamically
        self.output_label.text = message
        self.output_label.color = (0, 1, 0, 1) if success else (1, 0, 0, 1)  # Green for success, red for error
        self.output_label.bold = True

        popup_layout = BoxLayout(orientation="vertical", padding=20, spacing=10)
        popup_layout.add_widget(Label(
            text=message,
            halign="center",
            font_size="16sp",
            color=(0, 1, 0, 1) if success else (1, 0, 0, 1)
        ))
        close_button = Button(
            text="Close",
            size_hint=(1, 0.3),
            background_color=(0.8, 0.2, 0.2, 1)
        )
        popup = Popup(
            title="Success" if success else "Error",
            content=popup_layout,
            size_hint=(0.7, 0.4)
        )
        close_button.bind(on_press=popup.dismiss)
        popup_layout.add_widget(close_button)
        popup.open()

class ModernExcelAppUI(App):
    def build(self):
        return ModernExcelApp()

if __name__ == "__main__":
    ModernExcelAppUI().run()
