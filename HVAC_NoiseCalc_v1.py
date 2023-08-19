import tkinter as tk
from tkinter import ttk
import pandas as pd
import math
import numpy as np

# Main lookup table for duct elements
main_lookup_table = pd.DataFrame({
    'Element Type': ['Straight', 'Elbow', 'Branch', 'Silencer']
    # 'Lining': [True, False, False],
    # 'IL_63Hz': [2.5, 1.8, 2.0],
    # 'IL_125Hz': [3.0, 2.2, 2.5],
    # 'IL_250Hz': [2.8, 1.9, 2.2],
    # 'IL_500Hz': [2.2, 1.6, 1.8],
    # 'IL_1000Hz': [1.5, 1.2, 1.5],
    # 'IL_2000Hz': [1.0, 0.8, 1.0],
    # 'IL_4000Hz': [0.8, 0.6, 0.8],
    # 'IL_8000Hz': [0.5, 0.4, 0.5]
})


# Straight element lookup tables
lined_straight_lookup_table = pd.DataFrame({
    'Width': ['6x6', '12x12', '12x24', '24x24', '48x48', '72x72'],
    'IL_63Hz': [2.5, 3.0, 2.8, 2.2, 1.5, 1.0],
    'IL_125Hz': [3.0, 3.5, 3.3, 2.7, 1.8, 1.2],
    'IL_250Hz': [2.8, 3.3, 3.0, 2.5, 1.6, 1.0],
    'IL_500Hz': [2.2, 2.7, 2.5, 2.0, 1.2, 0.8],
    'IL_1000Hz': [1.5, 2.0, 1.8, 1.5, 0.8, 0.5],
    'IL_2000Hz': [1.0, 1.5, 1.3, 1.0, 0.5, 0.3],
    'IL_4000Hz': [0.8, 1.0, 0.9, 0.7, 0.3, 0.2],
    'IL_8000Hz': [0.5, 0.8, 0.7, 0.5, 0.2, 0.1]
})

unlined_straight_lookup_table = pd.DataFrame({
    'Width': ['6x6', '12x12', '12x24', '24x24', '48x48', '72x72'],
    'IL_63Hz': [1.8, 2.2, 1.9, 1.6, 1.2, 0.8],
    'IL_125Hz': [2.2, 2.6, 2.3, 1.9, 1.4, 1.0],
    'IL_250Hz': [1.9, 2.3, 2.0, 1.7, 1.3, 0.9],
    'IL_500Hz': [1.6, 1.9, 1.7, 1.4, 1.0, 0.6],
    'IL_1000Hz': [1.2, 1.5, 1.3, 1.0, 0.7, 0.4],
    'IL_2000Hz': [0.8, 1.1, 1.0, 0.8, 0.4, 0.2],
    'IL_4000Hz': [0.6, 0.8, 0.7, 0.5, 0.3, 0.1],
    'IL_8000Hz': [0.4, 0.6, 0.5, 0.4, 0.2, 0.1]
})

# Branch element lookup tables
lined_branch_lookup_table = pd.DataFrame({
    'Width': ['6x6', '12x12', '12x24', '24x24', '48x48', '72x72'],
    'IL_63Hz': [2.0, 2.3, 2.1, 1.7, 1.4, 1.1],
    'IL_125Hz': [2.5, 2.8, 2.6, 2.2, 1.9, 1.6],
    'IL_250Hz': [2.2, 2.5, 2.3, 1.9, 1.6, 1.3],
    'IL_500Hz': [1.8, 2.1, 1.9, 1.6, 1.3, 1.0],
    'IL_1000Hz': [1.5, 1.8, 1.6, 1.3, 1.0, 0.7],
    'IL_2000Hz': [1.0, 1.3, 1.1, 0.9, 0.6, 0.3],
    'IL_4000Hz': [0.8, 1.0, 0.9, 0.7, 0.4, 0.2],
    'IL_8000Hz': [0.5, 0.7, 0.6, 0.5, 0.3, 0.1]
})

unlined_branch_lookup_table = pd.DataFrame({
    'Width': ['6x6', '12x12', '12x24', '24x24', '48x48', '72x72'],
    'IL_63Hz': [2.0, 2.5, 2.2, 1.8, 1.5, 1.2],
    'IL_125Hz': [2.5, 3.0, 2.7, 2.3, 2.0, 1.7],
    'IL_250Hz': [2.2, 2.7, 2.4, 2.0, 1.7, 1.4],
    'IL_500Hz': [1.8, 2.3, 2.0, 1.6, 1.3, 1.0],
    'IL_1000Hz': [1.5, 2.0, 1.7, 1.3, 1.0, 0.7],
    'IL_2000Hz': [1.0, 1.5, 1.2, 0.8, 0.5, 0.2],
    'IL_4000Hz': [0.8, 1.0, 0.9, 0.7, 0.4, 0.1],
    'IL_8000Hz': [0.5, 0.8, 0.7, 0.5, 0.2, 0.1]
})

# Elbow element lookup tables
lined_elbow_lookup_table = pd.DataFrame({
    'Width': ['6x6', '12x12', '12x24', '24x24', '48x48', '72x72'],
    'IL_63Hz': [1.8, 2.2, 1.9, 1.6, 1.3, 1.0],
    'IL_125Hz': [2.2, 2.6, 2.3, 1.9, 1.6, 1.3],
    'IL_250Hz': [1.9, 2.3, 2.0, 1.7, 1.4, 1.1],
    'IL_500Hz': [1.6, 1.9, 1.7, 1.4, 1.1, 0.8],
    'IL_1000Hz': [1.2, 1.5, 1.3, 1.0, 0.7, 0.4],
    'IL_2000Hz': [0.8, 1.1, 1.0, 0.8, 0.5, 0.2],
    'IL_4000Hz': [0.6, 0.8, 0.7, 0.5, 0.3, 0.1],
    'IL_8000Hz': [0.4, 0.6, 0.5, 0.4, 0.2, 0.1]
})

unlined_elbow_lookup_table = pd.DataFrame({
    'Width': ['6x6', '12x12', '12x24', '24x24', '48x48', '72x72'],
    'IL_63Hz': [1.8, 2.2, 1.9, 1.6, 1.3, 1.0],
    'IL_125Hz': [2.2, 2.6, 2.3, 1.9, 1.6, 1.3],
    'IL_250Hz': [1.9, 2.3, 2.0, 1.7, 1.4, 1.1],
    'IL_500Hz': [1.6, 1.9, 1.7, 1.4, 1.1, 0.8],
    'IL_1000Hz': [1.2, 1.5, 1.3, 1.0, 0.7, 0.4],
    'IL_2000Hz': [0.8, 1.1, 1.0, 0.8, 0.5, 0.2],
    'IL_4000Hz': [0.6, 0.8, 0.7, 0.5, 0.3, 0.1],
    'IL_8000Hz': [0.4, 0.6, 0.5, 0.4, 0.2, 0.1]
})

# Silencer element lookup tables
lined_silencer_lookup_table = pd.DataFrame({
    'Width': ['6x6', '12x12', '12x24', '24x24', '48x48', '72x72'],
    'IL_per_length_63Hz': [0.010, 0.012, 0.014, 0.015, 0.020, 0.025],
    'IL_per_length_125Hz': [0.012, 0.014, 0.016, 0.018, 0.023, 0.028],
    'IL_per_length_250Hz': [0.014, 0.016, 0.018, 0.020, 0.025, 0.030],
    'IL_per_length_500Hz': [0.016, 0.018, 0.020, 0.022, 0.027, 0.032],
    'IL_per_length_1000Hz': [0.018, 0.020, 0.022, 0.024, 0.029, 0.034],
    'IL_per_length_2000Hz': [0.020, 0.022, 0.024, 0.026, 0.031, 0.036],
    'IL_per_length_4000Hz': [0.022, 0.024, 0.026, 0.028, 0.033, 0.038],
    'IL_per_length_8000Hz': [0.024, 0.026, 0.028, 0.030, 0.035, 0.040]
})

unlined_silencer_lookup_table = pd.DataFrame({
    'Width': ['6x6', '12x12', '12x24', '24x24', '48x48', '72x72'],
    'IL_per_length_63Hz': [0.010, 0.012, 0.014, 0.015, 0.020, 0.025],
    'IL_per_length_125Hz': [0.012, 0.014, 0.016, 0.018, 0.023, 0.028],
    'IL_per_length_250Hz': [0.014, 0.016, 0.018, 0.020, 0.025, 0.030],
    'IL_per_length_500Hz': [0.016, 0.018, 0.020, 0.022, 0.027, 0.032],
    'IL_per_length_1000Hz': [0.018, 0.020, 0.022, 0.024, 0.029, 0.034],
    'IL_per_length_2000Hz': [0.020, 0.022, 0.024, 0.026, 0.031, 0.036],
    'IL_per_length_4000Hz': [0.022, 0.024, 0.026, 0.028, 0.033, 0.038],
    'IL_per_length_8000Hz': [0.024, 0.026, 0.028, 0.030, 0.035, 0.040]
})

# Function to calculate insertion loss for a given duct element at desired frequencies
def calculate_insertion_loss(element_type, length, width, is_lined):
    insertion_loss = []
    for freq in desired_frequencies:
        try:
            if element_type == 'Straight':
                if is_lined:
                    lookup_table = lined_straight_lookup_table
                else:
                    lookup_table = unlined_straight_lookup_table
            elif element_type == 'Elbow':
                if is_lined:
                    lookup_table = lined_elbow_lookup_table
                else:
                    lookup_table = unlined_elbow_lookup_table
            elif element_type == 'Branch':
                if is_lined:
                    lookup_table = lined_branch_lookup_table
                else:
                    lookup_table = unlined_branch_lookup_table
            elif element_type == 'Silencer':
                if is_lined:
                    lookup_table = lined_silencer_lookup_table
                else:
                    lookup_table = unlined_silencer_lookup_table
            else:
                raise ValueError(f"Invalid element type: {element_type}")

            # Lookup insertion loss from the appropriate table based on the element type, lining, and width
            insertion_loss_value = lookup_table[(lookup_table['Width'] == width)][f'IL_{freq}Hz'].values[0]
            insertion_loss.append(insertion_loss_value)
        except IndexError:
            raise ValueError(f"No matching width {width} found in the lookup table for element type: {element_type}, is_lined: {is_lined}")

    total_insertion_loss = sum(insertion_loss)
    total_insertion_loss_per_length = total_insertion_loss / length
    return total_insertion_loss_per_length


# Desired frequencies for insertion loss calculations
desired_frequencies = [63, 125, 250, 500, 1000, 2000, 4000, 8000]


# Function to calculate sound pressure at the receiver
def calculate_sound_pressure(source_power, insertion_loss, receiver_distance, receiver_room_size):
    sound_pressure = []
    for power, loss in zip(source_power, insertion_loss):
        if loss is None:
            sound_pressure.append(None)
        else:
            if loss < 0:  # Silencer element reduces sound power
                power -= abs(loss)
            else:  # Other elements add or are neutral to transmitting sound power
                power += loss
            sound_pressure.append(10 * math.log10(power) + 94 + 20 * math.log10(receiver_distance) - receiver_room_size)
    for freq in desired_frequencies:
        sound_pressure_value = source_power[desired_frequencies.index(freq)] - (20 * math.log10(receiver_distance)) + (20 * math.log10(receiver_room_size))
        sound_pressure.append(sound_pressure_value)
    return sound_pressure

class PathFrame(tk.Frame):
    def __init__(self, master, element_type_options, element_table, total_insertion_loss_table):
        super().__init__(master)
        self.master = master
        self.element_table = element_table
        self.total_insertion_loss_table = total_insertion_loss_table

        self.element_type_var = tk.StringVar()
        self.element_type_var.set(element_type_options[0])
        label_element_type = tk.Label(self, text="Element Type:")
        label_element_type.pack()
        option_menu_element_type = tk.OptionMenu(self, self.element_type_var, *element_type_options)
        option_menu_element_type.pack()

        label_length = tk.Label(self, text="Length:")
        label_length.pack()
        self.entry_length = tk.Entry(self)
        self.entry_length.pack()

        label_width = tk.Label(self, text="Width:")
        label_width.pack()
        self.entry_width = tk.Entry(self)
        self.entry_width.pack()

        self.lined_var = tk.BooleanVar()
        self.lined_var.set(True)
        checkbox_lined = tk.Checkbutton(self, text="Lined", variable=self.lined_var)
        checkbox_lined.pack()

        button_add_element = tk.Button(self, text="Add Element", command=self.add_element)
        button_add_element.pack()

    def add_element(self):

        element_type = self.element_type_var.get()
        length = float(self.entry_length.get())
        width = self.entry_width.get()
        is_lined = self.lined_var.get()

        insertion_loss = calculate_insertion_loss(element_type, length, width, is_lined)

        self.element_table.add_element(element_type, insertion_loss)

        self.total_insertion_loss_table.update_total_insertion_loss()
    def get_elements(self):
        # gathering the elements info
        elements = []



class ElementTable(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master

        self.table = ttk.Treeview(self, columns=desired_frequencies, show="headings")
        self.table.pack(side=tk.LEFT, padx=5, pady=5)

        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.table.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.table.configure(yscrollcommand=scrollbar.set)

        self.table.heading("#0", text="Element Type")
        self.table.column("#0", width=100, anchor="center")
        for freq in desired_frequencies:
            self.table.heading(freq, text=freq)
            self.table.column(freq, width=80, anchor="center")

    def add_element(self, element_type, insertion_loss):
        insertion_loss_list = [insertion_loss] * len(desired_frequencies)
         # If it's None, set it to 0.0 to avoid concatenation error
        insertion_loss = [0.0 if il is None else il for il in insertion_loss_list]

        # Insert the element type and its insertion loss values into the table
        self.table.insert("", tk.END, text=element_type, values=insertion_loss)

class TotalInsertionLossTable(tk.Frame):
    def __init__(self, master, element_table):
        super().__init__(master)
        self.master = master
        self.element_table = element_table

        self.table = ttk.Treeview(self, columns=desired_frequencies, show="headings")
        self.table.pack(side=tk.LEFT, padx=5, pady=5)

        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.table.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.table.configure(yscrollcommand=scrollbar.set)

        self.table.heading("#0", text="Total Insertion Loss")
        self.table.column("#0", width=150, anchor="center")
        for freq in desired_frequencies:
            self.table.heading(freq, text=freq)
            self.table.column(freq, width=80, anchor="center")

    def update_total_insertion_loss(self):
        # self.table.delete(*self.table.get_children())  # Clear the existing data
        current_values = []
        for child in self.element_table.table.get_children():
            item = self.element_table.table.item(child)
            values = item['values']
            current_values.append(values)

        total_insertion_loss = [0.0] * len(desired_frequencies)
        for values in current_values:
            for i, il in enumerate(values):
                if il is not None:
                    total_insertion_loss[i] += float(il)

        self.table.delete(*self.table.get_children())  # Clear the existing data

        self.table.insert("", tk.END, text="Total", values=total_insertion_loss)

        # total_insertion_loss = [0] * len(desired_frequencies)

        # for child in self.element_table.table.get_children():
        #     item = self.element_table.table.item(child)
        #     values = item['values']
        #     values = [float(val) for val in values]  # Convert values to floats
        #     total_insertion_loss = [float(il) for il in total_insertion_loss]  # Convert total_insertion_loss to floats
        #     total_insertion_loss = [il + tl for il, tl in zip(values, total_insertion_loss)]



class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HVAC Duct Insertion Loss Calculator")
        self.geometry("1000x1000")

        self.entry_sources = []
        self.path_frames = []
        self.element_table = ElementTable(self)
        self.total_insertion_loss_table = TotalInsertionLossTable(self,self.element_table)

        self.create_widgets()

    def create_widgets(self):
        # Create entry fields for source sound power
        for freq in desired_frequencies:
            label_source = tk.Label(self, text=f"Source Sound Power {freq}Hz:")
            label_source.pack()
            entry_source = tk.Entry(self)
            entry_source.pack()
            self.entry_sources.append(entry_source)

        # Create entry fields for receiver distance and room size
        label_receiver_distance = tk.Label(self, text="Receiver Distance:")
        label_receiver_distance.pack()
        self.entry_receiver_distance = tk.Entry(self)
        self.entry_receiver_distance.pack()

        label_receiver_room_size = tk.Label(self, text="Receiver Room Size:")
        label_receiver_room_size.pack()
        self.entry_receiver_room_size = tk.Entry(self)
        self.entry_receiver_room_size.pack()
        # Create a frame for the paths and path elements list
        self.frame_paths = tk.Frame(self)
        self.frame_paths.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create a listbox to show the paths and elements
        self.listbox_paths = tk.Listbox(self.frame_paths, selectmode=tk.SINGLE)
        self.listbox_paths.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create a scrollbar for the listbox
        scrollbar = tk.Scrollbar(self.frame_paths, command=self.listbox_paths.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox_paths.configure(yscrollcommand=scrollbar.set)

        # Create a frame for the path elements details
        self.frame_path_elements = tk.Frame(self)
        self.frame_path_elements.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create a button to add a new path frame
        button_add_path = tk.Button(self.frame_path_elements, text="Add Path", command=self.add_path_frame)
        button_add_path.pack()

        # Create a button to calculate the results
        button_calculate = tk.Button(self.frame_path_elements, text="Calculate", command=self.calculate_results)
        button_calculate.pack()

        # # Create a frame for the paths
        # self.frame_paths = tk.Frame(self)
        # self.frame_paths.pack()

        # # Create a button to add a new path frame
        # button_add_path = tk.Button(self, text="Add Path", command=self.add_path_frame)
        # button_add_path.pack()

        # Create a frame for the element table
        self.frame_element_table = tk.Frame(self)
        self.frame_element_table.pack(fill=tk.BOTH, expand=True)

        self.element_table.pack(fill=tk.BOTH, expand=True)
        self.total_insertion_loss_table.pack(fill=tk.BOTH, expand=True)

    def add_path_frame(self):
        path_frame = PathFrame(self.frame_path_elements, main_lookup_table["Element Type"].tolist(), self.element_table, self.total_insertion_loss_table)
        path_frame.pack()
        self.path_frames.append(path_frame)
        self.listbox_paths.insert(tk.END, f"Path {len(self.path_frames)}")


    def calculate_results(self):
        # Get user inputs from the entry fields
        source_sound_power = [float(entry_source.get()) for entry_source in self.entry_sources]
        receiver_distance = float(self.entry_receiver_distance.get())
        receiver_room_size = float(self.entry_receiver_room_size.get())

        # Initialize total insertion loss list for the desired frequencies
        total_insertion_loss = [0] * len(desired_frequencies)

        for path_frame in self.path_frames:
            element_type = path_frame.element_type_var.get()
            length = float(path_frame.entry_length.get())
            width = path_frame.entry_width.get()
            is_lined = path_frame.lined_var.get()

            insertion_loss = calculate_insertion_loss(element_type, length, width, is_lined)

            self.element_table.add_element(element_type, insertion_loss)
            elements = path_frame["Elements"]
            for element in elements:
                insertion_loss = calculate_insertion_loss(element["Type"], element["Length"], element["Width"], element["Lined"])
                total_insertion_loss = [il + tl if il is not None else tl for il, tl in zip(insertion_loss, total_insertion_loss)]

            # insertion_loss = calculate_insertion_loss(element["Type"], element["Length"], element["Width"], element["Lined"])
            # total_insertion_loss = [il + tl if il is not None else tl for il, tl in zip(insertion_loss, total_insertion_loss)]

        self.total_insertion_loss_table.update_total_insertion_loss()

        # Calculate the combined insertion loss by subtracting the total insertion loss from the source sound power
        combined_insertion_loss = [sp - il for sp, il in zip(source_sound_power, total_insertion_loss)]

        # Calculate sound pressure at the receiver using the combined insertion loss, source sound power, distance, and room size
        receiver_sound_pressure = calculate_sound_pressure(combined_insertion_loss, receiver_distance, receiver_room_size)

        # Clear the entry sources
        for entry_source in self.entry_sources:
            entry_source.delete(0, tk.END)

        # Clear the entry fields for receiver distance and room size
        self.entry_receiver_distance.delete(0, tk.END)
        self.entry_receiver_room_size.delete(0, tk.END)

        # Output the calculated sound pressure values
        for freq, pressure in zip(desired_frequencies, receiver_sound_pressure):
            print(f"Frequency {freq}Hz: {pressure} dB")

app = Application()
app.mainloop()
