import tkinter
import customtkinter as ctk


class CustClr():
    """This class is for constructing the color tiles and storing all necessary information"""

    def __init__(self,
                 clr_name=None,
                 clr_hex=None,
                 name_entry=None,
                 hex_entry=None):

        self.clr_name = clr_name
        self.clr_hex = clr_hex
        self.name_entry = name_entry
        self.hex_entry = hex_entry

        self.color_fields = []
        self.hex_entry_fields = []
        self.switch_vars = []
        self.name_entry_fields = []

    def custclr_widget(self, clr_master, clr_table_row, clr_table_column):
        """The CustClr Widget holds all the parts of the color tiles"""

        clr_widget_base = ctk.CTkFrame(master=clr_master,
                                       width=70,
                                       height=83,
                                       corner_radius=0,
                                       border_width=0,
                                       fg_color="#181818")
        clr_widget_base.grid(
            row=clr_table_row, column=clr_table_column, padx=4, pady=5)

        self.name_entry = ctk.CTkEntry(master=clr_widget_base,
                                       placeholder_text="Color name",
                                       font=("Roboto", 10),
                                       width=80, height=20,
                                       corner_radius=0,
                                       border_width=0,
                                       fg_color="#141414",
                                       text_color="white")
        self.name_entry.pack(pady=2)
        self.name_entry_fields.append(self.name_entry)

        def switch_event():
            print("switch toggled, current value:", self.switch_var.get())

        self.switch_var = ctk.StringVar(value="off")
        self.switch = ctk.CTkSwitch(clr_widget_base,
                                    text=" ",
                                    height=10,
                                    width=20,
                                    switch_width=20,
                                    switch_height=10,
                                    command=switch_event,
                                    variable=self.switch_var,
                                    onvalue="on",
                                    progress_color="#76AE22",
                                    offvalue="off")
        self.switch.pack(anchor="center")
        self.switch_vars.append(self.switch_var)

        color_field = ctk.CTkFrame(master=clr_widget_base,
                                   width=20,
                                   height=20,
                                   corner_radius=4,
                                   border_color="#6D6D6D",
                                   fg_color="#1f1f1f",
                                   border_width=1)
        color_field.pack(anchor="center")
        self.color_fields.append(color_field)

        def limit_characters(entry, limit):
            value = entry.get()
            if len(value) > limit:
                entry.set(value[:limit])

        text_var = ctk.StringVar()
        self.hex_entry = ctk.CTkEntry(master=clr_widget_base,
                                      font=("Roboto", 10),
                                      textvariable=text_var,
                                      width=70,
                                      height=15,
                                      corner_radius=0,
                                      border_width=1,
                                      border_color="#6D6D6D",
                                      fg_color="#141414",
                                      text_color="white")
        self.hex_entry.pack(pady=5)
        self.hex_entry_fields.append(self.hex_entry)

        character_limit = 6
        text_var.trace_add(
            "write", lambda *args: limit_characters(text_var, character_limit))

        def update_color_field(color_field, hex_value):
            if not hex_value.startswith("#"):
                hex_value = f"#{hex_value}"
            try:
                if len(hex_value) == 7:
                    color_field.configure(fg_color=hex_value)
                else:
                    raise ValueError("Invalid Hex Value")
            except ValueError as e:
                print(e)

        def on_hex_entry_confirm(color_field, hex_entry):
            hex_value = hex_entry.get()
            update_color_field(color_field, hex_value)

        def create_lambda(color_field, hex_entry):
            return lambda event: on_hex_entry_confirm(color_field, hex_entry)

        self.hex_entry.bind("<Return>", create_lambda(
            color_field, self.hex_entry))

        return clr_widget_base

    def activate_fields(self):
        """When the color fields are initialized for the first time they are deactivated. 
        This function activates them"""
        self.hex_entry.configure(state="normal")
        self.name_entry.configure(state="normal")
        self.switch.configure(state="normal")

    def deactivate_fields(self):
        """And this one is used to deactivate them again when they are not supposed to be touched"""
        self.hex_entry.configure(state="disabled")
        self.name_entry.configure(state="disabled")
        self.switch.configure(state="disabled")

    def update_color_field_xml(self, clr_hex):
        """Updating the color fields(hex values) with the XML extracted values"""
        clr_val = clr_hex[1:]

        for color_field in self.color_fields:
            color_field.configure(fg_color=clr_hex)

        for hex_entry in self.hex_entry_fields:
            hex_entry.insert(0, clr_val)

    def update_color_name_field_xml(self, clr_name):
        """Updating the color fields(names) with the XML extracted values"""
        for name_entry in self.name_entry_fields:
            name_entry.insert(0, clr_name)

    def update_switches(self, switch_states):
        """Updating the color fields(switches) with the XML extracted values"""
        for switch_var in self.switch_vars:
            switch_var.set(switch_states)

    def build_name_string(self):
        """Gathering color name from the color field"""
        return self.name_entry.get()

    def build_hex_string(self):
        """Gathering hex value from the color field"""
        return self.hex_entry.get()

    def clear_all_custclr(self):
        """This is supposed to clear the custom color widgets completely"""
        base_color = "#181818"
        base_hex = ""
        base_name = ""
        base_switch_state = "off"

        for color_field, hex_entry, name_entry, switch_var in zip(self.color_fields,
                                                                  self.hex_entry_fields,
                                                                  self.name_entry_fields,
                                                                  self.switch_vars):
            color_field.configure(fg_color=base_color)
            hex_entry.delete(0, tkinter.END)
            hex_entry.insert(0, base_hex)
            name_entry.delete(0, tkinter.END)
            name_entry.insert(0, base_name)
            switch_var.set(base_switch_state)
