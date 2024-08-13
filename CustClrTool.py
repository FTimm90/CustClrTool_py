from modules import *
import customtkinter as ctk
from PIL import Image
import os
import xml.etree.ElementTree as ET
import sys

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

window = ctk.CTk()
window.geometry("1200x800")
window.title("CustClr Tool")
window.resizable(width=False, height=False)

bottom_bar = ctk.CTkFrame(master=window, 
                          height=20, 
                          corner_radius=0, 
                          fg_color="#1f1f1f")
bottom_bar.pack(side="bottom", fill="x")

output_label = ctk.CTkLabel(master=bottom_bar, 
                             width=600, 
                             anchor="w", 
                             text=" ", 
                             padx=10, 
                             text_color="#6D6D6D")
output_label.pack(side="left")

# Text updates for output label
def file_opened(lbl):
    themenumber = len(pptx_instance.found_themes)
    if themenumber > 0:
        str_end = ""
        if themenumber >= 2:
            str_end = "themes"
        else:
            str_end = "theme"
        newtext = "File opened, " + str(themenumber) + " " + str_end + " found"
        lbl.configure(text=newtext)
    else:
        newtext = "No themes found"
        lbl.configure(text=newtext)

def Clrs_found(lbl):
    Clrs_found = len(pptx_instance.hex_colorlist)
    if Clrs_found > 0:
        newtext = str(Clrs_found) + " custom colors found"
        lbl.configure(text=newtext)
    else:
        newtext = "No existing custom colors found"
        lbl.configure(text=newtext)

version_label = ctk.CTkLabel(master=bottom_bar, 
                             width=200, 
                             anchor="e", 
                             text="v 1.0", 
                             padx=10, 
                             text_color="#6D6D6D")
version_label.pack(side="right")

bottom_divider_line = ctk.CTkFrame(master=window, 
                           height=1, 
                           fg_color="#ffffff")
bottom_divider_line.pack(side="bottom", fill="x")

left_frame = ctk.CTkFrame(master=window, 
                          width=200, 
                          height=579, 
                          corner_radius=0, 
                          fg_color="#1f1f1f")
left_frame.pack(side="left", fill="y")

icon_path = resource_path("graphics/color_icon.png")
import_color_icon = Image.open(icon_path)
color_icon = ctk.CTkImage(import_color_icon)

CustClrBtn = ctk.CTkButton(master=left_frame, 
                            image=color_icon, 
                            text="Custom Colors", 
                            fg_color="#1f1f1f",
                            text_color="#A3A3A3", 
                            width=200, 
                            height=50, 
                            corner_radius=0, 
                            anchor="w",
                            compound="left",
                            hover=False)
CustClrBtn.pack(side="top", padx=30, pady=20)

content_area = ctk.CTkFrame(master=window, 
                            corner_radius=0, 
                            fg_color="#181818")
content_area.pack(side="right", fill="both", expand=True)

button_label_container = ctk.CTkFrame(master=content_area, 
                                      corner_radius=0, 
                                      fg_color="#181818")
button_label_container.pack(pady=30, padx=40, anchor="w")

pptx_instance = pptxClass.PptxFile()

open_file_button = ctk.CTkButton(master=button_label_container, 
                                 command=lambda:[clear_color_fields(),
                                                 pptx_instance.open_file(), 
                                                 create_dropdown(), 
                                                 pptx_file_name_label(file_name_label),
                                                 clear_name_label_func(theme_name_label),
                                                 file_opened(output_label)],
                                 text="choose file", 
                                 width=100, 
                                 height=30, 
                                 fg_color="#181818", 
                                 hover_color="#76AE22", 
                                 border_color="#76AE22", 
                                 border_width=1, 
                                 corner_radius=4)
open_file_button.pack(side="left")

file_name_label = ctk.CTkLabel(master=button_label_container, 
                               width=400, 
                               anchor="w", 
                               text_color="#6D6D6D",
                               text=" ")
file_name_label.pack(side="left", padx=20)

def pptx_file_name_label(lbl):
        filename = str(pptx_instance.pptx_path).split("/")[-1] + str(pptx_instance.pptx_format)
        lbl.configure(text=filename)

dropdown_label_container = ctk.CTkFrame(master=content_area, 
                                      corner_radius=0, 
                                      fg_color="#181818",
                                      width=600,
                                      height=20)
dropdown_label_container.pack(pady=0, padx=40, anchor="w", fill="x")

dropdown_container = ctk.CTkFrame(master=dropdown_label_container,
                                  corner_radius=0,
                                  fg_color="#181818",
                                  height=30,
                                  width=100)
dropdown_container.pack(side="left", anchor="w")

theme_name_label = ctk.CTkLabel(master=dropdown_label_container, 
                               width=400, 
                               anchor="w", 
                               text_color="#6D6D6D",
                               text=" ")
theme_name_label.pack(side="left", padx=20)

current_dropdown = None

def create_dropdown():
    global current_dropdown
    clear_color_fields

    if current_dropdown:
        current_dropdown.destroy()
        clear_color_fields

    optionmenu = ctk.CTkOptionMenu(master=dropdown_container, 
                                  values=pptx_instance.found_themes, 
                                  command=optionmenu_callback, 
                                  corner_radius=4, 
                                  fg_color="#1f1f1f", 
                                  button_color="#1f1f1f", 
                                  button_hover_color="#507216")
    optionmenu.pack(anchor="w")
    current_dropdown = optionmenu

def optionmenu_callback(selection):
    clear_color_fields
    selection = "ppt/theme/" + str(selection)
    print("Dropdown-selection:", selection)
    pptx_instance.xml_selection = selection
    pptxClass.PptxFile.find_custClrLst(pptx_instance)
    theme_name_label_func(theme_name_label)
    Clrs_found(output_label)
    return selection

def theme_name_label_func(lbl):
        filename_theme = str(pptx_instance.Theme_name)
        lbl.configure(text=filename_theme)

        fg_color_list = pptx_instance.complete_color_list()
        clr_namelist = pptx_instance.complete_name_list()
        switch_states = pptx_instance.complete_state_list()

        for instance, fg_color in zip(CustClr_instances, fg_color_list):
            instance.update_color_field_xml(Clr_hex=fg_color)

        for instance, name in zip(CustClr_instances, clr_namelist):
            instance.update_color_name_field_xml(Clr_name=name)

        for instance, switch_state in zip(CustClr_instances, switch_states):
            instance.update_switches(switch_states=switch_state)

def clear_name_label_func(lbl): 
        filename_theme = " "
        lbl.configure(text=filename_theme)

Color_field_frame = ctk.CTkFrame(master=content_area,
                                 fg_color="#181818",
                                 corner_radius=0)
Color_field_frame.pack(padx=40, pady=40, anchor="w")

CustClr_instances = []

for row in range(5):
    for column in range(10):
        CustClr_instance = CustClrClass.CustClr(Clr_master=Color_field_frame, Clr_table_row=row, Clr_table_column=column)
        CustClr_instance.CustClr_Widget(Clr_master=Color_field_frame, Clr_table_row=row, Clr_table_column=column)
        CustClr_instances.append(CustClr_instance)

def construct_output_strings(CustClr_instances):
  root = ET.Element("a:custClrLst", nsmap={"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
  
  for instance in CustClr_instances:
    if instance.switch_var.get() == "on":
      cust_clr = ET.SubElement(root, "a:custClr")
      name_string = instance.build_name_string() 
      cust_clr.set("name", name_string)
      srgb_clr = ET.SubElement(cust_clr, "a:srgbClr")
      srgb_clr.set("val", instance.build_hex_string())
  
  for elem in root.iter():
    if "nsmap" in elem.attrib:
      del elem.attrib["nsmap"]

  return root

def final_output_function():
    final_xml_tree = construct_output_strings(CustClr_instances)
    theme = pptx_instance.xml_selection
    pptx_instance.file_output(colors_string=final_xml_tree, theme_selection=theme)

def clear_color_fields():
    for instance in CustClr_instances:
        instance.clear_all_CustClr()

def confirm():
    try:
        final_output_function()
        print("Custom colors successfully added")
        newtext = "Custom colors succesfully added"
        output_label.configure(text=newtext)
        clear_color_fields()
        current_dropdown.destroy()
        file_name_label.configure(text=" ")
        theme_name_label.configure(text=" ")
    except Exception as e:
        print(f"Custom colors could not be added: {e}")
        newtext = "Custom colors could not be added"
        output_label.configure(text=newtext)

confirm_button = ctk.CTkButton(master=content_area,
                               command=confirm,
                               text="Confirm",
                               width=100,
                               height=30,
                               fg_color="#181818",
                               hover_color="#76AE22",
                               border_color="#76AE22",
                               border_width=1,
                               corner_radius=4)
confirm_button.pack(pady=30, padx=40, anchor="w")

window.mainloop()