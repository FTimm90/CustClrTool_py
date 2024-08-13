from customtkinter import filedialog 
from tkinter import *
import os
from zipfile import ZipFile 
import xml.etree.ElementTree as ET 
from io import BytesIO 
import re

def map_child_to_parent(root):
    parent_map = {c:p for p in root.iter() for c in p} 
    return parent_map

class PptxFile():

    def __init__(self, 
                 pptx_file=None, 
                 pptx_format=None, 
                 pptx_path=None, 
                 zip_filename=None, 
                 found_themes=None,
                 xml_selection=None,
                 hex_colorlist=None,
                 Theme_name=None,
                 full_color_list=None,
                 hex_namelist=None):
        
        self.pptx_file = pptx_file
        self.pptx_format = pptx_format
        self.pptx_path = pptx_path
        self.found_themes = found_themes
        self.zip_filename = zip_filename
        self.xml_selection = xml_selection
        self.hex_colorlist = hex_colorlist if hex_colorlist is not None else []
        self.Theme_name = Theme_name
        self.full_color_list = full_color_list
        self.hex_namelist = hex_namelist

    def open_file(self):

        if self.pptx_file:
            self.pptx_file = None
            self.clear_all_pptx

        self.pptx_file = filedialog.askopenfilename(filetypes=[("PowerPoint Files", "*.pptx *.potx")])

        if self.pptx_file is not None:
            self.pptx_path, self.pptx_format = os.path.splitext(self.pptx_file)

            self.zip_filename = self.pptx_path + ".zip"
            os.rename(self.pptx_file, self.zip_filename)
            print("File has been renamed to: " + self.zip_filename)

            # finding theme.xml files
            self.found_themes = []
            with ZipFile(self.zip_filename, "a") as zip:
                all_files = zip.namelist()
                theme_pattern = re.compile(r'theme\d+\.xml')
                for file in all_files:
                    if theme_pattern.search(file):
                        self.found_themes.append(os.path.basename(file))
                        print(f"Found themes: {self.found_themes}")

        print("the file format is " + self.pptx_format + " and the file name is " + self.pptx_path + self.pptx_format)
        return self.pptx_path, self.pptx_format, self.pptx_file, self.zip_filename, self.found_themes
    
    def register_all_namespaces(self, xml_data):
        namespaces = {}
        it = ET.iterparse(BytesIO(xml_data), events=["start-ns"])
        for _, elem in it:
            prefix, uri = elem
            namespaces[prefix] = uri
            for prefix, uri in namespaces.items():
                ET.register_namespace(prefix, uri)
                return namespaces

    # extracting colors from the opened theme.xml
    def find_custClrLst(self):
        with ZipFile(self.zip_filename, "a") as zip:
            with zip.open(self.xml_selection) as theme_xml:
                xml_data = theme_xml.read()
                namespaces = self.register_all_namespaces(xml_data)
                tree = ET.ElementTree(ET.fromstring(xml_data))
                root = tree.getroot()
                
                self.hex_colorlist = []
                self.hex_namelist = []
                self.temp_clr_list = []

                theme_element = None
                for elem in root.iter():
                    if elem.tag == "{http://schemas.openxmlformats.org/drawingml/2006/main}theme":
                        theme_element = elem
                        break

                if theme_element is not None:
                    theme_name = theme_element.get("name")
                    self.Theme_name = theme_name
                    print(f"Theme name: {self.Theme_name}")
                else:
                    print("Could not find theme name")

                custClrLst = root.find(".//a:custClrLst", namespaces)
                if custClrLst is not None:
                    self.hex_colorlist = [custClr.get("val") for custClr in custClrLst.findall(".//a:srgbClr", namespaces)]
                    self.hex_namelist = [custClr.get("name") for custClr in custClrLst.findall(".//a:custClr", namespaces) if custClr.get("name") is not None]
                    print(f"Existing custom colors: {self.hex_colorlist}")
                    print(f"Existing custom color names: {self.hex_namelist}")
                    return self.hex_colorlist, self.Theme_name, self.hex_namelist
                else:
                    print(f"There are no existing custom colors in {self.xml_selection}.")
                    return [], self.Theme_name

    def complete_color_list(self):
        base_color_element = "181818"
        self.full_color_list = [base_color_element] * 50
        replacement_number = len(self.hex_colorlist)
        self.full_color_list[:replacement_number] = self.hex_colorlist
        self.full_color_list = [f"#{color}" for color in self.full_color_list]
        return self.full_color_list
    
    def complete_name_list(self):
        base_name_element = ""
        self.full_name_list = [base_name_element] * 50
        replacement_number = len(self.hex_namelist)
        print("The numer of color names in the theme is: " + str(replacement_number))
        self.full_name_list[:replacement_number] = self.hex_namelist
        return self.full_name_list
    
    def complete_state_list(self):
        base_state_element = "off"
        changed_state_element = "on"
        full_state_list = [base_state_element] * 50
        replacement_number = len(self.hex_colorlist)
        print("The numer of hex values in the theme is: " + str(replacement_number))
        full_state_list[:replacement_number] = [changed_state_element] * replacement_number
        return full_state_list
    
    def clear_all_pptx(self):
        self.pptx_format=None, 
        self.pptx_path=None, 
        self.zip_filename=None,
        self.found_themes=None,
        self.xml_selection=None,
        self.hex_colorlist=None,
        self.Theme_name=None,
        self.full_color_list=None,
        self.hex_namelist=None

    def file_output(self, colors_string, theme_selection):
        with ZipFile(self.zip_filename, "a") as original_zip:
            with ZipFile("temp.zip", "w") as new_zip:
                for item in original_zip.namelist():
                    if item == theme_selection:
                        xml_data = original_zip.read(item)

                        self.register_all_namespaces(xml_data)
                        root = ET.fromstring(xml_data)
                        parent_map = map_child_to_parent(root) 

                        extLst_element = root.find(".//a:extLst", namespaces={"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
                        if (extLst_element is None):
                            extLst_element = ET.SubElement(root, "a:extLst")
                            print("a:extLst element created") 

                        CstClrExist = root.find(".//a:custClrLst", namespaces={"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})  # Searching for the <a:custClrLst> command
                        if CstClrExist is not None:
                            parent_of_CstClr = parent_map[CstClrExist]
                            parent_of_CstClr.remove(CstClrExist)
                            print("'a:custClrLst' was deleted successfully")
                        else:
                            print("'a:custClrLst' could not be found")

                        if extLst_element is not None:
                            parent = parent_map[extLst_element]
                            index = list(parent).index(extLst_element)
                            parent.insert(index, colors_string)
                        else:
                            print("<a:extLst> could not be found")

                        xml_bytes = BytesIO()
                        tree = ET.ElementTree(root)
                        tree.write(xml_bytes, encoding="utf-8", xml_declaration=True)
                        modified_xml_data = xml_bytes.getvalue()

                        new_zip.writestr(item, modified_xml_data)

                    else:
                        new_zip.writestr(item, original_zip.read(item))

        try:
            with ZipFile("temp.zip", "r") as temp_zip:
                pass
            print("Temporary zip created successfully")
        except Exception as e:
            print(f"Error writing temporary zip: {e}")

        os.remove(self.zip_filename)
        os.rename("temp.zip", self.pptx_path + self.pptx_format)
        self.pptx_file = None
        self.clear_all_pptx
        print("File created succesfully")