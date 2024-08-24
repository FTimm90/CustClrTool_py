import customtkinter as ctk

class ManualClass():
    """This class is for defining the labels that make up the explanation on the left side"""
    def __init__(self, manual_master):

        self.manual_master = manual_master

    def number_label(self, manual_master):
        """This is only the number for the explanation steps"""
        self.manual_number_label = ctk.CTkLabel(master=manual_master,
                                                font=("Roboto", 16, "bold"),
                                                text_color="white",
                                                width=10,
                                                justify="left",
                                                corner_radius=0,
                                                fg_color="#1f1f1f",
                                                compound="left")
        self.manual_number_label.configure(wraplength=160)
        self.manual_number_label.pack(anchor="w")

    def text_label(self, manual_master):
        """This is for the actual text of the explanation"""
        self.manual_text_label = ctk.CTkLabel(master=manual_master,
                                              font=("Roboto", 12),
                                              text_color="white",
                                              width=100,
                                              justify="left",
                                              corner_radius=0,
                                              fg_color="#1f1f1f",
                                              compound="left")
        self.manual_text_label.configure(wraplength=160)
        self.manual_text_label.pack(anchor="w")
