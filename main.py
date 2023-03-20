# ! The main script of SiPe.Sat Risat harvester

# ---------------------------- DOCUMENTATION ---------------------------- #

# The main file of SiPe.Sat
# This file is unified: all classes are declared in this file instead of
# splitted into separate directories/files

# Created on 2023-03-20

# WEB ARTICLES USED AS REFERENCES:
# 1. Tkinter Application to Switch Between Different Page Frames
#  -> https://www.geeksforgeeks.org/tkinter-application-to-switch-between-different-page-frames/
# 2. Python GUI examples (Tkinter Tutorial)
#  -> https://likegeeks.com/python-gui-examples-tkinter-tutorial/
# 3. Python Tkinter – Entry Widget
#  -> https://www.geeksforgeeks.org/python-tkinter-entry-widget/
# 4. How to justify text in label in Tkinter
#  -> https://stackoverflow.com/questions/37318060
# 5. RadioButton in Tkinter | Python
#  -> https://www.geeksforgeeks.org/radiobutton-in-tkinter-python/
# 6. Tkinter Grid
#  -> https://www.pythontutorial.net/tkinter/tkinter-grid/

# CODING CONVENTIONS:
# 1. Use single quote (') instead of double quotes (") when specifying strings
# 2. Equal signs used as function argument are not wrapped by empty space bars
# 3. 'SipesatScr...' class name prefix indicates a class of the superclass 'tk.Frame'

# VARIABLE CONVENTIONS:
# 1. 'Jenis Data' radio button in the harvester menu has the following possible IntVar values:
#  -> [0] = Data Ringkasan
#  -> [1] = Data Detil Lengkap
# 2. 'Data Hasil Panenan' radio button in the harvester menu has the following possible StringVar values:
#  -> ['dana'] = Risat Dana Penelitian/AbdiMas
#  -> ['arsip'] = Risat Arsip Penelitian/AbdiMas
#  -> ['lapakhir'] = = Risat Laporan Akhir Penelitian/AbdiMas

# --------------------------- CODE PREAMBLE --------------------------- #

# Modules import
from tkinter import scrolledtext
from tkinter import ttk
from tkinter import IntVar
from tkinter import StringVar
import tkinter as tk

# ------------------------ CLASSES DECLARATION ------------------------ #

# Main class declaration
# This class is called first to manage frames and Tkinter instances
class MainGUI(tk.Tk):
    
    # The init function for the MainGUI class
    def __init__(self, *args, **kwargs):
        
        # Instantiating Tkinter instance
        tk.Tk.__init__(self, *args, **kwargs)
        
        # Establishing the GUI container
        container = tk.Frame(self)
        container.pack(side='top', fill='both', expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        # Setting up the list of frames
        self.frames = {}
        
        # Iterating through frame list
        for F in FRAME_CLASSES:
            
            # Initializing the frame
            frame = F(container, self)
            
            # Assigning the initialized frame into
            # the frame list
            self.frames[F] = frame
            
            # Setting the frame's layout
            # Sticky: 'nsew' stands for 'north-south-east-west'
            frame.grid(row=0, column=0, sticky='nsew')
        
        # By default displaying the first frame class specified in
        # the tuple 'FRAME_CLASSES'
        # This will be called upon at the beginning of the app
        self.raise_frame(FRAME_CLASSES[0])
        
    # The function that displays the frame
    def raise_frame(self, frame_choice):
        frame = self.frames[frame_choice]
        frame.tkraise()

# The login frame
# This screen asks Risat login credentials, i.e., staff ID and password
class SipesatScrLogin(tk.Frame):
    
    # The function that will be triggered when the 'ABOUT' button is pressed
    def on_about_button_click(self):
        pass
    
    # The function that will be triggered when the 'LICENSE' button is pressed
    def on_license_button_click(self):
        pass
    
    # The function that will be triggered when the 'EXIT' button is pressed
    def on_exit_button_click(self):
        pass
    
    # The function that will be triggered when the 'LOGIN' form button is pressed
    def on_submit_button_click(self, username, password):
        pass
    
    # The __init__ function
    def __init__(self, parent, controller):
        
        # Instantiating tkinter.Frame instance
        tk.Frame.__init__(self, parent)
        
        # The header title displaying the screen's title 
        header_title = ttk.Label(self, text=STRING_HEADER_TITLE, font=FONT_HEADER_TITLE)
        header_title.grid(row=0, column=0, padx=10, pady=2)
        
        # The header description displaying additional information about the screen
        header_desc = ttk.Label(self, text=STRING_HEADER_DESC, font=FONT_HEADER_DESC)
        header_desc.grid(row=1, column=0, padx=10, pady=5)
        
        # The login display help
        help_text = 'Selamat datang di sistem pemanen Risat Satya Wacana!\nSilahkan masukkan kredensial berikut ini:'
        help_label = ttk.Label(self, text=help_text, font=FONT_REGULAR, anchor='w', justify='left')
        help_label.grid(row=2, column=0, padx=10, pady=10)
        
        # :::
        # The login form frame
        form_frame = tk.Frame(self)
        form_frame.grid(row=3, column=0, padx=10, pady=10)
        # Username display help
        user_text = 'ID Pegawai'
        user_label = ttk.Label(form_frame, text=user_text,)
        user_label.grid(row=0, column=0, padx=2, pady=2)
        # Username input prompt
        user_input = ttk.Entry(form_frame, width=50)
        user_input.grid(row=0, column=1, padx=2, pady=2)
        # Password display help
        pass_text = 'Password'
        pass_label = ttk.Label(form_frame, text=pass_text,)
        pass_label.grid(row=1, column=0, padx=2, pady=2)
        # Password input prompt
        pass_input = ttk.Entry(form_frame, width=50, show='*')
        pass_input.grid(row=1, column=1, padx=2, pady=2)
        # The submit button
        submit_text = 'LOGIN'
        submit_button = ttk.Button(form_frame, text=submit_text,
            command=lambda: self.on_submit_button_click(user_input.get(), pass_input.get()))
        submit_button.grid(row=2, column=0, padx=2, pady=2)
        
        # :::
        # The footer buttons frame container
        footer_frame = tk.Frame(self)
        footer_frame.grid(row=4, column=0, padx=10, pady=10)
        # The 'about' page button trigger
        about_text = 'TENTANG'
        about_button = ttk.Button(footer_frame, text=about_text,
            command=lambda: self.on_about_button_click())
        about_button.grid(row=0, column=0, padx=2, pady=2)
        # The 'license' page button trigger
        license_text = 'LISENSI'
        license_button = ttk.Button(footer_frame, text=license_text,
            command=lambda: self.on_license_button_click())
        license_button.grid(row=0, column=1, padx=2, pady=2)
        # The 'exit' page button trigger
        exit_text = 'KELUAR'
        exit_button = ttk.Button(footer_frame, text=exit_text,
            command=lambda: self.on_exit_button_click())
        exit_button.grid(row=0, column=2, padx=2, pady=2)

# The main menu
# This screen is displayed upon successful login
# Each button in this screen directs the user into the harvester menu
# of each category, e.g., 'Research' and 'Community Service'
class SipesatScrMainMenu(tk.Frame):
    
    # The function that will be triggered when the menu [1] button is selected/clicked
    def on_menu_1_button_click(self):
        pass
    
    # The function that will be triggered when the menu [2] button is selected/clicked
    def on_menu_2_button_click(self):
        pass
    
    # The __init__ function
    def __init__(self, parent, controller):
        
        # Instantiating tkinter.Frame instance
        tk.Frame.__init__(self, parent)
        
        # The header title displaying the screen's title 
        header_title = ttk.Label(self, text=STRING_HEADER_TITLE, font=FONT_HEADER_TITLE)
        header_title.grid(row=0, column=0, padx=10, pady=2)
        
        # The header description displaying additional information about the screen
        desc_text = 'Menu Utama'
        header_desc = ttk.Label(self, text=desc_text, font=FONT_HEADER_DESC)
        header_desc.grid(row=1, column=0, padx=10, pady=5)
        
        # The screen's display help
        help_text = 'Selamat datang di menu utama pemanen Risat Satya Wacana!\nPilih tombol di bawah ini untuk melanjutkan pemanenan'
        help_label = ttk.Label(self, text=help_text, font=FONT_REGULAR, anchor='w', justify='left')
        help_label.grid(row=2, column=0, padx=10, pady=10)
        
        # :::
        # The button layout frame
        layout_frame = tk.Frame(self)
        layout_frame.grid(row=3, column=0, padx=10, pady=10)
        # Menu [1]: Research
        menu_1_text = 'Risat Penelitian'
        menu_1_button = ttk.Button(layout_frame, text=menu_1_text, width=45,
            command=lambda: self.on_menu_1_button_click())
        menu_1_button.grid(row=0, column=0, padx=2, pady=2)
        # Menu [2]: Comm. Service
        menu_2_text = 'Risat Pengabdian Masyarakat'
        menu_2_button = ttk.Button(layout_frame, text=menu_2_text, width=45,
            command=lambda: self.on_menu_2_button_click())
        menu_2_button.grid(row=1, column=0, padx=2, pady=2)

# The harvester menu for 'research'
# - Displayed before the actual harvesting process is executed, for
# preconfiguring what kind of data and which database to be harvested
# - Called upon following the 'SipesatScrMainMenu' screen
class SipesatScrResearch(tk.Frame):
    
    # The function that will be triggered when the 'kembali' button is hit
    def on_back_button_click(self):
        pass
    
    # The function that will be triggered when the 'mulai' button is hit
    def on_start_button_click(self):
        pass
    
    # The __init__ function
    def __init__(self, parent, controller):
        
        # The radiobutton variables, used to group different radiobutton together
        self.radio_data_type = IntVar()
        self.radio_harvest_type = StringVar()
        
        # Instantiating tkinter.Frame instance
        tk.Frame.__init__(self, parent)
        
        # The header title displaying the screen's title 
        header_title = ttk.Label(self, text=STRING_HEADER_TITLE, font=FONT_HEADER_TITLE)
        header_title.grid(row=0, column=0, padx=10, pady=2)
        
        # The header description displaying additional information about the screen
        desc_text = 'Menu Utama » Risat Penelitian'
        header_desc = ttk.Label(self, text=desc_text, font=FONT_HEADER_DESC)
        header_desc.grid(row=1, column=0, padx=10, pady=5)
        
        # The screen's display help
        help_text = 'Silahkan tentukan basis data dan jenis data yang akan dipanen\nTekan "MULAI" untuk memulai proses pemanenan data Risat'
        help_label = ttk.Label(self, text=help_text, font=FONT_REGULAR, anchor='w', justify='left')
        help_label.grid(row=2, column=0, padx=10, pady=10)
        
        # :::
        # Layout [1]: Data type
        layout_frame_1 = tk.Frame(self)
        layout_frame_1.grid(row=3, column=0, padx=10, pady=10)
        # The help label
        datatype_text = 'Jenis Data'
        datatype_label = ttk.Label(layout_frame_1, text=datatype_text, font=FONT_REGULAR)
        datatype_label.grid(row=0, column=0, padx=2, pady=5, columnspan=2)
        # Radio button 1 -- 'Summary'
        datatype_summary_text = 'Ringkasan Data'
        datatype_summary_radio = ttk.Radiobutton(layout_frame_1, text=datatype_summary_text,
            value=0, variable=self.radio_data_type)
        datatype_summary_radio.grid(row=1, column=0, padx=2, pady=2, sticky='w')
        # Radio button 2 -- 'Detailed data'
        datatype_detailed_text = 'Data Detil Lengkap'
        datatype_detailed_radio = ttk.Radiobutton(layout_frame_1, text=datatype_detailed_text,
            value=1, variable=self.radio_data_type)
        datatype_detailed_radio.grid(row=1, column=1, padx=2, pady=2, sticky='w')
        
        # :::
        # Layout [2]: Harvested data output
        layout_frame_2 = tk.Frame(self)
        layout_frame_2.grid(row=4, column=0, padx=10, pady=10)
        # The help label
        harvesttype_text = 'Data Hasil Panenan'
        harvesttype_label = ttk.Label(layout_frame_2, text=harvesttype_text, font=FONT_REGULAR)
        harvesttype_label.grid(row=0, column=0, padx=2, pady=5)
        # Radio button 1 -- 'Dana'
        harvesttype_dana_text = 'Risat Dana Penelitian'
        harvesttype_dana_radio = ttk.Radiobutton(layout_frame_2, text=harvesttype_dana_text,
            value='dana', variable=self.radio_harvest_type)
        harvesttype_dana_radio.grid(row=1, column=0, padx=2, pady=2, sticky='w')
        # Radio button 2 -- 'Arsip'
        harvesttype_arsip_text = 'Risat Arsip Penelitian'
        harvesttype_arsip_radio = ttk.Radiobutton(layout_frame_2, text=harvesttype_arsip_text,
            value='arsip', variable=self.radio_harvest_type)
        harvesttype_arsip_radio.grid(row=2, column=0, padx=2, pady=2, sticky='w')
        # Radio button 3 -- 'Laporan Akhir'
        harvesttype_lapakhir_text = 'Risat Laporan Akhir Penelitian'
        harvesttype_lapakhir_radio = ttk.Radiobutton(layout_frame_2, text=harvesttype_lapakhir_text,
            value='lapakhir', variable=self.radio_harvest_type)
        harvesttype_lapakhir_radio.grid(row=3, column=0, padx=2, pady=2, sticky='w')
        
        # :::
        # Layout [3]: The action buttons
        layout_frame_3 = tk.Frame(self)
        layout_frame_3.grid(row=5, column=0, padx=10, pady=10)
        # The 'back' button trigger
        # Clicking this button calls out 'SipesatScrMainMenu'
        back_text = 'KEMBALI'
        back_button = ttk.Button(layout_frame_3, text=back_text, width=40,
            command=lambda: self.on_back_button_click())
        back_button.grid(row=0, column=0, padx=2, pady=2)
        # The 'start' button trigger
        # Clicking this button proceeds the program to the harvester screen
        start_text = 'MULAI'
        start_button = ttk.Button(layout_frame_3, text=start_text, width=40,
            command=lambda: self.on_start_button_click())
        start_button.grid(row=0, column=1, padx=2, pady=2)

# The harvester menu for 'community service'
# - Displayed before the actual harvesting process is executed, for
# preconfiguring what kind of data and which database to be harvested
# - Called upon following the 'SipesatScrMainMenu' screen
class SipesatScrComService(tk.Frame):
    
    # The function that will be triggered when the 'kembali' button is hit
    def on_back_button_click(self):
        pass
    
    # The function that will be triggered when the 'mulai' button is hit
    def on_start_button_click(self):
        pass
    
    # The __init__ function
    def __init__(self, parent, controller):
        
        # The radiobutton variables, used to group different radiobutton together
        self.radio_data_type = IntVar()
        self.radio_harvest_type = StringVar()
        
        # Instantiating tkinter.Frame instance
        tk.Frame.__init__(self, parent)
        
        # The header title displaying the screen's title 
        header_title = ttk.Label(self, text=STRING_HEADER_TITLE, font=FONT_HEADER_TITLE)
        header_title.grid(row=0, column=0, padx=10, pady=2)
        
        # The header description displaying additional information about the screen
        desc_text = 'Menu Utama » Risat Pengabdian Masyarakat'
        header_desc = ttk.Label(self, text=desc_text, font=FONT_HEADER_DESC)
        header_desc.grid(row=1, column=0, padx=10, pady=5)
        
        # The screen's display help
        help_text = 'Silahkan tentukan basis data dan jenis data yang akan dipanen\nTekan "MULAI" untuk memulai proses pemanenan data Risat'
        help_label = ttk.Label(self, text=help_text, font=FONT_REGULAR, anchor='w', justify='left')
        help_label.grid(row=2, column=0, padx=10, pady=10)
        
        # :::
        # Layout [1]: Data type
        layout_frame_1 = tk.Frame(self)
        layout_frame_1.grid(row=3, column=0, padx=10, pady=10)
        # The help label
        datatype_text = 'Jenis Data'
        datatype_label = ttk.Label(layout_frame_1, text=datatype_text, font=FONT_REGULAR)
        datatype_label.grid(row=0, column=0, padx=2, pady=5, columnspan=2)
        # Radio button 1 -- 'Summary'
        datatype_summary_text = 'Ringkasan Data'
        datatype_summary_radio = ttk.Radiobutton(layout_frame_1, text=datatype_summary_text,
            value=0, variable=self.radio_data_type)
        datatype_summary_radio.grid(row=1, column=0, padx=2, pady=2, sticky='w')
        # Radio button 2 -- 'Detailed data'
        datatype_detailed_text = 'Data Detil Lengkap'
        datatype_detailed_radio = ttk.Radiobutton(layout_frame_1, text=datatype_detailed_text,
            value=1, variable=self.radio_data_type)
        datatype_detailed_radio.grid(row=1, column=1, padx=2, pady=2, sticky='w')
        
        # :::
        # Layout [2]: Harvested data output
        layout_frame_2 = tk.Frame(self)
        layout_frame_2.grid(row=4, column=0, padx=10, pady=10)
        # The help label
        harvesttype_text = 'Data Hasil Panenan'
        harvesttype_label = ttk.Label(layout_frame_2, text=harvesttype_text, font=FONT_REGULAR)
        harvesttype_label.grid(row=0, column=0, padx=2, pady=5)
        # Radio button 1 -- 'Dana'
        harvesttype_dana_text = 'Risat Dana Pengabdian Masyarakat'
        harvesttype_dana_radio = ttk.Radiobutton(layout_frame_2, text=harvesttype_dana_text,
            value='dana', variable=self.radio_harvest_type)
        harvesttype_dana_radio.grid(row=1, column=0, padx=2, pady=2, sticky='w')
        # Radio button 2 -- 'Arsip'
        harvesttype_arsip_text = 'Risat Arsip Pengabdian Masyarakat'
        harvesttype_arsip_radio = ttk.Radiobutton(layout_frame_2, text=harvesttype_arsip_text,
            value='arsip', variable=self.radio_harvest_type)
        harvesttype_arsip_radio.grid(row=2, column=0, padx=2, pady=2, sticky='w')
        # Radio button 3 -- 'Laporan Akhir'
        harvesttype_lapakhir_text = 'Risat Laporan Akhir Pengabdian Masyarakat'
        harvesttype_lapakhir_radio = ttk.Radiobutton(layout_frame_2, text=harvesttype_lapakhir_text,
            value='lapakhir', variable=self.radio_harvest_type)
        harvesttype_lapakhir_radio.grid(row=3, column=0, padx=2, pady=2, sticky='w')
        
        # :::
        # Layout [3]: The action buttons
        layout_frame_3 = tk.Frame(self)
        layout_frame_3.grid(row=5, column=0, padx=10, pady=10)
        # The 'back' button trigger
        # Clicking this button calls out 'SipesatScrMainMenu'
        back_text = 'KEMBALI'
        back_button = ttk.Button(layout_frame_3, text=back_text, width=40,
            command=lambda: self.on_back_button_click())
        back_button.grid(row=0, column=0, padx=2, pady=2)
        # The 'start' button trigger
        # Clicking this button proceeds the program to the harvester screen
        start_text = 'MULAI'
        start_button = ttk.Button(layout_frame_3, text=start_text, width=40,
            command=lambda: self.on_start_button_click())
        start_button.grid(row=0, column=1, padx=2, pady=2)

# The harvester screen
# Displaying the current progress of the harvesting process
class SipesatScrHarvest(tk.Frame):
    
    # The __init__ function
    def __init__(self, parent, controller):
        
        # Instantiating tkinter.Frame instance
        tk.Frame.__init__(self, parent)
        
        # The header title displaying the screen's title 
        header_title = ttk.Label(self, text=STRING_HEADER_TITLE, font=FONT_HEADER_TITLE)
        header_title.grid(row=0, column=0, padx=10, pady=2)
        
        # The header description displaying additional information about the screen
        desc_text = 'Panen Data "[...placeholder...]"'
        header_desc = ttk.Label(self, text=desc_text, font=FONT_HEADER_DESC)
        header_desc.grid(row=1, column=0, padx=10, pady=5)
        
        # The screen's display help
        help_text = 'Data "[...placeholder...]" sedang dipanen. Silahkan menunggu...'
        help_label = ttk.Label(self, text=help_text, font=FONT_REGULAR, anchor='w', justify='left')
        help_label.grid(row=2, column=0, padx=10, pady=10)
        
        # :::
        # Layout harvester progress bar display
        progress_frame = tk.Frame(self)
        progress_frame.grid(row=3, column=0, padx=10, pady=10)
        # The progress status/value display
        progress_text = 'Status: 67%'
        progress_label = ttk.Label(progress_frame, text=progress_text, font=FONT_PROGRESS_VALUE, anchor='center', justify='center')
        progress_label.grid(row=0, column=0, padx=2, pady=5)
        # The progress bar
        progress_bar = ttk.Progressbar(progress_frame, length=600)
        progress_bar['value'] = 20
        progress_bar.grid(row=1, column=0, padx=2, pady=2, sticky='we')
        
        # The log message displayer
        message_area = scrolledtext.ScrolledText(self, width=100, height=10)
        message_area.grid(row=4, column=0, padx=5, pady=5)

# -------------------------- CONSTANT PRESETS -------------------------- #

# 'FRAME_CLASSES' is a tuple that defines all the frame classes of the file
# - The first class specified in this tuple will be the frame displayed
#   at the very beginning of the application after launch
# - Need to have at least two classes, otherwise the following error will be casted:
#   TypeError: 'type' object is not iterable
FRAME_CLASSES = (SipesatScrHarvest, SipesatScrComService, SipesatScrResearch, SipesatScrMainMenu, SipesatScrLogin)

# The following constants define font presets used in styling Tkinter widgets
FONT_HEADER_TITLE = ('Segoe UI', 30, 'bold')
FONT_HEADER_DESC = ('Segoe UI', 20, 'italic')
FONT_REGULAR = ('Segoe UI', 12)
FONT_FORM_INPUT = ('Courier New', 10, 'bold')
FONT_RADIO = ('Courier New', 10)
FONT_PROGRESS_VALUE = ('Courier New', 14)

# The following constants define the string presets used as template and localization
STRING_HEADER_TITLE = 'SiPe.Sat'
STRING_HEADER_DESC = 'Sistem Pemanen Satya Wacana'

# ------------------------- APPLICATION LAUNCH ------------------------- #

# Initializing the GUI
app_gui = MainGUI()
app_gui.mainloop()
