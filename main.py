# ! The main script of SiPe.Sat Risat harvester

# ---------------------------- DOCUMENTATION ---------------------------- #

# The main file of SiPe.Sat
# This file is unified: all classes are declared in this file instead of
# splitted into separate directories/files

# Created on 2023-03-20

# [1] WEB ARTICLES USED AS REFERENCES:
# 
# (On 2023-03-20)
#  1. Tkinter Application to Switch Between Different Page Frames
#   -> https://www.geeksforgeeks.org/tkinter-application-to-switch-between-different-page-frames/
#  2. Python GUI examples (Tkinter Tutorial)
#   -> https://likegeeks.com/python-gui-examples-tkinter-tutorial/
#  3. Python Tkinter – Entry Widget
#   -> https://www.geeksforgeeks.org/python-tkinter-entry-widget/
#  4. How to justify text in label in Tkinter
#   -> https://stackoverflow.com/questions/37318060
#  5. RadioButton in Tkinter | Python
#   -> https://www.geeksforgeeks.org/radiobutton-in-tkinter-python/
#  6. Tkinter Grid
#   -> https://www.pythontutorial.net/tkinter/tkinter-grid/
# 
# (On 2023-03-21)
#  7. How do I close a tkinter window?
#   -> https://stackoverflow.com/questions/110923
#  8. How to set the text/value/content of an `Entry` widget using a button in tkinter
#   -> https://stackoverflow.com/questions/16373887
#  9. Python Switch Statement – Switch Case Example
#   -> https://www.freecodecamp.org/news/python-switch-statement-switch-case-example/
# 10. Disable / Enable Button in TKinter
#   -> https://stackoverflow.com/questions/53580507
#
# (On 2023-03-28)
# 11. How to change the Tkinter label text?
#   -> https://www.geeksforgeeks.org/how-to-change-the-tkinter-label-text/
# 12. How to get the text out of a scrolledtext widget?
#   -> https://stackoverflow.com/questions/53937400
# 13. Create temporary files and directories using Python-tempfile
#   -> https://www.geeksforgeeks.org/create-temporary-files-and-directories-using-python-tempfile/
# 14. Python path separator [duplicate]
#   -> https://stackoverflow.com/a/50738724

# [2] CODING CONVENTIONS:
# 1. Use single quote (') instead of double quotes (") when specifying strings
# 2. Equal signs used as function argument are not wrapped by empty space bars
# 3. 'SipesatScr...' class name prefix indicates a class of the superclass 'tk.Frame'

# [3] VARIABLE CONVENTIONS:
# 1. 'Kategori' (category)
#  -> ['r'] = The Risat Research menu category
#  -> ['c'] = The Risat ComService menu category
# 2. 'Jenis Data' (datatype) radio button in the harvester menu has the following possible IntVar values:
#  -> [0] = Data Ringkasan
#  -> [1] = Data Detil Lengkap
# 3. 'Data Hasil Panenan' (output) radio button in the harvester menu has the following possible StringVar values:
#  -> ['dana'] = Risat Dana Penelitian/AbdiMas
#  -> ['arsip'] = Risat Arsip Penelitian/AbdiMas
#  -> ['lapakhir'] = = Risat Laporan Akhir Penelitian/AbdiMas

# [4] APPLICATION MECHANISM:
# 1. Upon successful login, the username & password credentials are stored
#    as variables inside the 'controller' class instance.
#    These variables will then be flushed out upon loggin out.

# [5] "RISAT DATA_PROMPT" ARRAY CONVENTIONS:
# data_prompt = {
#   'http_response'     --> the http_response after posting
#   'html_content'      --> 'http_response', converted into an XML-compatible HTML string
#   'viewstate'         --> the view state hidden ASPX form value (e.g. "/wEPaA8FDzhkYjBkY2I5ODVlM2YzZmTXigQ6dU1GUsc1Dgno6Z11")
#   'viewstategen'      --> the viewstategen hidden ASPX form value (e.g. "28239525")
#   'eventvalidation'   --> the eventvalidation hidden ASPX form value (e.g. "/wEdABiKTaCYJNZ8hl3vzC5OJBTzquwmUN/b9MQs90SJn/")
#   'button_name'       --> the entry row's submit button 'name' attribute (e.g. "ctl00$ContentPlaceHolder1$danacair1$repusulan1$ctl01$btdetil1")
#                           this data array data is used both in 'dana detil' and 'arsip detil' pages
#   'kodetran_prop'     --> the entry row's hidden 'kodetran' tag attribute name (e.g. "ctl00$ContentPlaceHolder1$danacair1$repusulan1$ctl01$kodetran1")
#                           this data array data is used only in 'dana detil' pages
#   'kodetran_val'      --> the entry row's hidden 'kodetran' tag attribute value (e.g. "842BA8DC-F7FA-48CE-A8B8-3106876A3B1E")
#                           this data array data is used only in 'dana detil' pages
#   'stat'              --> backward compatibility of 'stat_prop'
#   'stat_prop'         --> the entry row's hidden 'stat' tag attribute name (e.g. "ctl00$ContentPlaceHolder1$danacair1$repusulan1$ctl01$stat1")
#                           this data array data is used only in 'dana detil' pages
#   'stat_val'          --> the entry row's hidden 'stat' tag attribute value (e.g. "M" for "belum terealisasi, or "C" for "terealisasi")
#                           this data array data is used only in 'dana detil' pages
#   'idat_prop'         --> the entry row's hidden 'data' tag attribute value (e.g. "ctl00$ContentPlaceHolder1$arsip1$repusul1$ctl01$idat1")
#                           this data array data is used only in 'arsip detil' pages
#   'idat_val'          --> the entry row's hidden 'data' tag attribute value (e.g. "12AAFE6D-7C03-4998-9CB6-1BF4ECA37831")
#                           this data array data is used only in 'arsip detil' pages
#   'itgl_prop'         --> the entry row's hidden 'tanggal' tag attribute value (e.g. "ctl00$ContentPlaceHolder1$arsip1$repusul1$ctl01$itgl1")
#                           this data array data is used only in 'arsip detil' pages
#   'itgl_val'          --> the entry row's hidden 'tanggal' tag attribute value (e.g. "2021/01/21")
#                           this data array data is used only in 'arsip detil' pages
# }

# [6] REFERENCES TO FILES INSIDE THE AUTHOR'S PERSONAL COMPUTER
# 1. The harvester script that gets past the Risat login page
#   -> /ssynthesia/ghostcity/ar/dumper-2/24__2023.02.13__requestsrisat.py

# --------------------------- CODE PREAMBLE --------------------------- #

# Modules import
from lxml import html
from os.path import sep
from tkinter import messagebox
from tkinter import scrolledtext
from tkinter import ttk
from tkinter import IntVar
from tkinter import StringVar
import requests as rq
import tempfile as tmp
import tkinter as tk

# ------------------------ CLASSES DECLARATION ------------------------ #

# Main class declaration
# This class is called first to manage frames and Tkinter instances
class MainGUI(tk.Tk):
    
    # The init function for the MainGUI class
    def __init__(self, *args, **kwargs):
        
        # Instantiating Tkinter instance
        tk.Tk.__init__(self, *args, **kwargs)
        
        # Configuring app's identity
        self.title(APP_NAME)
        
        # Establishing the GUI container
        container = tk.Frame(self)
        container.pack(side='top', fill='both', expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        # Setting up the list of frames
        self.frames = {}
        
        # Initializing the credential variables and harvester arguments
        self.password = ''
        self.username = ''
        self.harvest_category = ''
        self.harvest_datatype = -1
        self.harvest_output = ''
        
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
        about_title = 'TENTANG'
        about_content = 'SiPe.Sat -- Sistem Pemanen Risat Satya Wacana\n\nDibuat oleh Samarthya Lykamanuella\n\nDirektorat Riset dan Pengabdian Masyarakat (DRPM)\n\nUniversitas Kristen Satya Wacana\n\n@ 2023'
        # Showing the info box
        messagebox.showinfo(about_title, about_content)
    
    # The function that will be triggered when the 'LICENSE' button is pressed
    def on_license_button_click(self):
        license_title = 'LISENSI'
        license_content = 'Program ini dilisensi menggunakan lisensi MIT\nKunjungi https://spdx.org/licenses/MIT.html untuk informasi lebih lanjut'
        messagebox.showinfo(license_title, license_content)
    
    # The function that will be triggered when the 'EXIT' button is pressed
    def on_exit_button_click(self):
        # Showing a message box that prompts the user whether to leave the app
        exit_title = 'KELUAR APLIKASI'
        exit_content = 'Apakah Anda yakin untuk keluar dari aplikasi?'
        res = messagebox.askquestion(exit_title, exit_content)
        if res == 'yes':
            # The user confirms to leave the app
            print('[SipesatScrLogin] :: See You :(')
            self.controller.destroy()
        else:
            # Just do nothing
            pass
    
    # The function that will be triggered when the 'LOGIN' form button is pressed
    def on_submit_button_click(self, username, password):
        
        # Validating the input credentials
        validation = BackEndLoginChecker()
        res = validation.validate(username, password)
        
        if res:
            # Login successful
            login_success_title = 'LOGIN BERHASIL'
            login_success_content = f'Anda telah berhasil log masuk.\nSelamat datang, {username}!'
            messagebox.showinfo(login_success_title, login_success_content)
            
            # Upon successful login, store the login credentials inside the 'controller'`s
            # class-wide variable
            self.controller.password = password
            self.controller.username = username
            print('[SipesatScrLogin] :: Credential variables set!')
            
            # Upon successful login, clear the password input
            self.pass_input.delete(0, tk.END)
            
            # Upon successful login, switch to main menu screen
            self.controller.raise_frame(SipesatScrMainMenu)
        else:
            # Login failed
            login_failed_title = 'LOGIN GAGAL'
            login_failed_content = 'Nama pengguna atau kata sandi salah!\n\nMungkin juga disebabkan karena koneksi internet yang bermasalah\n\nSilahkan coba lagi'
            messagebox.showerror(login_failed_title, login_failed_content)
    
    # The __init__ function
    def __init__(self, parent, controller):
        
        # Instantiating tkinter.Frame instance
        tk.Frame.__init__(self, parent)
        
        # Assigning the variables for use by other methods in the class
        self.parent = parent
        self.controller = controller
        
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
        self.pass_input = pass_input # --- for uses by other functions in the class
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
        self.controller.raise_frame(SipesatScrResearch)
    
    # The function that will be triggered when the menu [2] button is selected/clicked
    def on_menu_2_button_click(self):
        self.controller.raise_frame(SipesatScrComService)
    
    # The function that will be triggered when the logout button is selected/clicked
    def on_logout_button_click(self):
        # Showing a message box that prompts the user whehter to logout from the current session
        logout_title = 'LOG OUT'
        logout_content = 'Apakah Anda yakin untuk log keluar dari sesi ini?'
        res = messagebox.askquestion(logout_title, logout_content)
        if res == 'yes':
            # Upon logging out, flush out the credential variables
            self.controller.password = ''
            self.controller.username = ''
            print('[SipesatScrLogin] :: Credential variables flushed out!')
            
            # The user confirms to logout
            self.controller.raise_frame(SipesatScrLogin)
        else:
            # Just do nothing
            pass
    
    # The __init__ function
    def __init__(self, parent, controller):
        
        # Instantiating tkinter.Frame instance
        tk.Frame.__init__(self, parent)
        
        # Assigning the variables for use by other methods in the class
        self.parent = parent
        self.controller = controller
        
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
        
        # Logout button
        logout_text = 'LOG KELUAR'
        logout_button = ttk.Button(self, text=logout_text, width=30,
            command=lambda: self.on_logout_button_click())
        logout_button.grid(row=4, column=0, padx=5, pady=5)

# The harvester menu for 'research'
# - Displayed before the actual harvesting process is executed, for
# preconfiguring what kind of data and which database to be harvested
# - Called upon following the 'SipesatScrMainMenu' screen
class SipesatScrResearch(tk.Frame):
    
    # The function that will be triggered when the 'kembali' button is hit
    def on_back_button_click(self):
        self.controller.raise_frame(SipesatScrMainMenu)
    
    # The function that will be triggered when the 'lanjut' button is hit
    def on_next_button_click(self):
        # Setting the harvest arguments using the controller's 'self' variables
        self.controller.harvest_category = 'r'
        self.controller.harvest_datatype = self.radio_data_type.get()
        self.controller.harvest_output = self.radio_harvest_type.get()
        
        # Calling the harvester screen, begin the harvesting process
        self.controller.raise_frame(SipesatScrHarvest)
    
    # The __init__ function
    def __init__(self, parent, controller):
        
        # The radiobutton variables, used to group different radiobutton together
        self.radio_data_type = IntVar()
        self.radio_harvest_type = StringVar()
        
        # Instantiating tkinter.Frame instance
        tk.Frame.__init__(self, parent)
        
        # Assigning the variables for use by other methods in the class
        self.parent = parent
        self.controller = controller
        
        # The header title displaying the screen's title 
        header_title = ttk.Label(self, text=STRING_HEADER_TITLE, font=FONT_HEADER_TITLE)
        header_title.grid(row=0, column=0, padx=10, pady=2)
        
        # The header description displaying additional information about the screen
        desc_text = 'Menu Utama » Risat Penelitian'
        header_desc = ttk.Label(self, text=desc_text, font=FONT_HEADER_DESC)
        header_desc.grid(row=1, column=0, padx=10, pady=5)
        
        # The screen's display help
        help_text = 'Silahkan tentukan basis data dan jenis data yang akan dipanen\nTekan "LANJUT" untuk memulai proses pemanenan data Risat'
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
        # The 'next' button trigger
        # Clicking this button proceeds the program to the harvester screen
        next_text = 'LANJUT'
        next_button = ttk.Button(layout_frame_3, text=next_text, width=40,
            command=lambda: self.on_next_button_click())
        next_button.grid(row=0, column=1, padx=2, pady=2)

# The harvester menu for 'community service'
# - Displayed before the actual harvesting process is executed, for
# preconfiguring what kind of data and which database to be harvested
# - Called upon following the 'SipesatScrMainMenu' screen
class SipesatScrComService(tk.Frame):
    
    # The function that will be triggered when the 'kembali' button is hit
    def on_back_button_click(self):
        self.controller.raise_frame(SipesatScrMainMenu)
    
    # The function that will be triggered when the 'lanjut' button is hit
    def on_next_button_click(self):
        # Setting the harvest arguments using the controller's 'self' variables
        self.controller.harvest_category = 'c'
        self.controller.harvest_datatype = self.radio_data_type.get()
        self.controller.harvest_output = self.radio_harvest_type.get()
        
        # Calling the harvester screen, begin the harvesting process
        self.controller.raise_frame(SipesatScrHarvest)
    
    # The __init__ function
    def __init__(self, parent, controller):
        
        # The radiobutton variables, used to group different radiobutton together
        self.radio_data_type = IntVar()
        self.radio_harvest_type = StringVar()
        
        # Instantiating tkinter.Frame instance
        tk.Frame.__init__(self, parent)
        
        # Assigning the variables for use by other methods in the class
        self.parent = parent
        self.controller = controller
        
        # The header title displaying the screen's title 
        header_title = ttk.Label(self, text=STRING_HEADER_TITLE, font=FONT_HEADER_TITLE)
        header_title.grid(row=0, column=0, padx=10, pady=2)
        
        # The header description displaying additional information about the screen
        desc_text = 'Menu Utama » Risat Pengabdian Masyarakat'
        header_desc = ttk.Label(self, text=desc_text, font=FONT_HEADER_DESC)
        header_desc.grid(row=1, column=0, padx=10, pady=5)
        
        # The screen's display help
        help_text = 'Silahkan tentukan basis data dan jenis data yang akan dipanen\nTekan "LANJUT" untuk memulai proses pemanenan data Risat'
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
        # The 'next' button trigger
        # Clicking this button proceeds the program to the harvester screen
        next_text = 'LANJUT'
        next_button = ttk.Button(layout_frame_3, text=next_text, width=40,
            command=lambda: self.on_next_button_click())
        next_button.grid(row=0, column=1, padx=2, pady=2)

# The harvester screen
# Displaying the current progress of the harvesting process
class SipesatScrHarvest(tk.Frame):
    
    # The function that will be triggered when the 'cancel' button is selected/clicked
    def on_cancel_button_click(self):
        print('[SipesatScrHarvest] :: Operation cancelled!')
        if self.controller.harvest_category == 'c':
            self.controller.raise_frame(SipesatScrComService)
        elif self.controller.harvest_category == 'r':
            self.controller.raise_frame(SipesatScrResearch)
    
    # The function that will be triggered when the 'cancel' button is selected/clicked
    def on_start_button_click(self):    
        # Calling the back-end harvester class
        harvester = BackEndHarvester()
        # Begin the harvesting process
        harvester.execute_harvester(
            self, # --- the control object, so that this screen's progress bar values,
                  # etc., can be manipulated by the harvest executor
            self.controller.username,
            self.controller.password,
            self.controller.harvest_category,
            self.controller.harvest_datatype,
            self.controller.harvest_output
        )

    # The function that changes the header description of this screen
    def set_header_desc(self, string):
        self.header_desc.config(text=string)

    # The function that changes the help label of this screen
    def set_help_label(self, string):
        self.help_label.config(text=string)

    # The function that changes the progress label of this screen
    # i.e., changes the percentage display of the progress bar
    def set_progress_label(self, string):
        self.progress_label.config(text=string)

    # The function that changes the progress bar progression of this screen
    def set_progress_bar(self, value):
        # The value of the variable 'value' must be
        # in the range of 0 - 100
        if value > 100:
            value = 100
        elif value < 0:
            value = 0

        # Begin setting the progress bar's progression value
        self.progress_bar['value'] = value

    # The function that clears this screen's message area's harvest logging output
    def clear_message_area(self):
        self.message_area.delete(1.0, tk.END)

    # The function that sets (i.e., clear and write) this screen's message area's
    # harvest logging output
    def set_message_area(self, long_string):
        self.clear_message_area() # --- clearing the content first
        self.message_area.insert(tk.INSERT, long_string)

    # The function that appends long string to this screen's message area's
    # harvest logging output
    # New line character is concatenated in between the existing message area's
    # content and the string to be appended, by default
    def append_message_area(self, long_string, new_line=True):
        # Getting the current content of the message area
        str_ = self.message_area.get('1.0', tk.END)
        # Concat with a new line character,
        # if specified by the argument
        if new_line:
            str_ += '\n'
        # Appending 'long_string' to the existing content,
        # and then applying to the message area GUI element
        str_ += long_string
        self.set_message_area(str_)
    
    # The __init__ function
    def __init__(self, parent, controller):
        
        # -------------------- GUI LAYOUT -------------------- #

        # List of GUI elements that can be accessed by
        # other (non-'__init__') functions of this class
        #
        # self.header_desc      --> the header description
        # self.help_label       --> the label that indicates the status of the harvest process
        # self.progress_label   --> to display the percentage of the progress bar
        # self.progress_bar     --> the progress bar that indicates the progression of the harvest
        # self.message_area     --> the displayer of the harvest log
        
        # Instantiating tkinter.Frame instance
        tk.Frame.__init__(self, parent)
        
        # Assigning the variables for use by other methods in the class
        self.parent = parent
        self.controller = controller
        
        # The header title displaying the screen's title 
        header_title = ttk.Label(self, text=STRING_HEADER_TITLE, font=FONT_HEADER_TITLE)
        header_title.grid(row=0, column=0, padx=10, pady=2)
        
        # The header description displaying additional information about the screen
        desc_text = 'Panen Data "[...placeholder...]"'
        self.header_desc = ttk.Label(self, text=desc_text, font=FONT_HEADER_DESC)
        self.header_desc.grid(row=1, column=0, padx=10, pady=5)
        
        # The screen's display help
        help_text = 'Data "[...placeholder...]" sedang dipanen. Silahkan menunggu...'
        self.help_label = ttk.Label(self, text=help_text, font=FONT_REGULAR, anchor='w', justify='left')
        self.help_label.grid(row=2, column=0, padx=10, pady=10)
        
        # :::
        # Layout harvester progress bar display
        progress_frame = tk.Frame(self)
        progress_frame.grid(row=3, column=0, padx=10, pady=10)
        # The progress status/value display
        progress_text = 'Status: 67%'
        self.progress_label = ttk.Label(progress_frame, text=progress_text, font=FONT_PROGRESS_VALUE, anchor='center', justify='center')
        self.progress_label.grid(row=0, column=0, padx=2, pady=5)
        # The progress bar
        self.progress_bar = ttk.Progressbar(progress_frame, length=600)
        self.progress_bar['value'] = 20
        self.progress_bar.grid(row=1, column=0, padx=2, pady=2, sticky='we')
        
        # The log message displayer
        self.message_area = scrolledtext.ScrolledText(self, width=100, height=10)
        self.message_area.grid(row=4, column=0, padx=5, pady=5)
        
        # :::
        # Layout: The action buttons
        layout_actions = tk.Frame(self)
        layout_actions.grid(row=5, column=0, padx=10, pady=10)
        
        # Cancel button
        # - For debug purposes only: to allow testing out switching between screens
        #   in an efficient manner
        # - Upon release, this button should be disabled
        cancel_text = 'BATALKAN'
        cancel_button = ttk.Button(layout_actions, text=cancel_text, width=30,
            command=lambda: self.on_cancel_button_click())
        cancel_button.grid(row=0, column=0, padx=2, pady=2)
        # Setting the 'disabled' state of the button
        # Possible values: 'normal' and 'disabled'
        cancel_button['state'] = 'normal'
        
        # The 'start' button trigger
        # Clicking this button proceeds the program to begin the harvesting process
        start_text = 'MULAI PANEN'
        start_button = ttk.Button(layout_actions, text=start_text, width=40,
            command=lambda: self.on_start_button_click())
        start_button.grid(row=0, column=1, padx=2, pady=2)
    
# - The class that checks whether the input username/password credentials in
#   'SipesatScrLogin' screen are correct
class BackEndLoginChecker():
    
    # The __init__ function
    def __init__(self, *args, **kwargs):
        
        # Instantiating 'requests.Session'
        self.session = rq.Session()
                
        # Initializing variables
        self.password = ''
        self.username = ''
    
    # The variable that checks the validity of the input Risat credentials
    def validate(self, username, password):
        
        # Setting the variables according to the passed arguments
        self.password = password
        self.username = username
        
        # Opening the Risat homepage
        risat_homepage = 'https://risat.uksw.edu/login.aspx?ReturnUrl=%2f'
        r = self.session.get(risat_homepage)
        content = html.fromstring(r.content)
        
        # Obtaining the computer-generated hidden values of ASPX (before login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]
        
        # Preparing the http handler URL and payload
        HANDLER_URL = 'https://risat.uksw.edu/login.aspx?ReturnUrl=%2f'
        PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : viewstate,
            '__VIEWSTATEGENERATOR' : viewstategen,
            '__EVENTVALIDATION' : eventvalidation,
            # The values below are the login information provided by the prompt in the previous code block
            'txnip1': username,
            'txpwd1': password,
            'btlogin1': True
        }
        
        # Logging in to Risat administrator page
        post = self.session.post(HANDLER_URL, data=PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code
        
        # Finding the existence of HTML elements that exist when the input
        # login credentials are wrong
        try:
            prooftest = content.xpath('//div[@class="card-body"]/div[@class="form-group f12"]/text()')[0]
            if prooftest.__contains__('User atau password anda salah'):
                # The input credentials are wrong
                return False
        except IndexError:
            # The element is not present in the HTML page,
            # meaning the login credentials must be correct
            return True
        
        # The default state
        return False

# The essence of this application: the harvester class
# This class sorts harvest input arguments (such as data category,
# data type, and data output) and then execute harvesting the Risat data
class BackEndHarvester():
    
    # The __init__ function
    # This function instantiates 'requests.Session'
    # It does nothing else
    def __init__(self, *args, **kwargs):
        
        # Instantiating 'requests.Session'
        self.session = rq.Session()

        # Preparing the temporary folder that will be used
        # in the recursive scraper of the detail pages
        self.tmpdir = tmp.mkdtemp(prefix=SCRAPE_TEMP_DIR_PREFIX) + '/'
    
    # The harvest executor
    # Specify the username and password credentials to mitigate session timeout
    # Please refer to convention [3] for the possible values of 'category', 'datatype', and 'output'
    #
    # In the function 'execute_harvester()' below,
    # the argument 'control' refers to the SipesatScrHarvest
    # harvesting screen/GUI, which progress bar, labels, and
    # message area will be manipulated by this harvest executor
    # function as the harvesting process progresses
    def execute_harvester(self, control, username, password, category, datatype, output):
        
        # :::
        # Determining the cases of the harvesting arguments passed
        # This switching-cases require Python version >= v3.10
        
        # --------------------- RISAT PENELITIAN --------------------- #
        if category == 'r': # --- category selected: 'Risat Penelitian'
            
            # Determining the cases of the datatype
            if datatype == 0: # --- 'data ringkasan'
            
                # Determining the cases of the data output
                match output:
                    case 'arsip':
                        print(f'[BackEndHarvester] :: SELECTED_CASE: category={category}, datatype={datatype}, output={output}')
                        pass
                    case 'dana':
                        print(f'[BackEndHarvester] :: SELECTED_CASE: category={category}, datatype={datatype}, output={output}')
                        pass
                    case 'lapakhir':
                        print(f'[BackEndHarvester] :: SELECTED_CASE: category={category}, datatype={datatype}, output={output}')
                        pass
                        
            elif datatype == 1: # --- 'data detil lengkap'
                
                # Determining the cases of the data output
                match output:
                    case 'arsip':
                        print(f'[BackEndHarvester] :: SELECTED_CASE: category={category}, datatype={datatype}, output={output}')
                        pass
                    case 'dana':
                        print(f'[BackEndHarvester] :: SELECTED_CASE: category={category}, datatype={datatype}, output={output}')
                        pass
                    case 'lapakhir':
                        print(f'[BackEndHarvester] :: SELECTED_CASE: category={category}, datatype={datatype}, output={output}')
                        pass
        
        # ---------------- RISAT PENGABDIAN MASYARAKAT ---------------- #
        elif category == 'c': # --- category selected: 'Pengabdian Masyarakat'
            
            # Determining the cases of the datatype
            if datatype == 0: # --- 'data ringkasan'
            
                # Determining the cases of the data output
                match output:
                    case 'arsip':
                        print(f'[BackEndHarvester] :: SELECTED_CASE: category={category}, datatype={datatype}, output={output}')
                        pass
                    case 'dana':
                        print(f'[BackEndHarvester] :: SELECTED_CASE: category={category}, datatype={datatype}, output={output}')
                        pass
                    case 'lapakhir':
                        print(f'[BackEndHarvester] :: SELECTED_CASE: category={category}, datatype={datatype}, output={output}')
                        pass
                        
            elif datatype == 1: # --- 'data detil lengkap'
                
                # Determining the cases of the data output
                match output:
                    case 'arsip':
                        print(f'[BackEndHarvester] :: SELECTED_CASE: category={category}, datatype={datatype}, output={output}')
                        pass
                    case 'dana':
                        print(f'[BackEndHarvester] :: SELECTED_CASE: category={category}, datatype={datatype}, output={output}')
                        pass
                    case 'lapakhir':
                        print(f'[BackEndHarvester] :: SELECTED_CASE: category={category}, datatype={datatype}, output={output}')
                        pass
        
        # ======================== END ======================== #

    # This function recursively obtains the detail pages of Risat
    # For example, the 'Dana Penelitian Detil' and 'Abdimas Arsip Detil'
    # Requires no 'data_prompt' passed as argument,
    # but 'mode' argument is mandated to be passed
    #
    # Returns an array which elements are the paths to the
    # temporary files where each detail pages are stored into
    #
    # The possible values of 'mode' argument are as follows:
    # mode='dana_penelitian'        --> obtains the detail pages of 'Risat Dana Penelitian'
    # mode='arsip_penelitian'        --> obtains the detail pages of 'Risat Arsip Penelitian'
    # mode='dana_pengabdian'        --> obtains the detail pages of 'Risat Dana Pengabdian'
    # mode='arsip_pengabdian'        --> obtains the detail pages of 'Risat Arsip Pengabdian'
    def get_auto_risat_detil(self, mode):

        # :::
        # Determining the mode of the risat detil pages to be scraped recursively

        # --------------------- DANA PENELITIAN --------------------- #
        if mode == 'dana_penelitian':  # --- category selected: 'Dana Penelitian'

            # Preparing the 'data_prompt' arrays
            data_prompt = self.get_risat_login()
            data_prompt = self.get_risat_penelitian(data_prompt)
            data_prompt = self.get_risat_penelitian_dana_penelitian(data_prompt)
            
            # Parsing XML tree content
            content = data_prompt['html_content']
            
            # The base XPath location, pointing to each entry row
            base = '//div[@id="ContentPlaceHolder1_upd1"]/div[@class="mw-100"]//table[@width="100%"]//tr[@valign="top"]'
            
            # Reading the HTML entry row hidden ASPX values
            # These variables below are *arrays*, so they have indices
            all_kodetran_prop = content.xpath(base + '//input[1][@type="hidden"]/@name')
            all_kodetran_val = content.xpath(base + '//input[1][@type="hidden"]/@value')
            all_stat_prop = content.xpath(base + '//input[2][@type="hidden"]/@name')
            all_stat_val = content.xpath(base + '//input[2][@type="hidden"]/@value')
            all_submitbtn = content.xpath(base + '//input[@type="submit"]/@name')

            # The temporary 'data_prompt' to be used to revert back to the table page
            # which lists the detil pages
            temporary_prompt = data_prompt

            # The array which stores the list of temporary files
            # that stores the HTTP response of each detail page scraped
            scrape_array = []

            # Iterating through each entry row element
            # Assumes all the arrays in the previous code block
            # are of the same length/size
            for i in range(len(all_kodetran_prop)-1):

                # Preparing the AJAX payload
                detail_prompt = {
                    'viewstate' : temporary_prompt['viewstate'],
                    'viewstategen' : temporary_prompt['viewstategen'],
                    'eventvalidation' : temporary_prompt['eventvalidation'],
                    'button_name' : all_submitbtn[i],
                    'kodetran_prop' : all_kodetran_prop[i],
                    'kodetran_val' : all_kodetran_val[i],      
                    'stat_prop' : all_stat_prop[i],
                    'stat_val' : all_stat_val[i]
                }
                
                # Obtaining the response data of each individual entry row detail page
                data = self.get_risat_penelitian_dana_penelitian_detil(detail_prompt)
                content = data['html_content']
                response = data['http_response']

                # Writing to the temporary file in the temporary directory
                tmppath = self.tmpdir + sep + 'dana_penelitian-detil-' + str(i) + '.html'
                fo = open(tmppath, 'w')
                fo.write(response)
                fo.close()

                # Appending to the array
                scrape_array.insert(i, response)
                
                # Logging: printing the basic information
                judul = content.xpath('//span[@id="ContentPlaceHolder1_danacair1_kiri1_txjudul1"]/text()')[0].replace(
                    '\r', '').replace('\n', '').strip()
                print(f'ENTRY_NO: {i} --> {judul}')
                
                # Reopening the "Dana Penelitian" list page,
                # then assign the AJAX response to the temporary array 'temporary_prompt'
                # The 'data' array is obtained from opening individual entry row detail page
                temporary_prompt = self.get_risat_penelitian_dana_penelitian(data)
                continue

            # Returning the list of array
            return scrape_array

        # --------------------- ARSIP PENELITIAN --------------------- #
        elif mode == 'arsip_penelitian':  # --- category selected: 'Arsip Penelitian'
            pass

        # ---------------------- DANA PENGABDIAN ---------------------- #
        elif mode == 'dana_pengabdian':  # --- category selected: 'Dana Pengabdian'
            pass

        # ---------------------- ARSIP PENGABDIAN ---------------------- #
        elif mode == 'arsip_pengabdian':  # --- category selected: 'Arsip Pengabdian'
            pass

        # ======================== END ======================== #

    # This function gets past the barrier of Risat login page
    # - Requires two arguments: the username and password credentials
    # - Returns 'data_prompt' array
    def get_risat_login(self, username, password):
        # Opening the Risat homepage
        print('+ Opening Risat homepage...')
        risat_homepage = 'https://risat.uksw.edu/login.aspx?ReturnUrl=%2f'
        r = self.session.get(risat_homepage)
        content = html.fromstring(r.content)
        
        # Obtaining the computer-generated hidden values of ASPX (before login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]
        
        # Preparing the http handler URL and payload
        HANDLER_URL = 'https://risat.uksw.edu/login.aspx?ReturnUrl=%2f'
        PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : viewstate,
            '__VIEWSTATEGENERATOR' : viewstategen,
            '__EVENTVALIDATION' : eventvalidation,
            # The values below are the login information provided by the prompt in the previous code block
            'txnip1': username,
            'txpwd1': password,
            'btlogin1': True
        }
        
        # Logging in to Risat administrator page
        print('+ Logging in...')
        post = self.session.post(HANDLER_URL, data=PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code
        
        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]
        
        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }
        
        # Returning the http response string
        return data_prompt

    # This function opens "Penelitian" menu tab
    # Requires 'data_prompt' array obtained from the login function
    # Returns another 'data_prompt' array
    def get_risat_penelitian(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Penelitian" menu...')
        
        # Preparing the http handler URL and payload
        HANDLER_URL = 'https://risat.uksw.edu/bp3mpage.aspx'
        PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            'ctl00$menu2': True
        }
        
        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(HANDLER_URL, data=PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code
        
        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]
        
        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }
        
        # Returning the http response string
        return data_prompt
    
    # This function opens "Dana Penelitian" menu after opening the tab "Penelitian"
    # - Requires 'data_prompt' array as an unary argument obtained from get_risat_penelitian() function
    # - Returns also another 'data_prompt' array
    def get_risat_penelitian_dana_penelitian(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Penelitian --> Dana Penelitian" menu...')
        
        # Preparing the http handler URL and payload
        HANDLER_URL = 'https://risat.uksw.edu/bp3mpage.aspx'
        PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            '__EVENTTARGET' : 'ctl00$ContentPlaceHolder1$menu10'
        }
        
        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(HANDLER_URL, data=PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code
        
        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]
        
        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }
        
        # Returning 'data_prompt' array
        return data_prompt

    # This function opens the detail page of "Dana Penelitian" entry row
    # Requires one argument:
    # - the data prompt value array in compliance with convention [5] of this file
    # The 'data_prompt' argument passed into this function must be
    # rooted from the function 'get_risat_penelitian_dana_penelitian()'
    def get_risat_penelitian_dana_penelitian_detil(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Dana Penelitian" entry row detail page...')
        
        # Preparing the http handler URL and payload
        LOGIN_HANDLER_URL = 'https://risat.uksw.edu/bp3mpage.aspx'
        LOGIN_PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            # The entry row's submit button to 'hit'
            data_prompt['button_name'] : 'Detil',
            data_prompt['kodetran_prop'] : data_prompt['kodetran_val'],
            data_prompt['stat_prop'] : data_prompt['stat_val']
        }
        
        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(LOGIN_HANDLER_URL, data=LOGIN_PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code
        
        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]
        
        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }
        
        # Returning the http response string
        return data_prompt

    # This function opens "Pelaksanaan Kegiatan" menu after opening the tab "Penelitian"
    # - Requires 'data_prompt' array as an unary argument obtained from get_risat_penelitian() function
    # - Returns also another 'data_prompt' array
    def get_risat_penelitian_pelak_kegi(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Penelitian --> Pelaksanaan Kegiatan" menu...')

        # Preparing the http handler URL and payload
        HANDLER_URL = 'https://risat.uksw.edu/bp3mpage.aspx'
        PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            '__EVENTTARGET' : 'ctl00$ContentPlaceHolder1$menu9'
        }

        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(HANDLER_URL, data=PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code

        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]

        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }

        # Returning 'data_prompt' array
        return data_prompt

    # This function opens "Laporan Akhir" menu after opening the tab "Penelitian -> Pelaksanaan Kegiatan"
    # - Requires 'data_prompt' array as an unary argument obtained from get_risat_penelitian_pelak_kegi() function
    # - Returns also another 'data_prompt' array
    def get_risat_penelitian_pelak_kegi_lapakhir(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Penelitian --> Pelaksanaan Kegiatan --> Laporan Akhir" menu...')

        # Preparing the http handler URL and payload
        HANDLER_URL = 'https://risat.uksw.edu/bp3mpage.aspx'
        PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            '__EVENTTARGET' : 'ctl00$ContentPlaceHolder1$submenu3'
        }

        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(HANDLER_URL, data=PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code

        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]

        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }

        # Returning 'data_prompt' array
        return data_prompt

    # This function opens "Arsip" menu after opening the tab "Penelitian"
    # - Requires 'data_prompt' array as an unary argument obtained from get_risat_penelitian() function
    # - Returns also another 'data_prompt' array
    def get_risat_penelitian_arsip(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Penelitian --> Arsip" menu...')

        # Preparing the http handler URL and payload
        HANDLER_URL = 'https://risat.uksw.edu/bp3mpage.aspx'
        PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            '__EVENTTARGET' : 'ctl00$ContentPlaceHolder1$menu11'
        }

        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(HANDLER_URL, data=PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code

        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]

        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }

        # Returning 'data_prompt' array
        return data_prompt

    # This function opens the detail page of "Arsip" entry row
    # Requires one argument:
    # - the data prompt value array in compliance with convention [5] of this file
    # The 'data_prompt' argument passed into this function must be
    # rooted from the function 'get_risat_penelitian_arsip()'
    def get_risat_penelitian_arsip_detil(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Arsip" entry row detail page...')

        # Preparing the http handler URL and payload
        LOGIN_HANDLER_URL = 'https://risat.uksw.edu/bp3mpage.aspx'
        LOGIN_PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            # The entry row's submit button to 'hit'
            data_prompt['button_name'] : 'Detil',
            data_prompt['idat_prop'] : data_prompt['idat_val'],
            data_prompt['itgl_prop'] : data_prompt['itgl_val']
        }

        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(LOGIN_HANDLER_URL, data=LOGIN_PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code

        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]

        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }

        # Returning the http response string
        return data_prompt

    # This function opens "Pengabdian" menu tab
    # Requires 'data_prompt' array obtained from the login function
    # Returns another 'data_prompt' array
    def get_risat_pengabdian(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Pengabdian" menu...')

        # Preparing the http handler URL and payload
        HANDLER_URL = 'https://risat.uksw.edu/bp3mpageabdimas.aspx'
        PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            'ctl00$menu3': True
        }

        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(HANDLER_URL, data=PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code

        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]

        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }

        # Returning the http response string
        return data_prompt

    # This function opens "Dana Pengabdian" menu after opening the tab "Pengabdian"
    # - Requires 'data_prompt' array as an unary argument obtained from get_risat_pengabdian() function
    # - Returns also another 'data_prompt' array
    def get_risat_pengabdian_dana_pengabdian(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Pengabdian --> Dana Pengabdian" menu...')

        # Preparing the http handler URL and payload
        HANDLER_URL = 'https://risat.uksw.edu/bp3mpageabdimas.aspx'
        PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            '__EVENTTARGET' : 'ctl00$ContentPlaceHolder1$menu10'
        }

        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(HANDLER_URL, data=PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code

        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]

        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }

        # Returning 'data_prompt' array
        return data_prompt

    # This function opens the detail page of "Dana Pengabdian" entry row
    # Requires one argument:
    # - the data prompt value array in compliance with convention [5] of this file
    # The 'data_prompt' argument passed into this function must be
    # rooted from the function 'get_risat_pengabdian_dana_pengabdian()'
    def get_risat_pengabdian_dana_pengabdian_detil(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Dana Pengabdian" entry row detail page...')

        # Preparing the http handler URL and payload
        LOGIN_HANDLER_URL = 'https://risat.uksw.edu/bp3mpageabdimas.aspx'
        LOGIN_PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            # The entry row's submit button to 'hit'
            data_prompt['button_name'] : 'Detil',
            data_prompt['kodetran_prop'] : data_prompt['kodetran_val'],
            data_prompt['stat_prop'] : data_prompt['stat_val']
        }

        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(LOGIN_HANDLER_URL, data=LOGIN_PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code

        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]

        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }

        # Returning the http response string
        return data_prompt

    # This function opens "Pelaksanaan Kegiatan" menu after opening the tab "Pengabdian"
    # - Requires 'data_prompt' array as an unary argument obtained from get_risat_pengabdian() function
    # - Returns also another 'data_prompt' array
    def get_risat_pengabdian_pelak_kegi(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Pengabdian --> Pelaksanaan Kegiatan" menu...')

        # Preparing the http handler URL and payload
        HANDLER_URL = 'https://risat.uksw.edu/bp3mpageabdimas.aspx'
        PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            '__EVENTTARGET' : 'ctl00$ContentPlaceHolder1$menu9'
        }

        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(HANDLER_URL, data=PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code

        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]

        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }

        # Returning 'data_prompt' array
        return data_prompt

    # This function opens "Laporan Akhir" menu after opening the tab "Pengabdian -> Pelaksanaan Kegiatan"
    # - Requires 'data_prompt' array as an unary argument obtained from get_risat_pengabdian_pelak_kegi() function
    # - Returns also another 'data_prompt' array
    def get_risat_pengabdian_pelak_kegi_lapakhir(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Pengabdian --> Pelaksanaan Kegiatan --> Laporan Akhir" menu...')

        # Preparing the http handler URL and payload
        HANDLER_URL = 'https://risat.uksw.edu/bp3mpageabdimas.aspx'
        PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            '__EVENTTARGET' : 'ctl00$ContentPlaceHolder1$submenu3'
        }

        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(HANDLER_URL, data=PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code

        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]

        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }

        # Returning 'data_prompt' array
        return data_prompt

    # This function opens "Arsip" menu after opening the tab "Pengabdian"
    # - Requires 'data_prompt' array as an unary argument obtained from get_risat_pengabdian() function
    # - Returns also another 'data_prompt' array
    def get_risat_pengabdian_arsip(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Pengabdian --> Arsip" menu...')

        # Preparing the http handler URL and payload
        HANDLER_URL = 'https://risat.uksw.edu/bp3mpageabdimas.aspx'
        PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            '__EVENTTARGET' : 'ctl00$ContentPlaceHolder1$menu11'
        }

        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(HANDLER_URL, data=PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code

        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]

        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }

        # Returning 'data_prompt' array
        return data_prompt

    # This function opens the detail page of "Arsip" entry row
    # Requires one argument:
    # - the data prompt value array in compliance with convention [5] of this file
    # The 'data_prompt' argument passed into this function must be
    # rooted from the function 'get_risat_pengabdian_arsip()'
    def get_risat_pengabdian_arsip_detil(self, data_prompt):
        # Logging the calling of the function
        print('+ Opening "Arsip" entry row detail page...')

        # Preparing the http handler URL and payload
        LOGIN_HANDLER_URL = 'https://risat.uksw.edu/bp3mpageabdimas.aspx'
        LOGIN_PAYLOAD = {
            # The values below are computer-generated
            '__VIEWSTATE' : data_prompt['viewstate'],
            '__VIEWSTATEGENERATOR' : data_prompt['viewstategen'],
            '__EVENTVALIDATION' : data_prompt['eventvalidation'],
            # The entry row's submit button to 'hit'
            data_prompt['button_name'] : 'Detil',
            data_prompt['idat_prop'] : data_prompt['idat_val'],
            data_prompt['itgl_prop'] : data_prompt['itgl_val']
        }

        # Posting the http payloads
        print('+ Posting http payloads...')
        post = self.session.post(LOGIN_HANDLER_URL, data=LOGIN_PAYLOAD)
        response = post.text # --- Obtaining the response text
        content = html.fromstring(response) # --- Scraping the HTML code

        # Obtaining the computer-generated hidden values of ASPX (after login)
        viewstate = content.xpath('//*[@id="__VIEWSTATE"]/@value')[0]
        viewstategen = content.xpath('//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        eventvalidation = content.xpath('//*[@id="__EVENTVALIDATION"]/@value')[0]

        # Building the 'data_prompt' array
        data_prompt = {
            'http_response' : response,
            'html_content' : content,
            'viewstate' : viewstate,
            'viewstategen' : viewstategen,
            'eventvalidation' : eventvalidation
        }

        # Returning the http response string
        return data_prompt

# -------------------------- CONSTANT PRESETS -------------------------- #

# 'FRAME_CLASSES' is a tuple that defines all the frame classes of the file
# - The first class specified in this tuple will be the frame displayed
#   at the very beginning of the application after launch
# - Need to have at least two classes, otherwise the following error will be casted:
#   TypeError: 'type' object is not iterable
FRAME_CLASSES = (SipesatScrLogin, SipesatScrMainMenu, SipesatScrComService, SipesatScrResearch, SipesatScrHarvest,)

# The following constants define font presets used in styling Tkinter widgets
FONT_HEADER_TITLE = ('Segoe UI', 30, 'bold')
FONT_HEADER_DESC = ('Segoe UI', 20, 'italic')
FONT_REGULAR = ('Segoe UI', 12)
FONT_FORM_INPUT = ('Courier New', 10, 'bold')
FONT_RADIO = ('Courier New', 10)
FONT_PROGRESS_VALUE = ('Courier New', 14)

# The following constants define the string presets used as template and localization
STRING_HEADER_TITLE = 'SiPe.Sat'
STRING_HEADER_DESC = 'Sistem Pemanen Risat'

# Constants that define application identity
APP_NAME = 'Sistem Pemanen Risat - UKSW 2023'

# Back-End constants that specify the temporary folder name prefix
# in which the recursively scraped 'detail' pages are stored into
SCRAPE_TEMP_DIR_PREFIX = 'sipesat-123-'

# -------------------------- DEVELOPMENT TEST -------------------------- #

# Modules import (for development purposes only)
from getpass import getpass as input_pass

# The development test class
class DevelopmentTest():

    # On 2023-03-28
    # Testing out opening the detail page of 'penelitian arsip' entry row
    def experiment_1(self):
        # Preamble
        print('+ Running experiment_1: Opening "Penelitian Arsip Detil" page ...')

        # Initializing the harvester back-end
        backend = BackEndHarvester()

        # Opening the 'Penelitian Arsip' page
        data_prompt = backend.get_risat_login(self.user_, self.pass_)
        data_prompt = backend.get_risat_penelitian(data_prompt)
        data_prompt = backend.get_risat_penelitian_arsip(data_prompt)

        # Setting of sample detail page
        data_prompt['button_name'] = 'ctl00$ContentPlaceHolder1$arsip1$repusul1$ctl01$btdetil1'
        data_prompt['idat_prop'] = 'ctl00$ContentPlaceHolder1$arsip1$repusul1$ctl01$idat1'
        data_prompt['idat_val'] = '12AAFE6D-7C03-4998-9CB6-1BF4ECA37831'
        data_prompt['itgl_prop'] = 'ctl00$ContentPlaceHolder1$arsip1$repusul1$ctl01$itgl1'
        data_prompt['itgl_val'] = '2021/01/21'

        # Opening the detail page of 'Penelitian Arsip'
        data_prompt = backend.get_risat_penelitian_arsip_detil(data_prompt)

        # Writing the HTTP response to an external file
        print('+ Dumping the HTTP response ...')
        fo = open(HTTP_RESPONSE_DUMPER, 'w')
        fo.write(str(data_prompt['http_response']))

    # On 2023-03-28
    # Testing scraping the 'Penelitian > Pelaksanaan Kegiatan > Lap. Akhir' page
    def experiment_2(self):
        # Preamble
        print('+ Running experiment_2: Opening "Penelitian Lap. Akhir" page ...')

        # Initializing the harvester back-end
        backend = BackEndHarvester()

        # Opening the 'Penelitian Lap. Akhir' page
        data_prompt = backend.get_risat_login(self.user_, self.pass_)
        data_prompt = backend.get_risat_penelitian(data_prompt)
        data_prompt = backend.get_risat_penelitian_pelak_kegi(data_prompt)
        data_prompt = backend.get_risat_penelitian_pelak_kegi_lapakhir(data_prompt)

        # Writing the the HTTP response to an external file
        print('+ Dumping the HTTP response ...')
        fo = open(HTTP_RESPONSE_DUMPER, 'w')
        fo.write(str(data_prompt['http_response']))

    # On 2023-03-28
    # Testing scraping the 'Pengabdian > Pelaksanaan Kegiatan > Lap. Akhir' page
    def experiment_3(self):
        # Preamble
        print('+ Running experiment_3: Opening "Pengabdian Lap. Akhir" page ...')

        # Initializing the harvester back-end
        backend = BackEndHarvester()

        # Opening the 'Pengabdian Lap. Akhir' page
        data_prompt = backend.get_risat_login(self.user_, self.pass_)
        data_prompt = backend.get_risat_pengabdian(data_prompt)
        data_prompt = backend.get_risat_pengabdian_pelak_kegi(data_prompt)
        data_prompt = backend.get_risat_pengabdian_pelak_kegi_lapakhir(data_prompt)

        # Writing the the HTTP response to an external file
        print('+ Dumping the HTTP response ...')
        fo = open(HTTP_RESPONSE_DUMPER, 'w')
        fo.write(str(data_prompt['http_response']))

    # On 2023-03-28
    # Testing out opening the detail page of 'pengabdian dana pengabdian' entry row
    def experiment_4(self):
        # Preamble
        print('+ Running experiment_4: Opening "Pengabdian Dana Pengabdian Detil" page ...')

        # Initializing the harvester back-end
        backend = BackEndHarvester()

        # Opening the 'Pengabdian Dana Pengabdian' page
        data_prompt = backend.get_risat_login(self.user_, self.pass_)
        data_prompt = backend.get_risat_pengabdian(data_prompt)
        data_prompt = backend.get_risat_pengabdian_dana_pengabdian(data_prompt)

        # Setting of sample detail page
        data_prompt['button_name'] = 'ctl00$ContentPlaceHolder1$danacair1$repusulan1$ctl01$btdetil1'
        data_prompt['kodetran_prop'] = 'ctl00$ContentPlaceHolder1$danacair1$repusulan1$ctl01$kodetran1'
        data_prompt['kodetran_val'] = 'D480D002-80B3-42A4-919F-08EDD11F62D5'
        data_prompt['stat_prop'] = 'ctl00$ContentPlaceHolder1$danacair1$repusulan1$ctl01$stat1'
        data_prompt['stat_val'] = 'C'

        # Opening the detail page of 'Pengabdian Dana Pengabdian'
        data_prompt = backend.get_risat_pengabdian_dana_pengabdian_detil(data_prompt)

        # Writing the HTTP response to an external file
        print('+ Dumping the HTTP response ...')
        fo = open(HTTP_RESPONSE_DUMPER, 'w')
        fo.write(str(data_prompt['http_response']))

    # On 2023-03-28
    # Testing out opening the detail page of 'pengabdian arsip' entry row
    def experiment_5(self):
        # Preamble
        print('+ Running experiment_5: Opening "Pengabdian Arsip Detil" page ...')

        # Initializing the harvester back-end
        backend = BackEndHarvester()

        # Opening the 'Pengabdian Arsip' page
        data_prompt = backend.get_risat_login(self.user_, self.pass_)
        data_prompt = backend.get_risat_pengabdian(data_prompt)
        data_prompt = backend.get_risat_pengabdian_arsip(data_prompt)

        # Setting of sample detail page
        data_prompt['button_name'] = 'ctl00$ContentPlaceHolder1$arsip1$repusul1$ctl01$btdetil1'
        data_prompt['idat_prop'] = 'ctl00$ContentPlaceHolder1$arsip1$repusul1$ctl01$idat1'
        data_prompt['idat_val'] = 'EB969B69-8F86-4790-995F-3833400C42D8'
        data_prompt['itgl_prop'] = 'ctl00$ContentPlaceHolder1$arsip1$repusul1$ctl01$itgl1'
        data_prompt['itgl_val'] = '2021/02/09'

        # Opening the detail page of 'Penelitian Arsip'
        data_prompt = backend.get_risat_pengabdian_arsip_detil(data_prompt)

        # Writing the HTTP response to an external file
        print('+ Dumping the HTTP response ...')
        fo = open(HTTP_RESPONSE_DUMPER, 'w')
        fo.write(str(data_prompt['http_response']))

    # On 2023-03-28
    # Testing out running the recursive autoscraper of 'Dana Penelitian'
    def experiment_6(self):
        # Preamble
        print('+ Running experiment_6: Running the recursive autoscraper of "Dana Penelitian" ...')

        # Initializing the harvester back-end
        backend = BackEndHarvester()

        # Getting the array of files of the scraped detail pages
        array_of_files = backend.get_auto_risat_detil('dana_penelitian')

        # Printing the array
        print('+ Printing the array ...')
        print(array_of_files)

    # Wrapper of the developmental tester
    def launch_experiment_wrapper(self):

        # Processing the request
        # This switching-cases require Python version >= v3.10
        match self.nmbr_:
            case 1:
                self.experiment_1()
            case 2:
                self.experiment_2()
            case 3:
                self.experiment_3()
            case 4:
                self.experiment_4()
            case 5:
                self.experiment_5()
            case 6:
                self.experiment_6()
            case _:
                print('+ Error! Command not available!')

        # At the end of the experiment test, exit the app
        # so that the main GUI won't load
        print('+ Exitting the development test ...')
        exit()

    # The '__init__' function
    def __init__(self):
        # Preamble
        print('+ Beginning the development test ...')

        # Prompting username and password
        self.user_ = input('Please enter your username\n >>> ')
        self.pass_ = input_pass(f'Enter the password for [{self.user_}]\n >>> ')
        self.nmbr_ = int(input('Please enter the experiment number (integer)\n >>> '))

# Constants used in development tests only
HTTP_RESPONSE_DUMPER = '/tmp/http_dumper.html'

# ------------------------- APPLICATION LAUNCH ------------------------- #

# Development testing
# Should be commented out by final release to the public
dev = DevelopmentTest()
dev.launch_experiment_wrapper()

# Initializing the GUI
app_gui = MainGUI()
app_gui.mainloop()
