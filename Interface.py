import datetime as dt
import logging
import os
import shutil
import subprocess
import threading as thread
import time
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import tkinter.simpledialog as sd
import tkinter.ttk as tkk
import webbrowser
from tkinter import *

from Registry import Registry

import Parse
import Reports
import Verification


class UserInterface:

    def __init__(self, master):
        self.master = master
        self.report_log = ""
        self.timeline_log = []
        self.investigator = "RegSmart"
        self.intro()

        # Declarations
        self.directory = ""
        self.stop_processing = False
        self.finished_parse = [False, False, False, False, False]
        self.current_user = None
        self.current_config = None
        self.local_machine = None
        self.users = None
        self.session = ""

        self.session_name = StringVar()

        self.full_session = ""
        self.location = ""
        self.business = ""
        self.hash_gui = ""
        self.has_report = ""

        self.business_setting = ""
        self.location_setting = ""
        self.settings = ""
        self.db = {
            "services": IntVar(),
            "user_registered": IntVar(),
            "user_start": IntVar(),
            "user_installed": IntVar(),
            "system_registered": IntVar(),
            "system_start": IntVar(),
            "system_installed": IntVar(),
            "services_hash": "a4a40a3e5b91043137b5072da9a70832",
            "user_registered_hash": "d41d8cd98f00b204e9800998ecf8427e",
            "user_start_hash": "d41d8cd98f00b204e9800998ecf8427e",
            "user_installed_hash": "d41d8cd98f00b204e9800998ecf8427e",
            "system_registered_hash": "d41d8cd98f00b204e9800998ecf8427e",
            "system_start_hash": "d41d8cd98f00b204e9800998ecf8427e",
            "system_installed_hash": "d41d8cd98f00b204e9800998ecf8427e"
        }

        self.report = ""
        self.shim = ""
        self.shim_options = {
            "system": IntVar(),
            "ntuser": IntVar(),
            "software": IntVar(),
            "default": IntVar(),
            "sam": IntVar(),
            "security": IntVar()
        }
        self.system_report = IntVar()
        self.os_report = IntVar()
        self.app_report = IntVar()
        self.network_report = IntVar()
        self.device_report = IntVar()
        self.user_app_report = IntVar()

        # Hives
        self.default = ""
        self.ntuser = ""
        self.sam = ""
        self.security = ""
        self.software = ""
        self.system = ""

        self.control_set = ""
        self.sa_windir = ""
        self.sa_processor = ""
        self.sa_computer_name = ""
        self.sa_process_num = 0
        self.sa_path = ""
        self.sa_curr_version = ""
        self.sa_shutdown = ""
        self.sa_bios_vendor = "Processing ..."
        self.sa_bios_version = "Processing ..."
        self.sa_system_manufacturer = "Processing ..."
        self.sa_system_product_name = "Processing ..."
        self.os = {}

        # Function initializations
        self.toolbar()
        self.main_menu()
        self.load_settings()

        self.status = Label(self.master, text="STATUS: Ready", bd=1, relief=SUNKEN, anchor=W)
        self.progress = tkk.Progressbar(self.master, orient="horizontal", length=200, mode="indeterminate")
        self.progress.step(5)
        self.progress["value"] = 0

        self.status.pack(side=BOTTOM, fill=X, anchor=S, expand=True)
        self.progress.pack(side=RIGHT, fill=X, anchor=S, padx=2, in_=self.status)

        self.frame = Frame(self.master, width=800, height=500)

        self.root_tree = tkk.Treeview(self.master, height=400, columns=('Created', 'Modified'),
                                      selectmode='extended')

        self.session_frame = Frame(self.frame, width=300, height=500)
        self.session_frame.grid_rowconfigure(0, weight=1)
        self.session_frame.grid_columnconfigure(0, weight=1)
        self.session_frame.pack(side=LEFT, fill=BOTH, expand=True)

        xscrollbar = Scrollbar(self.session_frame, orient=HORIZONTAL)
        xscrollbar.grid(row=1, column=0, sticky=E + W)

        yscrollbar = Scrollbar(self.session_frame)
        yscrollbar.grid(row=0, column=1, sticky=N + S)

        self.canvas = Canvas(self.session_frame, bg='white', width=300, height=500)
        self.canvas_frame = Frame(self.canvas)
        self.canvas.config(xscrollcommand=xscrollbar.set, yscrollcommand=xscrollbar.set)
        yscrollbar.config(command=self.canvas.yview)
        xscrollbar.config(command=self.canvas.xview)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.canvas.create_window((0, 0), window=self.canvas_frame, anchor=N+W, width=550)

        self.canvas_frame.bind("<Configure>", lambda event, canvas=self.canvas: self.onFrameConfigure(self.canvas))

        r = 0
        Label(self.canvas_frame, text="Sessions", font="Arial 14 bold", fg="white", bg="orange").pack(fill=BOTH, expand=True)
        try:
            if not os.path.exists(os.getcwd()+"\\data\\sessions"):
                os.makedirs(os.getcwd() + "\\data\\sessions")

            for filename in os.listdir(os.getcwd()+"\\data\\sessions"):
                r += 1
                image = PhotoImage(file="data/img/regticksession.png", height=50, width=50)
                image.zoom(50, 50)
                tmp = filename

                session_info = self.get_config(filename)
                b = Button(self.canvas_frame, image=image, compound=LEFT, text=session_info, anchor=W, justify=LEFT,
                           command=lambda tmp=tmp: self.load_session(tmp))
                b.image = image
                b.pack(fill=BOTH, expand=True)
        except Exception:
            logging.error('[RegSmart] An error occurred in (Session loading)', exc_info=True, extra={'investigator': 'RegSmart'})
            self.display_message('error', 'An error occurred while Loading sessions.\nPlease try again.')

        tool_frame = Frame(self.frame, width=500, height=500)
        tool_frame.pack(side=RIGHT)

        Label(tool_frame, textvariable=self.session_name, font="Algerian 14 bold")\
            .grid(row=0, column=0, columnspan=6, sticky="nsew")

        image = PhotoImage(file="data/img/system.png", height=50, width=50)
        image.zoom(50, 50)
        b = Button(tool_frame, text="System Analysis", image=image, compound=TOP, command=self.system_analysis)
        b.image = image
        b.grid(row=2, column=2, sticky="nsew")

        image = PhotoImage(file="data/img/os.png", height=50, width=50)
        image.zoom(50, 50)
        b = Button(tool_frame, text="OS Analysis", image=image, compound=TOP, command=self.os_analysis)
        b.image = image
        b.grid(row=2, column=4, sticky="nsew")

        image = PhotoImage(file="data/img/application.png", height=50, width=50)
        image.zoom(50, 50)
        b = Button(tool_frame, text="Application Analysis", image=image, compound=TOP,
                   command=self.application_analysis)
        b.image = image
        b.grid(row=4, column=2, sticky="nsew")

        image = PhotoImage(file="data/img/network.png", height=50, width=50)
        image.zoom(50, 50)
        b = Button(tool_frame, text="Network Analysis", image=image, compound=TOP, command=self.network_analysis)
        b.image = image
        b.grid(row=4, column=4, sticky="nsew")

        image = PhotoImage(file="data/img/device.png", height=50, width=50)
        image.zoom(50, 50)
        b = Button(tool_frame, text="Device Analysis", image=image, compound=TOP, command=self.device_analysis)
        b.image = image
        b.grid(row=6, column=2, sticky="nsew")

        image = PhotoImage(file="data/img/cache.png", height=50, width=50)
        image.zoom(50, 50)
        b = Button(tool_frame, text="ShimCache Analysis", image=image, compound=TOP, command=self.shim_cache_gui)
        b.image = image
        b.grid(row=6, column=4, sticky="nsew")

        image = PhotoImage(file="data/img/regview.png", height=50, width=50)
        image.zoom(50, 50)
        b = Button(tool_frame, text="Registry Viewer", image=image, compound=TOP, command=self.regview)
        b.image = image
        b.grid(row=8, column=2, sticky="nsew")

        image = PhotoImage(file="data/img/report.png", height=50, width=50)
        image.zoom(50, 50)
        b = Button(tool_frame, text="Report", image=image, compound=TOP, command=self.make_report)
        b.image = image
        b.grid(row=8, column=4, sticky="nsew")

        image = PhotoImage(file="data/img/hash.png", height=50, width=50)
        image.zoom(50, 50)
        b = Button(tool_frame, text="Hash Generator", image=image, compound=TOP, command=self.hash_checker)
        b.image = image
        b.grid(row=10, column=2, sticky="nsew")

        image = PhotoImage(file="data/img/settings.png", height=50, width=50)
        image.zoom(50, 50)
        b = Button(tool_frame, text="Settings", image=image, compound=TOP, command=self.settings_gui)
        b.image = image
        b.grid(row=10, column=4, sticky="nsew")

        self.frame.pack(expand=True, fill=BOTH)

        self.master.update()

    def help(self):
        url = 'file://' + os.getcwd() + "\\data\\help\\index.html#userman"
        webbrowser.open_new(url)

    def load_settings(self):
        try:
            i = 0
            with open(os.getcwd() + "\\data\\config\\regsmart.conf", 'r') as file:
                for line in file:
                    if i < 7:
                        (key, val) = line.split()
                        key = key.strip(":")
                        self.db[str(key)].set(int(val))

                    elif i > 6 and i < 14:
                        (key, val) = line.split()
                        key = key.strip(":")
                        self.db[str(key)] = val
                    elif i == 14:
                        self.business_setting = line.split(":")[1]
                    elif i == 15:
                        self.location_setting = line.split(":")[1].split(",")
                    i += 1

        except Exception as ee:
            print(ee)
            logging.error('[RegSmart] An error occurred in (load_settings)', exc_info=True,
                          extra={'investigator': 'RegSmart'})
            return "Error occurred"

    def update_settings(self, display=None):
        self.rep_log("Saved settings")
        try:
            with open(os.getcwd() + "\\data\\config\\regsmart.conf", 'w') as file:
                final = ""
                b = self.business_setting.strip("\n")
                loc = ""
                for i in self.location_setting:
                    if i != "":
                        loc += i.strip("\n") + ","
                loc = loc[:-1]
                j = 0
                for i, k in self.db.items():
                    if j < 7:
                        final += i + ": " + str(k.get()) + "\n"
                    else:
                        final += i + ": " + k + "\n"
                    j += 1
                final += "business_name:" + b + "\n"
                final += "business_address:" + loc
                file.write(final)
            if not display:
                self.display_message("info", "Your settings have been updated successfully")
                self.settings.destroy()
        except Exception as ee:
            logging.error('[RegSmart] An error occurred in (Update settings)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

    def hash_folder(self):
        try:
            filename = fd.askdirectory()
            hash = Verification.get_hash(filename)
            self.display_message("info", "Filename: " + filename + "\n\nHash: " + hash
                                 + "\n\nNote: Contents copied to clipboard!")
            self.hash_gui.clipboard_clear()
            self.hash_gui.clipboard_append("Filename: " + filename + "\nHash: " + hash)
            self.hash_gui.focus_force()
        except Exception as ee:
            logging.error('[RegSmart] An error occurred in (hash folder)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

    def hash_file(self):
        try:
            filename = fd.askopenfilename()
            hash = Verification.hash_file(filename)
            self.display_message("info", "Filename: " + filename + "\n\nHash: " + hash
                                 + "\n\nNote: Contents copied to clipboard!")
            self.hash_gui.clipboard_clear()
            self.hash_gui.clipboard_append("Filename: " + filename + "\nHash: " + hash)
            self.hash_gui.focus_force()
        except Exception as ee:
            logging.error('[RegSmart] An error occurred in (hash file)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

    def hash_checker(self):
        self.rep_log("Opened Hash Generator")
        self.hash_gui = Toplevel()
        self.center_window(self.hash_gui, 353, 150)
        self.hash_gui.title("RegSmart: Hash Generator")
        self.hash_gui.iconbitmap("data/img/icon.ico")

        r = 1
        Label(self.hash_gui, font="Arial 16 bold", fg="black", bg="orange",
              text="Hash Generator") \
            .grid(row=0, column=0, columnspan=2, sticky="nsew")
        r += 1

        image = PhotoImage(file="data/img/folder.png", height=50, width=50)
        image.zoom(50, 50)
        b = Button(self.hash_gui, text="Hash Folder", image=image, compound=TOP, command=self.hash_folder)
        b.image = image
        b.grid(row=1, column=0, pady=10, sticky="nsew")

        image = PhotoImage(file="data/img/file.png", height=50, width=50)
        image.zoom(50, 50)
        b = Button(self.hash_gui, text="Hash File", image=image, compound=TOP, command=self.hash_file)
        b.image = image
        b.grid(row=1, column=1, pady=10, sticky="nsew")

        Label(self.hash_gui, font="Arial 10 bold", fg="black", bg="grey",
              text="Hash information is automatically copied to clipboard!") \
            .grid(row=2, column=0, columnspan=2, sticky="nsew")

    def human_bytes(self, B):
        B = float(B)
        KB = float(1024)
        MB = float(KB ** 2)  # 1,048,576
        GB = float(KB ** 3)  # 1,073,741,824
        TB = float(KB ** 4)  # 1,099,511,627,776

        if B < KB:
            return '{0} {1}'.format(B, 'Bytes' if 0 == B > 1 else 'Byte')
        elif KB <= B < MB:
            return '{0:.2f} KB'.format(B / KB)
        elif MB <= B < GB:
            return '{0:.2f} MB'.format(B / MB)
        elif GB <= B < TB:
            return '{0:.2f} GB'.format(B / GB)
        elif TB <= B:
            return '{0:.2f} TB'.format(B / TB)

    def get_config(self, filename):
        filename = "data\\sessions\\" + filename + "\\regsmart.session"
        try:
            with open(filename, 'r') as file:
                return file.read()
        except Exception:
            return "Error occurred"

    def get_size(self, start_path='.'):
        total_size = 0
        for dirpath, dirnames, filenames in os.walk(start_path):
            for f in filenames:
                fp = os.path.join(dirpath, f)
                total_size += os.path.getsize(fp)
        return total_size

    def donothing(self):
        filewin = Toplevel(self.master)
        button = Button(filewin, text="Do nothing button")
        button.pack()

    def onFrameConfigure(self, t=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def main_menu(self):
        menubar = Menu(self.master)
        self.master.config(menu=menubar)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="New Session", command=self.open_directory)
        # filemenu.add_command(label="Open ", command=self.open_directory)
        # filemenu.add_command(label="Save", command=self.donothing)
        # filemenu.add_command(label="Save as...", command=self.donothing)
        filemenu.add_command(label="Close Session", command=self.close_session)

        filemenu.add_separator()

        filemenu.add_command(label="Exit", command=self.confirm_quit)
        menubar.add_cascade(label="File", menu=filemenu)
        editmenu = Menu(menubar, tearoff=0)

        editmenu.add_separator()

        editmenu.add_command(label="System Analysis", command=self.system_analysis)
        editmenu.add_command(label="OS Analysis", command=self.os_analysis)
        editmenu.add_command(label="Application Analysis", command=self.application_analysis)
        editmenu.add_command(label="Network Analysis", command=self.network_analysis)
        editmenu.add_command(label="Device Analysis", command=self.device_analysis)
        editmenu.add_command(label="ShimCache Analysis", command=self.shim_cache_gui)
        editmenu.add_command(label="Registry Viewer", command=self.regview)
        editmenu.add_command(label="Hash Generator", command=self.hash_checker)
        editmenu.add_command(label="Report", command=self.make_report)

        menubar.add_cascade(label="Tools", menu=editmenu)
        helpmenu = Menu(menubar, tearoff=0)
        helpmenu.add_command(label="Help", command=self.help)
        helpmenu.add_command(label="About...", command=self.intro)
        menubar.add_cascade(label="Help", menu=helpmenu)

    def display_message(self, types, message):
        if types == "info":
            # thread.Thread(target=mb.showinfo, args=("RegSmart", message,)).start()
            mb.showinfo("RegSmart", message)
        if types == "error":
            # thread.Thread(target=mb.showerror, args=("RegSmart", message,)).start()
            mb.showerror("RegSmart", message)
        if types == "warning":
            # thread.Thread(target=mb.showwarning, args=("RegSmart", message,)).start()
            mb.showwarning("RegSmart", message)
        if types == "success":
            # mb._show(title="RegSmart", message=message, _icon="info", type="ok")
            tmp = Toplevel()
            tmp.title("RegSmart")
            tmp.iconbitmap("data/img/icon.ico")
            Frame(tmp, width=200, height=8).pack()
            photo = PhotoImage(file="data/img/regticksmall.png", height=100, width=100)
            label = Label(tmp, image=photo)
            label.image = photo
            label.pack(side=LEFT, padx=20, pady=10)
            Label(tmp, text="Successfully verified dumps!", font="Bold", pady=20).pack()
            Button(tmp, text="OK", command=tmp.destroy).pack(side=RIGHT, padx=100, pady=20, ipadx=20, ipady=2)
            self.center_window(tmp, 400, 150)
            tmp.focus_force()

    def get_answer(self, message):
        return mb.askquestion("RegSmart",message)

    def stop(self):
        self.stop_processing = True
        self.set_status("Stopped")
        self.progress.stop()
        logging.info("Stopped processing", extra={'investigator': self.investigator})

    def toolbar(self):
        toolbar = Frame(self.master, bg="gray")
        # Button(toolbar, text="Quit", command=self.stop, anchor=S).pack(side=LEFT)
        toolbar.pack(side=TOP, fill=X)

    def set_status(self, msg):
        self.status['text'] = "STATUS: " + msg

    def intro(self):
        self.rep_log("Opened about window")
        logging.info("Initiated about window", extra={'investigator': self.investigator})

        about = Toplevel(bg='white')
        about.title("RegSmart")
        about.iconbitmap("data/img/icon.ico")
        frame = Frame(about, width=400, height=100, bg='white')
        reg = PhotoImage(file="data/img/regsmart.png")
        label = Label(about, image=reg)
        label.image = reg
        label.pack()

        welcome = open("data/info/welcome", 'r').read()
        Label(about, text=welcome, bd=1, wraplength=500, bg='white').pack(fill=X)
        disclaimer = open("data/info/disclaimer", 'r').read()
        year = dt.datetime.now().year
        Label(about, text="Copyright@"+str(year), bd=1, relief=SUNKEN, anchor=E, wraplength=200)\
            .pack(side=BOTTOM, fill=X)
        Label(about, text=disclaimer, bd=1, bg="lightgrey", relief=SUNKEN, anchor=S).pack(side=BOTTOM, fill=X)
        frame.pack()
        self.center_window(about, 750, 400)
        about.focus_force()
        self.rep_log("Getting Investigator information")
        if self.investigator == "RegSmart":
            self.investigator = sd.askstring("RegSmart", "Investigators Name:")
            if not self.investigator or self.investigator == "" or self.investigator == " ":
                self.investigator = "RegSmart"
                self.rep_log("User failed to enter their name on the first try.")
                self.display_message("warning", "Please enter your name")
                self.investigator = sd.askstring("RegSmart", "Investigators Name:")
                if not self.investigator or self.investigator == "" or self.investigator == " ":
                    logging.info('Failed to get investigator name exiting ...', extra={'investigator': 'RegSmart'})
                    mb.showerror("RegSmart", "Failed to get Investigator Name.\nExiting ...")
                    exit(0)
            id = sd.askstring("RegSmart", "Investigators ID: ")
            if not id or id == "" or id == " ":
                self.rep_log("User failed to enter their ID on the first try.")
                self.display_message("warning", "Please enter your ID")
                id = sd.askstring("RegSmart", "Investigators ID:")
                if not id or id == "" or id == " ":
                    logging.info('Failed to get investigator id exiting ...', extra={'investigator': 'RegSmart'})
                    mb.showerror("RegSmart", "Failed to get Investigator ID.\nExiting ...")
                    exit(0)

            self.rep_log("Investigator entered name: " + self.investigator)
            self.rep_log("Investigator entered id: " + id)
            self.display_message("info", "Hello " + self.investigator + ".\nPlease wait while sessions are loading.")
            self.investigator += " (" + id + ")"

        else:
            about.focus_force()

        # self.master.update()
        self.master.deiconify()
        about.focus_force()

    def confirm_quit(self):
        tmp = self.get_answer("Do you want to quit?")
        if tmp == "yes":
            logging.info("Exiting RegSmart ...", extra={'investigator': self.investigator})
            self.master.destroy()
            exit(0)

    def update_loading(self):

        if self.directory != "":
            self.progress.start()
            self.set_status("Verifying dumps ...")
            self.display_message("info", "Verifying dumps for session. \nPress OK to continue.")
            tmp = Verification.verify_dump(self.directory)
            # tmp = 1
            if tmp == 2:
                self.display_message("warning", "No dumps were found in this directory please choose another.")
                self.set_status("No dumps found")
            elif tmp:
                self.display_message("info", "Successfully verified integrity of dumps")
                self.set_status("Processing ...")
                # name = self.directory.split("\\")
                # fold = name[len(name) - 1]
                # self.session = fold
                # self.session_name.set(fold.split("_")[0])
                thread.Thread(target=self.read_dumps).start()
            else:
                self.set_status("Dumps are not authentic")
                self.display_message("error", "The integrity of the dumps is not valid.")
                self.display_message("error", "RegSmart cannot process due to difference in original file"
                                              " and forensic copy.")
                self.display_message("info", "Please get new dumps and make sure that the dumps were"
                                             " successfully acquired.")
            self.progress['value'] = 0
            self.progress.stop()
            thread.Thread(target=self.reset_progress).start()

    def read_dumps(self):

        if self.directory != "":
            self.set_status("Processing dumps ...")
            self.progress.start()
            self.master.update()
            try:
                for filename in os.listdir(self.directory):
                    if filename == "DEFAULT":
                        self.default = Registry.Registry(self.directory + "/" + filename)
                        # thread.Thread(target=self.get_keys, args=(self.default.root(), self.u,)).start()

                    elif filename == "NTUSER.DAT":
                        self.ntuser = Registry.Registry(self.directory + "/" + filename)
                        # thread.Thread(target=self.get_keys, args=(self.ntuser.root(), self.cu,)).start()

                    elif filename == "SAM":
                        self.sam = Registry.Registry(self.directory + "/" + filename)
                        # thread.Thread(target=self.get_keys, args=(self.sam.root(), self.lm,)).start()

                    elif filename == "SECURITY":
                        self.security = Registry.Registry(self.directory + "/" + filename)
                        # thread.Thread(target=self.get_keys, args=(self.security.root(), self.lm,)).start()

                    elif filename == "SOFTWARE":
                        self.software = Registry.Registry(self.directory + "/" + filename)
                        # thread.Thread(target=self.get_keys, args=(self.software.root(), self.lm,)).start()

                    elif filename == "SYSTEM":
                        self.system = Registry.Registry(self.directory + "/" + filename)

                key = self.system.open("Select")
                for v in key.values():
                    if v.name() == "Current":
                        self.control_set = str(v.value())

                key = self.system.open("ControlSet00" + self.control_set + "\\Control\\Session Manager\\Environment")
                for v in key.values():
                    if v.name() == "PROCESSOR_ARCHITECTURE":
                        self.sa_processor = v.value()
            except Exception as ee:
                logging.error('An error occurred in (parsing registry)', exc_info=True,
                              extra={'investigator': 'RegSmart'})
                self.display_message("error", "Failed to Parse registry dumps.\nSession is now closing.")
                self.close_session()
        self.set_status("Ready")

    def utf8(self, data):
        # Returns a Unicode object on success, or None on failure

        try:
            return data.decode('utf-8')
        except UnicodeDecodeError:
            return None

    def get_keys(self, key, tree, depth=0):
        create = ""
        mod = ""
        if "DateCreated" in key.values():
            create = key.value("DateCreated")
        if "DateLastModified" in key.values():
            mod = key.value("DateLastModified")
        if "DateLastConnected" in key.values():
            mod = key.value("DateLastConnected")

        self.root_tree.insert(tree, 'end', text=key.path(), values=(create, mod))

        for subkey in key.subkeys():
            self.get_keys(subkey, tree, depth + 1)

    def reset_progress(self):
        time.sleep(2)
        self.progress['value'] = 0
        self.progress.stop()

    def close_session(self):
        self.rep_log("Session closed [" + self.full_session + "]")
        logging.info("Closed session ["+self.full_session+"]", extra={'investigator': self.investigator})
        self.stop_processing = False
        self.finished_parse = [False, False, False, False, False]
        self.current_user = None
        self.current_config = None
        self.local_machine = None
        self.users = None
        self.session = ""
        self.full_session = ""
        self.session_name.set("")

        # Hives
        self.default = ""
        self.ntuser = ""
        self.sam = ""
        self.security = ""
        self.software = ""
        self.system = ""
        self.session_name.set("")

        self.control_set = ""
        self.sa_windir = ""
        self.sa_processor = ""
        self.sa_computer_name = ""
        self.sa_process_num = 0
        self.sa_path = ""
        self.sa_curr_version = ""
        self.sa_shutdown = ""
        self.sa_bios_vendor = "Processing ..."
        self.sa_bios_version = "Processing ..."
        self.sa_system_manufacturer = "Processing ..."
        self.sa_system_product_name = "Processing ..."
        self.os = {}
        self.directory = ""
        self.master.title("RegSmart")

    def load_session(self, dir):
        self.close_session()
        sess = os.getcwd()+"\\data\\sessions\\"+dir
        if self.get_answer("Are you sure you want to load this session?\n"+dir) == "yes":
            self.rep_log("Loading session: " + dir)
            self.directory = sess
            tmp = self.directory.split("\\")
            self.full_session = tmp[len(tmp) - 1]
            if self.get_config(self.full_session) == "Error occurred":
                self.display_message("error", "Invalid session selected. Please re-import the session.")
                return
            self.session_name.set(self.full_session.split("_")[0])
            self.master.title("RegSmart: [" + tmp[len(tmp) - 1] + "]")
            self.update_loading()
            logging.info("Loaded session [" + self.full_session + "]", extra={'investigator': self.investigator})

    def reload_sessions(self):
        self.canvas.delete("all")
        self.canvas.pack_forget()
        self.session_frame.pack_forget()
        self.session_frame = Frame(self.frame, width=300, height=500)
        self.session_frame.grid_rowconfigure(0, weight=1)
        self.session_frame.grid_columnconfigure(0, weight=1)
        self.session_frame.pack(side=LEFT, fill=BOTH, expand=True)

        xscrollbar = Scrollbar(self.session_frame, orient=HORIZONTAL)
        xscrollbar.grid(row=1, column=0, sticky=E + W)

        yscrollbar = Scrollbar(self.session_frame)
        yscrollbar.grid(row=0, column=1, sticky=N + S)

        self.canvas = Canvas(self.session_frame, bg='white', width=300, height=500)
        self.canvas_frame = Frame(self.canvas)
        self.canvas.config(xscrollcommand=xscrollbar.set, yscrollcommand=xscrollbar.set)
        yscrollbar.config(command=self.canvas.yview)
        xscrollbar.config(command=self.canvas.xview)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.canvas.create_window((0, 0), window=self.canvas_frame, anchor=N + W, width=550)

        self.canvas_frame.bind("<Configure>", lambda event, canvas=self.canvas: self.onFrameConfigure(self.canvas))

        r = 0
        Label(self.canvas_frame, text="Sessions", font="Arial 14 bold", fg="white", bg="orange").pack(fill=BOTH,
                                                                                                      expand=True)
        try:
            for filename in os.listdir(os.getcwd() + "\\data\\sessions"):
                r += 1
                image = PhotoImage(file="data/img/regticksession.png", height=50, width=50)
                image.zoom(50, 50)
                tmp = filename

                session_info = self.get_config(filename)
                b = Button(self.canvas_frame, image=image, compound=LEFT, text=session_info, anchor=W, justify=LEFT,
                           command=lambda tmp=tmp: self.load_session(tmp))
                b.image = image
                b.pack(fill=BOTH, expand=True)
        except Exception:
            logging.error('[RegSmart] An error occurred in (Session loading)', exc_info=True,
                          extra={'investigator': 'RegSmart'})
            self.display_message('error', 'An error occurred while Loading sessions.\nPlease try again.')

    def copy_dumps(self, src, dst, symlinks=False, ignore=None):
        for item in os.listdir(src):
            s = os.path.join(src, item)
            d = os.path.join(dst, item)
            if os.path.isdir(s):
                shutil.copytree(s, d, symlinks, ignore)
            else:
                shutil.copy2(s, d)

    def create_config(self, case, dest, folder_name):
        folder_name = folder_name.split("_")
        file = open(dest+"\\regsmart.session", 'w')
        file.write("  Name: " + folder_name[0] + "\n")
        file.write("  Machine: " + folder_name[1] + "\n")
        file.write("  Date: " + folder_name[2] + "\n")
        file.write("  Size: " + self.human_bytes(self.get_size(dest)) + "\t\t\t\t")
        file.write("  Case: " + str(case) + "\n")
        file.close()

    def import_concurrent(self, dest, case, name):

        if not os.path.exists(dest):
            os.makedirs(dest)
            self.copy_dumps(self.directory, dest)
            self.create_config(case, dest, name)
            self.reload_sessions()
            self.update_loading()
        else:
            if self.get_answer("Session already exists.\nDo you want to replace it?") == "yes":
                try:
                    shutil.rmtree(dest)
                    os.makedirs(dest)
                    self.copy_dumps(self.directory, dest)
                    self.reload_sessions()
                    self.update_loading()
                except Exception:
                    self.display_message("error", "Failed to overwrite session.\nAccess is denied.\n"
                                                  "Please delete the folder from data/sessions/[session]")
            else:
                self.display_message("warning", "Please note that this new session is not going to be created.")

    def open_directory(self):
        self.rep_log("Creating new session")
        self.set_status("Ready")
        self.directory = fd.askdirectory()
        if self.directory != "":
            if Verification.is_valid_regacquire(self.directory):
                case = sd.askstring("RegSmart", "Enter Case Number:")
                if not case or case == "" or case == " ":
                    self.display_message("warning", "Please enter case number")
                    case = sd.askstring("RegSmart", "Enter Case Number:")
                    if not case or case == "" or case == " ":
                        self.display_message("error", "Failed to get case number, session will not be imported.")
                        return

                tmp = self.directory.split("/")
                self.master.title("RegSmart: ["+tmp[len(tmp)-1]+"]")
                self.session = tmp[len(tmp)-1].split("_")[0]
                self.full_session = tmp[len(tmp)-1]
                self.rep_log("New session [" + tmp[len(tmp)-1] + "]")
                self.rep_log("Case number [" + case + "]")
                self.session_name.set(self.session)
                self.set_status("Importing dumps...")
                self.progress.start()
                self.display_message("info", "Importing dumps...\nPlease wait\n\nYou will be notified when it's done.")

                dest = os.getcwd() + "\\data\\sessions\\"+tmp[len(tmp)-1]
                name = tmp[len(tmp) - 1]

                thread.Thread(target=self.import_concurrent, args=(dest, case, name,)).start()
            else:
                self.close_session()
                self.display_message("error", "This is not a valid RegAcquire folder!")
        else:
            self.close_session()
            self.display_message("warning", "Failed to create new session!")

    def center_window(self, tmp, width=300, height=200):
        # get screen width and height
        screen_width = tmp.winfo_screenwidth()
        screen_height = tmp.winfo_screenheight()

        # calculate position x and y coordinates
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        tmp.geometry('%dx%d+%d+%d' % (width, height, x, y))

    def system_analysis_data(self):
        services = []
        system = {}
        system["release_id"] = ""
        try:
            key = self.system.open("ControlSet00" + self.control_set + "\\Control\\Session Manager\\Environment")
            for v in key.values():
                if v.name() == "windir":
                    system["windir"] = v.value()
        except Exception:
            logging.error('An error occurred in (system_analysis - WinDir)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.system.open("ControlSet00" + self.control_set + "\\Control\\Session Manager\\Environment")
            for v in key.values():
                if v.name() == "PROCESSOR_ARCHITECTURE":
                    system["process_arch"] = v.value()
                if v.name() == "NUMBER_OF_PROCESSORS":
                    system["process_num"] = str(v.value())
                if v.name() == "Path":
                    system["path"] = v.value()
        except Exception:
            logging.error('An error occurred in (system_analysis - Processor)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.system.open("ControlSet00" + self.control_set + "\\Control\\ComputerName\\ComputerName")
            for v in key.values():
                if v.name() == "ComputerName":
                    system["computer_name"] = v.value()
        except Exception:
            logging.error('An error occurred in (system_analysis - Computer Name)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.system.open("ControlSet00" + self.control_set + "\\Control\\Windows")
            for v in key.values():
                if v.name() == "ShutdownTime":
                    system["shutdown"] = Parse.hex_windows_to_date(v.value().hex())
        except Exception:
            logging.error('An error occurred in (system_analysis - Shutdown time)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.software.open("Microsoft\\Windows NT\\CurrentVersion")
            for v in key.values():
                if v.name() == "ProductName":
                    system["product_name"] = v.value()
                if v.name() == "ReleaseId":
                    system["release_id"] = v.value()
        except Exception:
            logging.error('An error occurred in (system_analysis - Product info)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.system.open("ControlSet00" + self.control_set + "\\Services")
            for v in key.subkeys():
                services.append(v.name())
        except Exception:
            logging.error('An error occurred in (system_analysis - Services)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        return services, system

    def system_analysis(self):
        self.rep_log("Viewed system analysis")
        try:
            if self.directory != "" and self.system != "":
                logging.info("System Analysis on [" + self.full_session + "]", extra={'investigator': self.investigator})
                tk = Tk()
                tk.grid_columnconfigure(0, weight=1)
                tk.grid_columnconfigure(1, weight=1)
                tk.grid_columnconfigure(2, weight=1)
                tk.grid_columnconfigure(3, weight=1)
                tk.grid_rowconfigure(0, weight=1)
                tk.grid_rowconfigure(1, weight=1)
                self.center_window(tk, 500, 600)
                tk.title("RegSmart: System Analysis")
                tk.iconbitmap("data/img/icon.ico")

                services, system = self.system_analysis_data()

                r = 1
                Label(tk, font="Arial 14 bold", fg="blue", bg="yellow",
                      text="System Analysis \n["+self.full_session+"]")\
                    .grid(row=0, columnspan=4, sticky="nsew")
                r += 1

                Label(tk, text='Computer Name: ').grid(row=r, column=0)
                Label(tk, font="Helvetica 10 bold italic", text=system["computer_name"]).grid(row=r, column=1)
                r += 1

                Label(tk, text='OS: ').grid(row=r, column=0)
                Label(tk, font="Helvetica 10 bold italic", text=system["product_name"] + " " + system['release_id'])\
                    .grid(row=r, column=1)
                r += 1

                Label(tk, text='Last Shutdown: ').grid(row=r, column=0)
                Label(tk, font="Helvetica 10 bold italic", text=system["shutdown"]).grid(row=r, column=1)
                r += 1

                Label(tk, text='Processor Architecture: ').grid(row=r, column=0)
                Label(tk, font="Helvetica 10 bold italic", text=system["process_arch"]).grid(row=r, column=1)
                r += 1

                Label(tk, text='# of Processors: ').grid(row=r, column=0)
                Label(tk, font="Helvetica 10 bold italic", text=system["process_num"]).grid(row=r, column=1)
                r += 1

                # Label(tk, text="BIOS Vendor: ").grid(row=r, column=0)
                # v = Label(tk, font="Helvetica 10 bold italic", text=self.sa_bios_vendor)
                # v.grid(row=r, column=1)
                # r += 1
                #
                # Label(tk, text="BIOS Version: ").grid(row=r, column=0)
                # vv = Label(tk, font="Helvetica 10 bold italic", text=self.sa_bios_version)
                # vv.grid(row=r, column=1)
                # r += 1
                #
                # Label(tk, text="System Manufacturer: ").grid(row=r, column=0)
                # m = Label(tk, font="Helvetica 10 bold italic", text=self.sa_system_manufacturer)
                # m.grid(row=r, column=1)
                # r += 1
                #
                # Label(tk, text="System Product Name: ").grid(row=r, column=0)
                # mm = Label(tk, font="Helvetica 10 bold italic", text=self.sa_system_product_name)
                # mm.grid(row=r, column=1)
                # r += 1

                txt_frm = Frame(tk, width=350, height=200)
                txt_frm.grid(row=r, column=1, sticky="nsew")
                txt_frm.grid_propagate(False)
                txt_frm.grid_rowconfigure(0, weight=1)
                txt_frm.grid_columnconfigure(0, weight=1)
                txt = Text(txt_frm, borderwidth=3, relief="sunken")

                txt.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollb = Scrollbar(txt_frm, command=txt.yview)
                scrollb.grid(row=0, column=1, sticky='nsew')
                txt['yscrollcommand'] = scrollb.set
                Label(tk, text='Path: ').grid(row=r, column=0)
                for p in system["path"].split(";"):
                    txt.insert(END, p + "\n")
                txt.config(font=("consolas", 10), state=DISABLED)
                r += 2

                Label(tk, text='Services: ').grid(row=r, column=0)
                lb_frm = Frame(tk, width=350, height=210)
                lb_frm.grid(row=r, column=1, sticky="nsew")
                lb_frm.grid_propagate(False)
                lb_frm.grid_rowconfigure(0, weight=1)
                lb_frm.grid_columnconfigure(0, weight=1)
                lb = Listbox(lb_frm)
                for s in services:
                    lb.insert(END, s)
                lb.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollb = Scrollbar(lb_frm, command=lb.yview)
                scrollb.grid(row=0, column=1, sticky='nsew')
                lb['yscrollcommand'] = scrollb.set
                scrollbx = Scrollbar(lb_frm, command=lb.xview, orient=HORIZONTAL)
                scrollbx.grid(row=1, column=0, sticky='nsew')
                lb['xscrollcommand'] = scrollbx.set

                # if self.sa_system_product_name == "Processing ...":
                #     self.system_analysis_regsmart()

                # v['text'] = self.sa_bios_vendor
                # vv['text'] = self.sa_bios_version
                # m['text'] = self.sa_system_manufacturer
                # mm['text'] = self.sa_system_product_name
                tk.lift()

            else:
                self.rep_log("No session loaded")
                self.display_message('error', 'Please click on a session to load!')

        except Exception:
            logging.error('An error occurred in (system_analysis)', exc_info=True, extra={'investigator': 'RegSmart'})
            self.display_message('error', 'An error occurred while trying to extract information.\nPlease try again.')

    def os_analysis_data(self):
        try:
            self.os['RegisteredOrganization'] = "N/A"
            self.os['RegisteredOwner'] = "N/A"
            self.os['ReleaseId'] = ""
            key = self.software.open("Microsoft\\Windows NT\\CurrentVersion")
            for v in key.values():
                if v.name() == "ReleaseId":
                    self.os['ReleaseId'] = v.value()
                if v.name() == "ProductName":
                    self.os['ProductName'] = v.value()
                if v.name() == "ProductId":
                    self.os['ProductId'] = v.value()
                if v.name() == "PathName":
                    self.os['PathName'] = v.value()
                if v.name() == "InstallDate":
                    self.os['InstallDate'] = time.strftime('%a %b %d %H:%M:%S %Y (UTC)', time.gmtime(v.value()))
                if v.name() == "RegisteredOrganization":
                    self.os['RegisteredOrganization'] = v.value()
                    if self.os['RegisteredOrganization'] == "":
                        self.os['RegisteredOrganization'] = "N/A"
                if v.name() == "RegisteredOwner":
                    self.os['RegisteredOwner'] = v.value()
                if v.name() == "CurrentBuild":
                    self.os['CurrentBuild'] = v.value()
        except Exception:
            logging.error('An error occurred in (OS_analysis - Current Version)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.system.open("ControlSet00" + self.control_set + "\\Control\\Session Manager\\Environment")
            for v in key.values():
                if v.name() == "windir":
                    self.sa_windir = v.value()
        except Exception:
            logging.error('An error occurred in (OS_analysis - WinDir)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        sid_list = []
        users_paths_list = []
        mapping_list = []
        accounts = []

        try:
            key = self.software.open("Microsoft\\Windows NT\\CurrentVersion\\ProfileList")
            for v in key.subkeys():
                sid_list.append(v.name())
        except Exception:
            logging.error('An error occurred in (OS_analysis - SID)', exc_info=True, extra={'investigator': 'RegSmart'})

        try:
            for sid in sid_list:
                k = self.software.open("Microsoft\\Windows NT\\CurrentVersion\\ProfileList\\" + sid)
                for v in k.values():
                    if v.name() == "ProfileImagePath":
                        name = v.value().split("\\")
                        mapping_list.append(name[len(name) - 1])
                        users_paths_list.append(v.value())
        except Exception:
            logging.error('An error occurred in (OS_analysis - Profile)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.sam.open("SAM\\Domains\\Account\\Users\\Names")
            for v in key.subkeys():
                accounts.append(v.name())
        except Exception:
            logging.error('An error occurred in (OS_analysis - User accounts)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        return self.os, sid_list, users_paths_list, mapping_list, accounts

    def os_analysis(self):
        self.rep_log("Viewed os analysis")
        try:
            if self.directory != "" and self.software != "":
                logging.info("OS Analysis on [" + self.full_session + "]", extra={'investigator': self.investigator})
                tk = Tk()
                tk.grid_columnconfigure(0, weight=1)
                tk.grid_columnconfigure(1, weight=1)
                tk.grid_columnconfigure(2, weight=1)
                tk.grid_columnconfigure(3, weight=1)
                tk.grid_rowconfigure(0, weight=1)
                tk.grid_rowconfigure(1, weight=1)
                self.center_window(tk, 500, 520)
                tk.title("RegSmart: OS Analysis")
                tk.iconbitmap("data/img/icon.ico")

                self.os, sid_list, users_paths_list, mapping_list, accounts = self.os_analysis_data()

                r = 1
                Label(tk, font="Arial 14 bold", fg="white", bg="cyan", text="OS Analysis \n["+self.full_session+"]")\
                    .grid(row=0, columnspan=4, sticky="nsew")
                r += 1
                r += 1

                Label(tk, text='Product Name: ').grid(row=r, column=0)
                Label(tk, font="Helvetica 10 bold italic", text=self.os['ProductName']).grid(row=r, column=1)
                r += 1

                Label(tk, text='Release Id: ').grid(row=r, column=0)
                Label(tk, font="Helvetica 10 bold italic", text=self.os['ReleaseId']).grid(row=r, column=1)
                r += 1

                Label(tk, text='Current Build: ').grid(row=r, column=0)
                Label(tk, font="Helvetica 10 bold italic", text=self.os['CurrentBuild']).grid(row=r, column=1)
                r += 1

                Label(tk, text='Product Id: ').grid(row=r, column=0)
                Label(tk, font="Helvetica 10 bold italic", text=self.os['ProductId']).grid(row=r, column=1)
                r += 1

                Label(tk, text='Path Name: ').grid(row=r, column=0)
                Label(tk, font="Helvetica 10 bold italic", text=self.os['PathName']).grid(row=r, column=1)
                r += 1

                Label(tk, text="Install Date: ").grid(row=r, column=0)
                v = Label(tk, font="Helvetica 10 bold italic", text=self.os['InstallDate'])
                v.grid(row=r, column=1)
                r += 1

                Label(tk, text="Registered Organization: ").grid(row=r, column=0)
                vv = Label(tk, font="Helvetica 10 bold italic", text=self.os['RegisteredOrganization'])
                vv.grid(row=r, column=1)
                r += 1

                Label(tk, text="Registered Owner: ").grid(row=r, column=0)
                m = Label(tk, font="Helvetica 10 bold italic", text=self.os['RegisteredOwner'])
                m.grid(row=r, column=1)
                r += 1

                Label(tk, text="Windows Directory: ").grid(row=r, column=0)
                m = Label(tk, font="Helvetica 10 bold italic", text=self.sa_windir)
                m.grid(row=r, column=1)
                r += 1

                txt_frm = Frame(tk, width=350, height=150)
                txt_frm.grid(row=r, column=1, sticky="nsew")
                txt_frm.grid_propagate(False)
                txt_frm.grid_rowconfigure(0, weight=1)
                txt_frm.grid_columnconfigure(0, weight=1)
                tv = tkk.Treeview(txt_frm)
                tv['columns'] = ('MAP', 'PATH')
                tv.heading("#0", text='USERNAME')
                tv.column('#0', stretch=True)
                tv.heading('MAP', text='SID')
                tv.column('MAP', stretch=True)
                tv.heading('PATH', text='PATH')
                tv.column('PATH', stretch=True)

                tv.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollb = Scrollbar(txt_frm, command=tv.yview)
                scrollb.grid(row=0, column=1, sticky='nsew')
                tv['yscrollcommand'] = scrollb.set
                scrollbx = Scrollbar(txt_frm, command=tv.xview, orient=HORIZONTAL)
                scrollbx.grid(row=1, column=0, sticky='nsew')
                tv['xscrollcommand'] = scrollbx.set
                Label(tk, text='User Profiles: ').grid(row=r, column=0)

                for i in range(0, len(sid_list)):
                    tv.insert('', 'end', text=mapping_list[i], values=(sid_list[i], users_paths_list[i]))
                r += 1

                Label(tk, text='User Accounts: ').grid(row=r, column=0)
                lb_frm = Frame(tk, width=350, height=110)
                lb_frm.grid(row=r, column=1, sticky="nsew")
                lb_frm.grid_propagate(False)
                lb_frm.grid_rowconfigure(0, weight=1)
                lb_frm.grid_columnconfigure(0, weight=1)
                lb = Listbox(lb_frm)
                for a in accounts:
                    lb.insert(END, a)
                lb.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollb = Scrollbar(lb_frm, command=lb.yview)
                scrollb.grid(row=0, column=1, sticky='nsew')
                lb['yscrollcommand'] = scrollb.set
                scrollbx = Scrollbar(lb_frm, command=lb.xview, orient=HORIZONTAL)
                scrollbx.grid(row=1, column=0, sticky='nsew')
                lb['xscrollcommand'] = scrollbx.set
                tk.lift()

            else:
                self.rep_log("No session loaded")
                self.display_message('error', 'Please click on a session to load!')
        except Exception:
            logging.error('An error occurred in (OS_analysis)', exc_info=True, extra={'investigator': 'RegSmart'})
            self.display_message('error', 'An error occurred while processing\n Please try again.')

    def shim_cache_gui(self):
        self.rep_log("Opened the Shim Cache menu")
        try:
            if self.directory != "" and self.software != "":
                self.shim = Toplevel()
                self.center_window(self.shim, 355, 250)
                self.shim.title("RegSmart: Shim Cache")
                self.shim.iconbitmap("data/img/icon.ico")

                r = 1
                Label(self.shim, font="Arial 10 bold", fg="black", bg="grey",
                      text="Shim Cache") \
                    .grid(row=0, column=0, columnspan=2, sticky="nsew")
                r += 1

                Checkbutton(
                    self.shim, text="System Cache", variable=self.shim_options["system"], anchor=W, justify=LEFT
                ).grid(row=r, column=0, sticky=W)

                Checkbutton(
                    self.shim, text="Software Cache", variable=self.shim_options["software"], anchor=W, justify=LEFT
                ).grid(row=r, column=1, sticky=W)

                r += 1
                Checkbutton(
                    self.shim, text="Default Cache", variable=self.shim_options["default"], anchor=W, justify=LEFT
                ).grid(row=r, column=0, sticky=W)

                Checkbutton(
                    self.shim, text="Security Cache", variable=self.shim_options["security"], anchor=W, justify=LEFT
                ).grid(row=r, column=1, sticky=W)

                r += 1
                Checkbutton(
                    self.shim, text="Sam Cache", variable=self.shim_options["sam"], anchor=W, justify=LEFT
                ).grid(row=r, column=0, sticky=W)
                Checkbutton(
                    self.shim, text="NTuser Cache", variable=self.shim_options["ntuser"], anchor=W, justify=LEFT
                ).grid(row=r, column=1, sticky=W)
                r += 1
                # image = PhotoImage(file="data/img/report.png", height=50, width=50)
                # image.zoom(50, 50)
                sc = Button(self.shim, compound=LEFT, text="Generate results", anchor=W, justify=LEFT,
                             command=self.process_shim_cache)
                # rep.image = image
                sc.grid(row=r, columnspan=2, pady=20)
                r += 1
                Label(self.shim, font="Arial 8", fg="lightgrey", bg="black",
                      text="Cache that is not found will not be written!\n\n" +
                           "Shim Cache obtained from https://github.com/mandiant/ShimCacheParser", anchor=S) \
                    .grid(row=r, column=0, columnspan=2, sticky="nsew")
            else:
                self.rep_log("No session loaded")
                self.display_message('error', 'Please click on a session to load!')

        except Exception as e:
            print(e)
            logging.error('An error occurred in (Shim Cache)', exc_info=True, extra={'investigator': 'RegSmart'})
            self.display_message('error', 'An error occurred while processing\n Please try again.')

    def process_shim_cache(self):
        self.rep_log("Getting target destination for shim cache")
        self.progress.start()
        self.set_status("Generating Shim Cache ...")

        file = fd.askdirectory()

        if not file or file == "" or file == " ":
            file = fd.askdirectory()
        if not file or file == "" or file == " ":
            self.display_message("error", "Failed to get target directory. Shim Cache cannot be generated.")
            self.report.destroy()
            return

        path = os.getcwd() + "\\data\\sessions\\" + self.full_session + "\\"
        for filename in os.listdir(os.getcwd() + "\\data\\sessions\\" + self.full_session):
            if filename == "DEFAULT":
                if self.shim_options["default"].get():
                    self.save_shim_cache(path + filename, file+"\\default_cache.csv", "default_cache")

            if filename == "NTUSER.DAT":
                if self.shim_options["ntuser"].get():
                    self.save_shim_cache(path + filename, file+"\\user_cache.csv", "user_cache")

            if filename == "SAM":
                if self.shim_options["sam"].get():
                    self.save_shim_cache(path + filename, file + "\\sam_cache.csv", "sam_cache")

            if filename == "SECURITY":
                if self.shim_options["security"].get():
                    self.save_shim_cache(path + filename, file+"\\security_cache.csv", "security_cache")

            if filename == "SOFTWARE":
                if self.shim_options["software"].get():
                    self.save_shim_cache(path + filename, file+"\\software_cache.csv", "software_cache")

            if filename == "SYSTEM":
                if self.shim_options["system"].get():
                    self.save_shim_cache(path + filename, file+"\\system_cache.csv", "system_cache")

        self.display_message('info', 'Shim cache files have been saved.')
        self.shim.destroy()
        self.progress.stop()
        self.set_status("DONE")

    def regview(self):

        self.rep_log("Browsing Registry with Registry Viewer")
        try:
            if self.directory != "" and self.software != "":
                hives = ""
                path = os.getcwd() + "\\data\\sessions\\" + self.full_session + "\\"

                hives += " " + path + "DEFAULT"
                hives += " " + path + "NTUSER.DAT"
                hives += " " + path + "SAM"
                hives += " " + path + "SECURITY"
                hives += " " + path + "SOFTWARE"
                hives += " " + path + "SYSTEM"

                python3_command = os.getcwd() + "\\data\\lib\\regview\\RegView.exe " + hives
                # os.system(python3_command)
                process = subprocess.Popen(python3_command, stdin=None, stdout=None,
                                           close_fds=False, creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)

            else:
                self.display_message('error', 'Please click on a session to load!')

        except Exception as ee:
            print(ee)
            logging.error('An error occurred in (regview)', exc_info=True, extra={'investigator': 'RegSmart'})
            self.display_message('error', 'An error occurred while opening regview\n Please try again.')

    def save_shim_cache(self, hive, out, name):
        self.rep_log("Saving shim cache information [" + name + "]")
        try:
            python3_command = os.getcwd() + "\\data\\lib\\python\\py.exe " + os.getcwd() + \
                              "\\data\\lib\\ShimCacheParser\\ShimCacheParser.py -i " \
                              + hive + " -o " + out

            process = subprocess.Popen(python3_command, shell=True)
            # output, error = process.communicate()  # receive output from the python2 script
            # print(output)
            # print(error)
        except IOError as ee:
            print(ee)
            # self.display_message("error", "Error finding and saving Shim cache for " + hive)
            logging.error('An error occurred in (shim cache)', exc_info=True,  extra={'investigator': 'RegSmart'})

    def application_analysis_data(self):
        start_applications = []
        registered_applications = []
        installed_applications = []
        user_start_applications = []
        user_registered_applications = []
        user_installed_applications = []

        try:
            key = self.software.open("Microsoft\\Windows\\CurrentVersion\\Run")
            for v in key.values():
                start_applications.append(v.name())
        except Exception:
            logging.error('An error occurred in (application_analysis - Run)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.software.open("RegisteredApplications")
            for v in key.values():
                registered_applications.append(v.name())
        except Exception:
            logging.error('An error occurred in (application_analysis - Registered)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.software.open("Microsoft\\Windows\\CurrentVersion\\Uninstall")
            for v in key.subkeys():
                installed_applications.append(v.name())
        except Exception:
            logging.error('An error occurred in (application_analysis - Installed)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        # ===
        try:
            key = self.ntuser.open("Software\\Microsoft\\Windows\\CurrentVersion\\Run")
            for v in key.values():
                user_start_applications.append(v.name())
        except Exception:
            logging.error('An error occurred in (application_analysis - User Run)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.ntuser.open("Software\\RegisteredApplications")
            for v in key.values():
                user_registered_applications.append(v.name())
        except Exception:
            logging.error('An error occurred in (application_analysis - User Registerd)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.ntuser.open("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
            for v in key.subkeys():
                user_installed_applications.append(v.name())
        except Exception:
            logging.error('An error occurred in (application_analysis - User Installed)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            if "64" in self.sa_processor:
                key = self.software.open("WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
                for v in key.subkeys():
                    installed_applications.append(v.name())

                key = self.software.open("WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Run")
                for v in key.values():
                    start_applications.append(v.name())
        except Exception:
            logging.error('An error occurred in (application_analysis - 64 Run and Installed)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        return start_applications, registered_applications, installed_applications, user_start_applications, \
               user_registered_applications, user_installed_applications

    def application_analysis(self):
        self.rep_log("Viewed application analysis")
        try:
            if self.directory != "" and self.software != "":
                logging.info("Application Analysis on [" + self.full_session + "]", extra={'investigator': self.investigator})
                tk = Tk()
                tk.grid_columnconfigure(0, weight=1)
                tk.grid_columnconfigure(1, weight=1)
                tk.grid_columnconfigure(2, weight=1)
                tk.grid_columnconfigure(3, weight=1)
                tk.grid_rowconfigure(0, weight=1)
                tk.grid_rowconfigure(1, weight=1)
                self.center_window(tk, 850, 520)
                tk.title("RegSmart: Application Analysis")
                tk.iconbitmap("data/img/icon.ico")

                start_applications, registered_applications, installed_applications, user_start_applications, \
                user_registered_applications, user_installed_applications = self.application_analysis_data()

                r = 1
                Label(tk, font="Arial 14 bold", fg="white", bg="brown", text="Application Analysis \n["+self.full_session+"]")\
                    .grid(row=0, columnspan=4, sticky="nsew")
                r += 1
                r += 1
                Label(tk, font="Arial 12 bold", text="System").grid(row=r, column=1)
                Label(tk, font="Arial 12 bold", text="User [" + self.session_name.get() + "]").grid(row=r, column=3)
                r += 2
                Label(tk, text='StartUp Programs: ').grid(row=r, column=0)
                lb_frm = Frame(tk, width=350, height=110)
                lb_frm.grid(row=r, column=1, sticky="nsew")
                lb_frm.grid_propagate(False)
                lb_frm.grid_rowconfigure(0, weight=1)
                lb_frm.grid_columnconfigure(0, weight=1)
                lb = Listbox(lb_frm)
                for a in start_applications:
                    lb.insert(END, a)
                lb.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollb = Scrollbar(lb_frm, command=lb.yview)
                scrollb.grid(row=0, column=1, sticky='nsew')
                lb['yscrollcommand'] = scrollb.set
                scrollbx = Scrollbar(lb_frm, command=lb.xview, orient=HORIZONTAL)
                scrollbx.grid(row=1, column=0, sticky='nsew')
                lb['xscrollcommand'] = scrollbx.set

                lb_frm = Frame(tk, width=350, height=110,)
                lb_frm.grid(row=r, column=3, sticky="nsew")
                lb_frm.grid_propagate(False)
                lb_frm.grid_rowconfigure(0, weight=1)
                lb_frm.grid_columnconfigure(0, weight=1)
                lb = Listbox(lb_frm)
                for a in user_start_applications:
                    lb.insert(END, a)
                lb.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollb = Scrollbar(lb_frm, command=lb.yview)
                scrollb.grid(row=0, column=1, sticky='nsew')
                lb['yscrollcommand'] = scrollb.set
                scrollbx = Scrollbar(lb_frm, command=lb.xview, orient=HORIZONTAL)
                scrollbx.grid(row=1, column=0, sticky='nsew')
                lb['xscrollcommand'] = scrollbx.set
                r += 2

                Label(tk, text='Registered Programs: ').grid(row=r, column=0)
                lbr_frm = Frame(tk, width=350, height=110)
                lbr_frm.grid(row=r, column=1, sticky="nsew")
                lbr_frm.grid_propagate(False)
                lbr_frm.grid_rowconfigure(0, weight=1)
                lbr_frm.grid_columnconfigure(0, weight=1)
                lbr = Listbox(lbr_frm)
                for a in registered_applications:
                    lbr.insert(END, a)
                lbr.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollbr = Scrollbar(lbr_frm, command=lbr.yview)
                scrollbr.grid(row=0, column=1, sticky='nsew')
                lbr['yscrollcommand'] = scrollbr.set
                scrollbxr = Scrollbar(lbr_frm, command=lbr.xview, orient=HORIZONTAL)
                scrollbxr.grid(row=1, column=0, sticky='nsew')
                lbr['xscrollcommand'] = scrollbxr.set

                lbr_frm = Frame(tk, width=350, height=110)
                lbr_frm.grid(row=r, column=3, sticky="nsew" )
                lbr_frm.grid_propagate(False)
                lbr_frm.grid_rowconfigure(0, weight=1)
                lbr_frm.grid_columnconfigure(0, weight=1)
                lbr = Listbox(lbr_frm)
                for a in user_registered_applications:
                    lbr.insert(END, a)
                lbr.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollbr = Scrollbar(lbr_frm, command=lbr.yview)
                scrollbr.grid(row=0, column=1, sticky='nsew')
                lbr['yscrollcommand'] = scrollbr.set
                scrollbxr = Scrollbar(lbr_frm, command=lbr.xview, orient=HORIZONTAL)
                scrollbxr.grid(row=1, column=0, sticky='nsew')
                lbr['xscrollcommand'] = scrollbxr.set
                r += 2

                Label(tk, text='Installed Programs: ').grid(row=r, column=0)
                lbr_frm = Frame(tk, width=350, height=220)
                lbr_frm.grid(row=r, column=1, sticky="nsew")
                lbr_frm.grid_propagate(False)
                lbr_frm.grid_rowconfigure(0, weight=1)
                lbr_frm.grid_columnconfigure(0, weight=1)
                lbr = Listbox(lbr_frm)
                for i in installed_applications:
                    lbr.insert(END, i)
                lbr.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollbr = Scrollbar(lbr_frm, command=lbr.yview)
                scrollbr.grid(row=0, column=1, sticky='nsew')
                lbr['yscrollcommand'] = scrollbr.set
                scrollbxr = Scrollbar(lbr_frm, command=lbr.xview, orient=HORIZONTAL)
                scrollbxr.grid(row=1, column=0, sticky='nsew')
                lbr['xscrollcommand'] = scrollbxr.set

                lbr_frm = Frame(tk, width=350, height=220)
                lbr_frm.grid(row=r, column=3, sticky="nsew")
                lbr_frm.grid_propagate(False)
                lbr_frm.grid_rowconfigure(0, weight=1)
                lbr_frm.grid_columnconfigure(0, weight=1)
                lbr = Listbox(lbr_frm)
                for i in user_installed_applications:
                    lbr.insert(END, i)
                lbr.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollbr = Scrollbar(lbr_frm, command=lbr.yview)
                scrollbr.grid(row=0, column=1, sticky='nsew')
                lbr['yscrollcommand'] = scrollbr.set
                scrollbxr = Scrollbar(lbr_frm, command=lbr.xview, orient=HORIZONTAL)
                scrollbxr.grid(row=1, column=0, sticky='nsew')
                lbr['xscrollcommand'] = scrollbxr.set

                tk.lift()

            else:
                self.rep_log("No session loaded")
                self.display_message('error', 'Please click on a session to load!')
        except Exception:
            logging.error('An error occurred in (OS_analysis)', exc_info=True, extra={'investigator': 'RegSmart'})
            self.display_message('error', 'An error occurred while processing\n Please try again.')

    def network_analysis_data(self):
        cards = []
        intranet = []
        wireless = []
        matched = []

        try:
            key = self.software.open("Microsoft\\Windows NT\\CurrentVersion\\NetworkCards")
            for v in key.subkeys():
                for n in v.values():
                    if n.name() == "Description":
                        cards.append(n.value())
        except Exception:
            logging.error('An error occurred in (network_analysis - Cards)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.software.open("Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Nla\\Cache\\Intranet")
            for v in key.subkeys():
                intranet.append(v.name())
        except Exception:
            logging.error('An error occurred in (network_analysis - Intranets)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.software.open("Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Nla\\Wireless")
            for v in key.subkeys():
                wireless.append(v.name())
        except Exception:
            logging.error('An error occurred in (network_analysis - Wireless)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.software.open("Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Profiles")
            for v in key.subkeys():
                tmp = {}
                tmp["ID"] = v.name()
                for s in v.values():
                    if s.name() == "Description":
                        tmp["Description"] = s.value()
                    if s.name() == "DateCreated":
                        tmp["Created"] = Parse.parse_date(s.value().hex())
                    if s.name() == "DateLastConnected":
                        tmp["Modified"] = Parse.parse_date(s.value().hex())
                matched.append(tmp)
        except Exception:
            logging.error('An error occurred in (network_analysis - Extracting Wireless profiles)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        return cards, intranet, wireless, matched

    def network_analysis(self):
        self.rep_log("Viewed network analysis")
        try:
            if self.directory != "" and self.software != "":
                logging.info("Network Analysis on [" + self.full_session + "]", extra={'investigator': self.investigator})
                tk = Tk()
                tk.grid_columnconfigure(0, weight=1)
                tk.grid_columnconfigure(1, weight=1)
                tk.grid_columnconfigure(2, weight=1)
                tk.grid_columnconfigure(3, weight=1)
                tk.grid_rowconfigure(0, weight=1)
                tk.grid_rowconfigure(1, weight=1)
                self.center_window(tk, 800, 520)
                tk.title("RegSmart: Application Analysis")
                tk.iconbitmap("data/img/icon.ico")

                cards, intranet, wireless, matched = self.network_analysis_data()

                r = 1
                Label(tk, font="Arial 14 bold", fg="white", bg="green", text="Network Analysis \n["+self.full_session+"]") \
                    .grid(row=0, columnspan=4, sticky="nsew")
                r += 1

                Label(tk, text='Network Cards: ').grid(row=r, column=0)
                lb_frm = Frame(tk, width=350, height=110)
                lb_frm.grid(row=r, column=1, sticky="nsew")
                lb_frm.grid_propagate(False)
                lb_frm.grid_rowconfigure(0, weight=1)
                lb_frm.grid_columnconfigure(0, weight=1)
                lb = Listbox(lb_frm)
                for a in cards:
                    lb.insert(END, a)
                lb.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollb = Scrollbar(lb_frm, command=lb.yview)
                scrollb.grid(row=0, column=1, sticky='nsew')
                lb['yscrollcommand'] = scrollb.set
                scrollbx = Scrollbar(lb_frm, command=lb.xview, orient=HORIZONTAL)
                scrollbx.grid(row=1, column=0, sticky='nsew')
                lb['xscrollcommand'] = scrollbx.set
                r += 1

                Label(tk, text='Intranet Networks: ').grid(row=r, column=0)
                lb_frm = Frame(tk, width=350, height=110)
                lb_frm.grid(row=r, column=1, sticky="nsew")
                lb_frm.grid_propagate(False)
                lb_frm.grid_rowconfigure(0, weight=1)
                lb_frm.grid_columnconfigure(0, weight=1)
                lb = Listbox(lb_frm)
                for a in intranet:
                    lb.insert(END, a)
                lb.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollb = Scrollbar(lb_frm, command=lb.yview)
                scrollb.grid(row=0, column=1, sticky='nsew')
                lb['yscrollcommand'] = scrollb.set
                scrollbx = Scrollbar(lb_frm, command=lb.xview, orient=HORIZONTAL)
                scrollbx.grid(row=1, column=0, sticky='nsew')
                lb['xscrollcommand'] = scrollbx.set
                r += 1

                txt_frm = Frame(tk, width=650, height=250)
                txt_frm.grid(row=r, column=1, sticky="nsew")
                txt_frm.grid_propagate(False)
                txt_frm.grid_rowconfigure(0, weight=1)
                txt_frm.grid_columnconfigure(0, weight=1)
                tv = tkk.Treeview(txt_frm)
                tv['columns'] = ('Created', 'LastConnected', 'ID')
                tv.heading("#0", text='Description')
                tv.column('#0', stretch=True)
                tv.heading('Created', text='Created')
                tv.column('Created', stretch=True)
                tv.heading('LastConnected', text='LastConnected')
                tv.column('LastConnected', stretch=True)
                tv.heading('ID', text='ID')
                tv.column('ID', stretch=True)

                tv.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
                scrollb = Scrollbar(txt_frm, command=tv.yview)
                scrollb.grid(row=0, column=1, sticky='nsew')
                tv['yscrollcommand'] = scrollb.set
                scrollbx = Scrollbar(txt_frm, command=tv.xview, orient=HORIZONTAL)
                scrollbx.grid(row=1, column=0, sticky='nsew')
                tv['xscrollcommand'] = scrollbx.set
                Label(tk, text='Wireless: ').grid(row=r, column=0)

                for u in matched:
                    tv.insert('', 'end', text=u["Description"], values=(u["Created"], u["Modified"], u["ID"]))
                r += 1

                tk.lift()

            else:
                self.rep_log("No session loaded")
                self.display_message('error', 'Please click on a session to load!')
        except Exception:
            logging.error('An error occurred in (OS_analysis)', exc_info=True, extra={'investigator': 'RegSmart'})
            self.display_message('error', 'An error occurred while processing\n Please try again.')

    def device_analysis_data(self):
        logging.info("Device Analysis on [" + self.full_session + "]", extra={'investigator': self.investigator})
        printer = []
        usb = []

        try:
            key = self.system.open("ControlSet00" + self.control_set + "\\Control\\Print\\Environments")
            for v in key.subkeys():
                for s in v.subkeys():
                    if s.name() == "Drivers":
                        for ss in s.subkeys():
                            if "Version" in ss.name():
                                for d in ss.subkeys():
                                    printer.append(d.name())
        except Exception:
            logging.error('An error occurred in (device_analysis - Print Drivers)', exc_info=True,
                          extra={'investigator': 'RegSmart'})

        try:
            key = self.system.open("ControlSet00" + self.control_set + "\\Enum\\USBSTOR")
            for v in key.subkeys():
                name = v.name().split("&")
                tmp = {}
                serial = []
                tmp["Type"] = name[0]
                tmp["Vendor"] = name[1].split("_")[1]
                tmp["Product"] = name[2].split("_")[1]
                tmp["Revision"] = name[3].split("_")[1]

                for s in v.subkeys():
                    serial_data = {}
                    serial_data["ID"] = s.name()
                    serial_data["InstallDate"] = "N/A"
                    serial_data["LastArrivalDate"] = "N/A"
                    serial_data["LastRemovalDate"] = "N/A"
                    for ss in s.subkeys():
                        if ss.name() == "Properties":
                            for sss in ss.subkeys():
                                for ssss in sss.subkeys():
                                    if ssss.name() == "0064":
                                        for ssssv in ssss.values():
                                            serial_data["InstallDate"] = ssssv.value().replace(microsecond=0)
                                    if ssss.name() == "0066":
                                        for ssssv in ssss.values():
                                            serial_data["LastArrivalDate"] = ssssv.value().replace(microsecond=0)
                                    if ssss.name() == "0067":
                                        for ssssv in ssss.values():
                                            serial_data["LastRemovalDate"] = ssssv.value().replace(microsecond=0)
                            serial.append(serial_data)
                tmp["Serial"] = serial
                usb.append(tmp)
        except Exception:
            logging.error('An error occurred in (device_analysis - USB)', exc_info=True,
                          extra={'investigator': 'RegSmart'})
        return printer, usb

    def device_analysis(self):
        self.rep_log("Viewed device analysis")
        try:
            if self.directory != "" and self.system != "":
                tk = Tk()
                tk.grid_columnconfigure(0, weight=1)
                tk.grid_columnconfigure(1, weight=2)
                tk.grid_columnconfigure(2, weight=2)
                tk.grid_columnconfigure(3, weight=2)
                tk.grid_rowconfigure(0, weight=1)
                tk.grid_rowconfigure(1, weight=1)
                self.center_window(tk, 950, 520)
                tk.title("RegSmart: Application Analysis")
                tk.iconbitmap("data/img/icon.ico")

                printer, usb = self.device_analysis_data()

                r = 1
                Label(tk, font="Arial 14 bold", fg="white", bg="Black", text="Device Analysis \n["+self.full_session+"]") \
                    .grid(row=0, columnspan=4, sticky="nsew")
                r += 1

                Label(tk, text='Printers: ').grid(row=r, column=0)
                lb_frm = Frame(tk, width=850, height=210)
                lb_frm.grid(row=r, column=1, sticky="nsew")
                lb_frm.grid_propagate(False)
                lb_frm.grid_rowconfigure(0, weight=1)
                lb_frm.grid_columnconfigure(0, weight=1)
                lb = Listbox(lb_frm)
                for a in printer:
                    lb.insert(END, a)
                lb.grid(row=0, column=0, sticky="nsew")
                scrollb = Scrollbar(lb_frm, command=lb.yview)
                scrollb.grid(row=0, column=1, sticky='nsew')
                lb['yscrollcommand'] = scrollb.set
                scrollbx = Scrollbar(lb_frm, command=lb.xview, orient=HORIZONTAL)
                scrollbx.grid(row=1, column=0, sticky='nsew')
                lb['xscrollcommand'] = scrollbx.set
                r += 1

                txt_frm = Frame(tk, width=850, height=250)
                txt_frm.grid(row=r, column=1, sticky="nsew")
                txt_frm.grid_propagate(False)
                txt_frm.grid_rowconfigure(0, weight=1)
                txt_frm.grid_columnconfigure(0, weight=1)
                tv = tkk.Treeview(txt_frm)
                tv['columns'] = ('Vendor', 'Product', 'Revision')
                tv.heading("#0", text='Type')
                tv.column('#0', stretch=True)
                tv.heading('Vendor', text='#Vendor')
                tv.column('Vendor', stretch=True)
                tv.heading('Product', text='#Product')
                tv.column('Product', stretch=True)
                tv.heading('Revision', text='#Revision')
                tv.column('Revision', stretch=True)

                tv.grid(row=0, column=0, sticky="nsew")
                scrollb = Scrollbar(txt_frm, command=tv.yview)
                scrollb.grid(row=0, column=1, sticky='nsew')
                tv['yscrollcommand'] = scrollb.set
                scrollbx = Scrollbar(txt_frm, command=tv.xview, orient=HORIZONTAL)
                scrollbx.grid(row=1, column=0, sticky='nsew')
                tv['xscrollcommand'] = scrollbx.set
                Label(tk, text='USB: ').grid(row=r, column=0)
                tv.tag_configure('new_serials', background='lightgrey')

                for u in usb:
                    tmp = tv.insert('', 'end', text=u["Type"], values=(u["Vendor"], u["Product"], u["Revision"]))
                    ser = tv.insert(tmp, 'end', text='#Serials', values=('#Installed Date', '#Plugged Date',
                                                                        '#Unplugged Date'), tags = ('new_serials',))
                    for z in u["Serial"]:
                        tv.insert(ser, 'end', text=z['ID'], values=(z['InstallDate'], z['LastArrivalDate'],
                                                                    z['LastRemovalDate']))
                r += 1
                tk.lift()

            else:
                self.rep_log("No session loaded")
                self.display_message('error', 'Please click on a session to load!')
        except Exception:
            logging.error('An error occurred in (OS_analysis)', exc_info=True, extra={'investigator': 'RegSmart'})
            self.display_message('error', 'An error occurred while processing\n Please try again.')

    def key_to_string(self, key):
        tmp = ""
        for value in [v for v in key.values() \
                      if v.value_type() == Registry.RegSZ or \
                                      v.value_type() == Registry.RegExpandSZ]:
            tmp += value.name() + ": " + value.value()
        return tmp

    def find_key(self, reg, k):
        try:
            key = reg.open(k)
            return key

        except Registry.RegistryKeyNotFoundException:
            return None

    def make_report(self):

        try:
            if self.directory != "" and self.system != "":
                self.rep_log("Opened the report menu")
                self.has_report = "True"
                self.report = Toplevel()
                self.center_window(self.report, 400, 450)
                self.report.title("RegSmart: Report")
                self.report.iconbitmap("data/img/icon.ico")

                r = 1
                Label(self.report, font="Arial 10 bold", fg="blue", bg="yellow",
                      text="Reporting Information") \
                    .grid(row=0, column=0, columnspan=2, sticky="nsew")
                r += 1

                Label(self.report, text="Enter Organization name: ") \
                    .grid(row=r, column=0, sticky="nsew")
                r += 1
                tmp_business = self.business_setting
                self.business = Entry(self.report)
                self.business.insert(0, tmp_business)
                self.business.grid(row=r, column=0, columnspan=2)
                r += 1
                Label(self.report, text="Enter Organization Location (Line separated): ")\
                    .grid(row=r, column=0, sticky="nsew")
                r += 1
                tmp_loc = self.location_setting
                self.location = Text(self.report, height=10, width=10)
                for i in tmp_loc:
                    self.location.insert(END, i + '\n')
                self.location.grid(row=r, column=0, sticky="nsew", columnspan=2)
                r += 1

                self.system_report.set(1)
                self.os_report.set(1)
                self.user_app_report.set(1)
                self.app_report.set(1)
                self.network_report.set(1)
                self.device_report.set(1)

                Checkbutton(
                    self.report, text="System Analysis", variable=self.system_report, anchor=W, justify=LEFT
                ).grid(row=r, column=0, sticky=W)

                Checkbutton(
                    self.report, text="OS Analysis", variable=self.os_report, anchor=W, justify=LEFT
                ).grid(row=r, column=1, sticky=W)
                r += 1
                Checkbutton(
                    self.report, text="User Application Analysis", variable=self.user_app_report, anchor=W, justify=LEFT
                ).grid(row=r, column=0, sticky=W)
                Checkbutton(
                    self.report, text="System Application Analysis", variable=self.app_report, anchor=W, justify=LEFT
                ).grid(row=r, column=1, sticky=W)
                r += 1
                Checkbutton(
                    self.report, text="Network Analysis", variable=self.network_report, anchor=W, justify=LEFT
                ).grid(row=r, column=0, sticky=W)

                Checkbutton(
                    self.report, text="Device Analysis", variable=self.device_report, anchor=W, justify=LEFT
                ).grid(row=r, column=1, sticky=W)
                r += 6

                rep = Button(self.report, compound=LEFT, text="Generate Report", anchor=W, justify=LEFT,
                           command=self.generate_report)

                rep.grid(row=r, columnspan=2, pady=20 )

            else:
                self.rep_log("No session loaded, not showing report menu")
                self.display_message('error', 'Please select a valid session')
        except Exception as ee:
            print(ee)
            self.display_message("error", "Failed to create report.\nThis could be due to several reasons, "
                                          "PDF file is already opened or incorrect options choosed.")
            logging.error('Report failed to be created', extra={'investigator': 'RegSmart'})

    def generate_report(self):
        try:
            self.set_status("Verifying dumps ...")
            self.progress.start()
            self.rep_log("Generating report")
            self.display_message("info", "Creating report\nPlease wait while dumps are being verified ...\nPress OK to continue.")

            tmp = Verification.verify_dump(self.directory)
            # tmp = 1
            case = self.get_config(self.full_session)
            self.progress.stop()
            self.set_status("Done")
            if case == "Error occurred":
                self.display_message("error", "Session is not valid. Report cannot be created.")
                return

            case = case.split(":")[5]

            if tmp:
                b = [self.business.get()]
                tmp = self.location.get(1.0, END).split('\n')
                self.business_setting = self.business.get().strip("\n")
                self.location_setting = tmp
                for i in tmp:
                    b.append(i)
                self.rep_log("Updating the business information.")
                self.update_settings(True)
                data = {}
                if self.system_report.get():
                    data["system"] = self.system_analysis_data()

                if self.os_report.get():
                    data["os"] = self.os_analysis_data()

                if self.network_report.get():
                    data["network"] = self.network_analysis_data()

                if self.user_app_report.get():
                    data["user_application"] = self.application_analysis_data()

                if self.app_report.get():
                    data["application"] = self.application_analysis_data()

                if self.device_report.get():
                    data["device"] = self.device_analysis_data()

                file = fd.asksaveasfilename(defaultextension=".pdf", initialfile=self.full_session,
                                        filetypes=[("PDF Document", "*.pdf")])

                if not file or file == "" or file == " ":
                    self.rep_log("Failed to set the saving destination.")
                    file = fd.asksaveasfilename(defaultextension=".pdf", initialfile=self.full_session,
                                            filetypes=[("PDF Document", "*.pdf")])
                if not file or file == "" or file == " ":
                    self.rep_log("Failed to set the saving destination for the second time. Not making the report.")
                    self.display_message("error", "Failed to get report target file. Report cannot be created.")
                    self.report.destroy()
                    return

                self.progress.start()
                self.set_status("Generating report ...")
                self.display_message("info", "Generating Report\nPlease click OK to continue.")
                Reports.standard_report(file, self.full_session + "@" + case, self.investigator, b, self.db,
                                        self.report_log, self.timeline_log, data)
                self.display_message('info', 'Your report has been saved.')

                self.report.destroy()
                self.progress.stop()
                self.set_status("DONE")

            else:
                self.display_message('error', 'Integrity of the session violated, Report cannot be created.')

        except Exception as ee:
            print(ee)
            self.display_message("error", "Failed to create report.\nThis could be due to several reasons, "
                                          "PDF file is already opened or incorrect options choosed.")
            logging.error('Report failed to be created', exc_info=True, extra={'investigator': 'RegSmart'})

    def settings_gui(self):
        self.rep_log("Opened settings menu")
        try:
            self.settings = Toplevel()
            self.center_window(self.settings, 410, 200)
            self.settings.title("RegSmart: Settings")
            self.settings.iconbitmap("data/img/icon.ico")

            r = 1
            Label(self.settings, font="Arial 10 bold", fg="black", bg="lightgreen",
                  text="Settings") \
                .grid(row=0, column=0, columnspan=2, sticky="nsew")
            r += 1
            Label(self.settings, font="Arial 8 bold", fg="black", bg="lightblue",
                  text="Exclusion Databases:") \
                .grid(row=r, column=0, columnspan=2, sticky="nsew")
            r += 1

            Checkbutton(
                self.settings, text="services.db", variable=self.db["services"], anchor=W, justify=LEFT
            ).grid(row=r, column=0, sticky=W)

            Checkbutton(
                self.settings, text="system_installed_applications.db",
                variable=self.db["system_installed"], anchor=W, justify=LEFT
            ).grid(row=r, column=1, sticky=W)

            r += 1
            Checkbutton(
                self.settings, text="system_registered_applications.db",
                variable=self.db["system_registered"], anchor=W, justify=LEFT
            ).grid(row=r, column=0, sticky=W)

            Checkbutton(
                self.settings, text="system_start_applications.db",
                variable=self.db["system_start"], anchor=W, justify=LEFT
            ).grid(row=r, column=1, sticky=W)

            r += 1
            Checkbutton(
                self.settings, text="user_installed_applications.db",
                variable=self.db["user_installed"], anchor=W, justify=LEFT
            ).grid(row=r, column=0, sticky=W)

            Checkbutton(
                self.settings, text="user_registered_applications.db",
                variable=self.db["user_registered"], anchor=W, justify=LEFT
            ).grid(row=r, column=1, sticky=W)
            r += 1
            Checkbutton(
                self.settings, text="user_start_applications.db",
                variable=self.db["user_start"], anchor=W, justify=LEFT
            ).grid(row=r, column=1, sticky=W)
            # image = PhotoImage(file="data/img/report.png", height=50, width=50)
            # image.zoom(50, 50)
            r += 1
            sc = Button(self.settings, compound=LEFT, font="Arial 10 bold", text="Save", anchor=W, justify=LEFT,
                        command=self.update_settings)
            # rep.image = image
            sc.grid(row=r, columnspan=2, pady=20)
            r += 1

        except Exception as ee:
            print(ee)
            logging.error('An error occurred in (Settings)', exc_info=True, extra={'investigator': 'RegSmart'})
            self.display_message('error', 'An error occurred trying to save your settings.\n Please try again.')

    def rep_log(self, msg):
        tmp = dt.datetime.now().replace(microsecond=0)
        self.report_log += str(tmp) + " - [" + self.investigator + "] => " + msg + "<br/>"
        self.timeline_log.append(tmp.time())