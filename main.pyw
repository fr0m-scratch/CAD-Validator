from utility.utility import *
from cad_entity import CADEntity
from cad_handler import CADHandler
import rich.progress
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
from ttkthemes import ThemedTk
import os
import sys
from tkterminal import Terminal
import threading
from rich.traceback import install
from validation import validate
from tkinter.font import Font
install()

class Application(ttk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.geometry("800x600")
        self.pack(padx=10, pady=10)
        self.create_widgets()
        

    def create_widgets(self):
        self.label = ttk.Label(self, text="比对文件路径")
        self.label.pack()

        self.entry = ttk.Entry(self, width=80)
        self.entry.pack(pady=5)

        self.browse_button = ttk.Button(self, text="选择比对文件", command=self.browse_file)
        self.browse_button.pack(pady=5)

        self.convert_button = ttk.Button(self, text="生成输出文件", command=self.start_conversion)
        self.convert_button.pack(pady=5)

        self.result_label = ttk.Label(self, text="输出文件路径")
        self.result_label.pack(pady=5)

        self.result_entry = ttk.Entry(self, width=80)
        self.result_entry.pack(pady=5)

        self.open_button = ttk.Button(self, text="打开输出文件", command=self.open_file)
        self.open_button.pack(pady=5)
        
        self.terminal = Terminal(pady=5, padx = 5, 
                                 font = 'Courier 16', 
                                 background = 'black', 
                                 foreground = 'white',
                                 height = 20,
                                 width = 45)
        self.terminal.basename = "CAD Validation"
        self.terminal.tag_config('basename', foreground='white', font = 'Courier 16')
        self.terminal.pack(pady = 5)
        cl = ConsoleLog(self.terminal)
        sys.stdout = cl
        

    def browse_file(self):
        filename = filedialog.askopenfilename()
        self.entry.delete(0, tk.END)
        self.entry.insert(tk.END, filename)

    def convert_file(self):
        source_file = self.entry.get()
        if not os.path.isfile(source_file):
            messagebox.showerror("Error", "Please select a valid file")
            return

        destination_file = validate(source_file)
        
        self.terminal.insert(tk.END, "文件比对成功，并存入: " + destination_file + "\n")
        self.result_entry.delete(0, tk.END)
        self.result_entry.insert(tk.END, destination_file)

    def start_conversion(self):
        thread = threading.Thread(target=self.convert_file)
        thread.start()

    def open_file(self):
        file_to_open = self.result_entry.get()
        if not os.path.isfile(file_to_open):
            messagebox.showerror("Error", "No converted file to open")
            return
        os.startfile(file_to_open)

class ConsoleLog:
    def __init__(self, console):
        self.console = console

    def write(self, message):
        self.console.insert(tk.END, message)

    def flush(self):
        pass

root = ThemedTk(theme="arc")
app = Application(master=root)

app.mainloop()
