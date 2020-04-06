import tkinter as tk
from tkinter import filedialog

from copy_sheets import Script
from new import NewSheet


class RunGUI(tk.Frame):
    entry_eshop_text = None
    entry_money_text = None

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.create_widgets()
        self.pack()

    def create_widgets(self):
        frame_money = tk.Frame(self, width=300, height=300)
        frame_eshop = tk.Frame(self, width=300, height=300)
        frame_money.grid(pady=5)
        frame_eshop.grid(pady=5)
        bottom_frame = tk.Frame(self)
        bottom_frame.grid(pady=5)

        money_label = tk.Label(frame_money, text='File name of money s4 sheet')
        money_label.pack(side=tk.LEFT)

        money_button = tk.Button(frame_money, text='Choose money s4 EXCEL file', command=self.get_money_sheet_filename)
        money_button.pack(side=tk.RIGHT)

        self.entry_money_text = tk.StringVar()
        entry_money = tk.Entry(frame_money, width=40, bd=3, state=tk.DISABLED, textvariable=self.entry_money_text)
        entry_money.pack(side=tk.RIGHT)

        eshop_label = tk.Label(frame_eshop, text='File name of e-shop sheet')
        eshop_label.pack(side=tk.LEFT, anchor='w')

        self.entry_eshop_text = tk.StringVar()
        entry_eshop = tk.Entry(frame_eshop, width=40, bd=3, state=tk.DISABLED, textvariable=self.entry_eshop_text)
        eshop_button = tk.Button(frame_eshop, text='Choose e-shop EXCEL file', command=self.get_eshop_sheet_filename)
        eshop_button.pack(side=tk.RIGHT)
        entry_eshop.pack(side=tk.RIGHT)

        run_script_button = tk.Button(bottom_frame, text='Create new sheet', command=self.run_script)
        run_script_button.pack(side=tk.RIGHT)

    def get_money_sheet_filename(self):
        filename = filedialog.askopenfilename(initialdir='/', title='Select excel file from money',
                                              filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))
        self.entry_money_text.set(filename.replace('/', '//'))

    def get_eshop_sheet_filename(self):
        filename = filedialog.askopenfilename(initialdir='/', title='Select excel file from money',
                                              filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))
        self.entry_eshop_text.set(filename.replace('/', '//'))

    def run_script(self):
        # script = Script(self.entry_money_text.get(), self.entry_eshop_text.get())
        # script.build_new_sheet(lang=0)
        # script.build_new_sheet(lang=1)

        new_sheet = NewSheet(self.entry_money_text.get())
        new_sheet.create_new_sheet()
        label = tk.Label(self, text='Files have been created')
        label.grid()
