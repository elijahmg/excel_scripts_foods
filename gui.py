import tkinter as tk
from tkinter import filedialog

from new import NewSheet


class RunGUI(tk.Frame):
    entry_eshop_text = None
    entry_money_text = None

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.v = tk.IntVar()
        self.info_label_text = None
        self.info_label = None
        self.create_widgets()
        self.pack()

    def create_widgets(self):
        frame_money = tk.Frame(self, width=300, height=500)
        frame_eshop = tk.Frame(self, width=300, height=300)
        frame_money.grid(pady=5)
        frame_eshop.grid(pady=5)
        bottom_frame = tk.Frame(self, height=300)
        bottom_frame.grid(pady=10)

        tk.Radiobutton(frame_eshop, text='Czech config', value=0, variable=self.v, command=self.v.set(0)).pack(
            side=tk.TOP, ipady=5)
        tk.Radiobutton(frame_eshop, text='English config', value=1, variable=self.v, command=self.v.set(1)).pack(
            side=tk.TOP, ipady=5)

        money_label = tk.Label(frame_money, text='File name of money s4 sheet')
        money_label.pack(side=tk.LEFT)

        money_button = tk.Button(frame_money, text='Choose money s4 EXCEL file', command=self.get_money_sheet_filename)
        money_button.pack(side=tk.RIGHT)

        self.entry_money_text = tk.StringVar()
        entry_money = tk.Entry(frame_money, width=40, bd=3, state=tk.DISABLED, textvariable=self.entry_money_text)
        entry_money.pack(side=tk.RIGHT)

        run_script_button = tk.Button(bottom_frame, text='Create new sheet', command=self.run_script)
        run_script_button.pack(side=tk.RIGHT)

        self.info_label_text = tk.StringVar()
        self.info_label = tk.Label(self, textvariable=self.info_label_text).grid()

    def get_money_sheet_filename(self):
        filename = filedialog.askopenfilename(title='Select excel file from money',
                                              filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))
        self.entry_money_text.set(filename.replace('/', '//'))

    def run_script(self):
        try:
            new_sheet = NewSheet(self.entry_money_text.get(), self.v.get())
            new_sheet.create_new_sheet()
            self.info_label_text.set('Files have been created')
        except KeyError:
            self.info_label_text.set('!!! NO CONFIG FILE !!!')
