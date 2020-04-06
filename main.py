import tkinter as tk

from gui import RunGUI

root = tk.Tk()
root.title('Money s4 to E-shop')
root.geometry('600x150')
app = RunGUI(master=root)
app.mainloop()
