import tkinter as tk

from gui import RunGUI

root = tk.Tk()
root.title('Money s4 to E-shop')
root.geometry('600x200')
app = RunGUI(master=root)
app.mainloop()
