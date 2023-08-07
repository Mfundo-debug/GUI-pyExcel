import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.title("pyExcel")

style = ttk.Style(root)
root.tk.call("source", "C:/Users/didit/Downloads/GUI-pyExcel/Forest-ttk-theme/forest-dark.tcl")
root.tk.call("source", "C:/Users/didit/Downloads/GUI-pyExcel/Forest-ttk-theme/forest-light.tcl")
style.theme_use("forest-dark")

root.mainloop()