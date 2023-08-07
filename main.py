# create the main window
import tkinter as tk
from tkinter import ttk
import openpyxl as xl
from openpyxl import Workbook
from tkinter import filedialog, messagebox, simpledialog
from tkinter.ttk import Treeview

#load the data
def load_data():
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if path == "":
        messagebox.showerror("Error", "No file selected")
        return
    else:
        try:
            global wb
            wb = xl.load_workbook(path)
            global sheet
            sheet = wb.active
            messagebox.showinfo("Success", "Data loaded successfully")
        except:
            messagebox.showerror("Error", "Invalid file")

    list_values = list(sheet.values)
    print(list_values)
    tree = ttk.Treeview(root)
    treeview = tree
    tree = ttk.Treeview(root)
    treeview = tree
    for col_name in list_values[0]:
        tree.insert("", tk.END, values=col_name)
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        tree.insert("", tk.END, values=value_tuple)

#insert row
def insert_row():
    name = name_entry.get()
    age = int(age_spinbox.get())
    subscription_status = status_combobox.get()
    employment_status = "Employed" if a.get() else "Unemployed"

    print(name, age, subscription_status, employment_status)

    #insert into excel
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if path == "":
        messagebox.showerror("Error", "No file selected")
        return
    else:
        try:
            wb = xl.load_workbook(path)
            sheet = wb.active
            sheet.append([name, age, subscription_status, employment_status])
            wb.save(path)
            messagebox.showinfo("Success", "Data inserted successfully")
        except:
            messagebox.showerror("Error", "Invalid file")

    #insert into treeview
    treeview.insert("", tk.END, values=(name, age, subscription_status, employment_status))

    #clear the entries
    name_entry.delete(0, tk.END)
    age_spinbox.delete(0, tk.END)
    age_spinbox.insert(0,'Age')
    status_combobox.set("Status")
    status_combobox.set(combo_list[0])
    checkbutton.state(["!selected"])

def toggle_mode():
    if mode_switch.instate(['selected']):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")
root = tk.Tk()
style = ttk.Style(root)

root.tk.call("source", "C:/Users/didit/Downloads/GUI-pyExcel/Forest-ttk-theme/forest-dark.tcl")
root.tk.call("source", "C:/Users/didit/Downloads/GUI-pyExcel/Forest-ttk-theme/forest-light.tcl")
style.theme_use("forest-dark")

combo_list = ["Subscribed", "Unsubscribed","Student"]

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Insert Data")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

name_entry = ttk.Entry(widgets_frame, width=30)
name_entry.insert(0, "Name")
name_entry.bind("<FocusIn>", lambda args: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, padx=10, pady=(0,5), sticky="ew")

age_spinbox = ttk.Spinbox(widgets_frame, from_=0, to=100, width=28)
age_spinbox.insert(0, "Age")
age_spinbox.bind("<FocusIn>", lambda args: age_spinbox.delete('0', 'end'))
age_spinbox.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

status_combobox = ttk.Combobox(widgets_frame, value=combo_list, width=27)
status_combobox.set("Status")
status_combobox.current(0)
status_combobox.bind("<<ComboboxSelected>>", lambda args: status_combobox.delete('0', 'end'))
status_combobox.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Employed", variable=a)
checkbutton.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

button = ttk.Button(widgets_frame, text="Insert", command=insert_row)
button.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame, orient="horizontal")
separator.grid(row=5, column=0, padx=5, pady=5, sticky="ew")

mode_switch = ttk.Checkbutton(widgets_frame, text="Light Mode", command=toggle_mode)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")

treeview_frame = ttk.LabelFrame(frame, text="Data")
treeview_frame.grid(row=0, column=1,pady=10)
treeScroll = ttk.Scrollbar(treeview_frame)
treeScroll.pack(side=tk.RIGHT, fill=tk.Y)

cols = ('Name', 'Age', 'Subscription Status', 'Employment Status')
treeview = ttk.Treeview(treeview_frame, columns=cols, show='headings', yscrollcommand=treeScroll.set)
treeview.column("Name", width=100)
treeview.column("Age", width=100)
treeview.column("Subscription Status", width=100)
treeview.column("Employment Status", width=100)
treeview.pack(side=tk.LEFT, fill=tk.Y)
treeScroll.config(command=treeview.yview)
load_data()

root.mainloop()