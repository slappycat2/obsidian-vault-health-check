import tkinter as tk
from tkinter import ttk

"""Mapping of bbox choices based on abox selection"""
choices = {
    "A": ["1", "2", "3"],
    "B": ["4", "6", "7"],
    "C": ["8", "8", "9"],
}
def aselected(box, var, values):
    """I get called when the abox selection changes"""
    box["values"] = values
    var.set(values[0])

root = tk.Tk()

key = list(choices.keys())[0]
value = choices[key]

avar = tk.StringVar(value=key)
bvar = tk.StringVar(value=value[0])

abox = ttk.Combobox(root, textvariable=avar, values=list(choices.keys()))
bbox = ttk.Combobox(root, textvariable=bvar, values=value)

abox.bind('<<ComboboxSelected>>', lambda event: aselected(bbox, bvar, choices[avar.get()]))
bbox.bind('<<ComboboxSelected>>', lambda event: print(bvar.get()))

abox.pack()
bbox.pack()
root.mainloop()