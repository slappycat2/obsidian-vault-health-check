
# import tkinter as tk
#
# def on_option_select(event):
#     selected_option.set(event)
# root = tk.Tk()
# root.title("Dropdown Menu Example")
# root.geometry("400x300")
#
# selected_option = tk.StringVar()
#
#
# options = ["Option 1", "Option 2", "Option 3", "Option 4"]
# dropdown = tk.OptionMenu(root, selected_option, *options)
# dropdown.current(2)
# dropdown.pack(pady=10)
#
# # Add a button to display the selected option
# show_button = tk.Button(root, text="Show Selection", command=lambda: on_option_select(selected_option.get()))
# show_button.pack()
#
# # Label to display the selected option
# result_label = tk.Label(root, text="")
# result_label.pack()
# root.mainloop()

#Import Tkinter library
from tkinter import *
from tkinter import ttk
#Create an instance of Tkinter frame or window
win= Tk()
#Set the geometry of tkinter frame
win.geometry("750x250")
#Create a Combobox
combobox= ttk.Combobox(win,state= "readonly")
combobox['values']=('C++','Java','Python')
combobox.current(2)
combobox.pack(pady=30, ipadx=20)
win.mainloop()