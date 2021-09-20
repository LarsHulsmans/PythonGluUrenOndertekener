import tkinter as tk
import os
from win32com import client
from openpyxl  import load_workbook
from TkinterDnD2 import DND_FILES, TkinterDnD


def drop_inside_list_box(event):
    if(event.data.endswith(".xlsx")):
        listb.insert("end", event.data)

def signCallback():
    items = listb.get(0, tk.END)
    for i in items:
        insertNameAndEmail(i)

def insertNameAndEmail(path):
    book = load_workbook(path)
    sheet = book.active

    sheet['d37'].value = nameEntry.get()
    sheet['d38'].value = emailEntry.get()

    book.save(filename=path)

    saveAsPDF(path)

def saveAsPDF(path):
    name = os.path.basename(path).replace(".xlsx", "")
    dir = os.path.dirname(path)

    excel = client.Dispatch("Excel.Application")
  
    # Read Excel File
    sheets = excel.Workbooks.Open(path)
    work_sheets = sheets.Worksheets[0]
    
    # Convert into PDF File
    work_sheets.ExportAsFixedFormat(0, f"{dir}/{name}.pdf")
    sheets.Close(True)
    listb.delete(0, 'end')

root = TkinterDnD.Tk();
root.geometry("800x500")

listb = tk.Listbox(root, selectmode=tk.SINGLE, background="#ffe0d6")
listb.pack(fill=tk.X)
listb.drop_target_register(DND_FILES)
listb.dnd_bind("<<Drop>>", drop_inside_list_box)

nameLabel = tk.Label(root, text="Name:")
nameEntry = tk.Entry(root)
nameEntry.insert(0, "Lars Hulsmans")
emailLabel = tk.Label(root, text="e-mail:")
emailEntry = tk.Entry(root)
emailEntry.insert(0, "lars.hulsmans@sensiks.com")

nameLabel.pack();
nameEntry.pack();
emailLabel.pack();
emailEntry.pack();

signb = tk.Button(root, text="Sign", command=signCallback)
signb.pack();

root.mainloop();