from customtkinter import *
root = CTk()
root.title("Entrybox")
root.geometry("888x444")
root.withdraw()

file_name = filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel file","*.xlsx")],title="Save Your Workbook")
if file_name:
    print(file_name)
else:
    print(file_name)
