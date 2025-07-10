from customtkinter import *
root = CTk()
root.title("Entrybox")
root.geometry("888x444")

entry = CTkEntry(master=root)
entry.pack(pady=30,padx=40)

root.mainloop()