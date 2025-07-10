from customtkinter import *

app = CTk()
app.geometry("555x333")
app.title("Using progress bar")

progress = CTkProgressBar(app,width=400,height=15,corner_radius=5,progress_color="green")
progress.pack(padx=30, pady=50)
progress.set(0)

def start_progress():
    progress.set(0)
    app.after(100, update_progress, 0)

def update_progress(value):
    if value <= 1.0:
        progress.set(value)
        app.after(100, update_progress, value + 0.01)

start_button = CTkButton(app, text="Start Work", command=start_progress)
start_button.pack(pady=20)

app.mainloop()