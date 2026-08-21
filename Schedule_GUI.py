import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
import pyexcel_io.writers
import pyexcel_io.readers
from Schedule_automation import run_excel_stuff

# Global variable for file path
selected_file_path = ""

def update_file_selection(file_path):
    global selected_file_path

    if file_path.startswith('{') and file_path.endswith('}'):
        file_path = file_path[1:-1]

    if file_path.lower().endswith(('.xlsx', '.xls')):
        selected_file_path = file_path
        label_status.config(text=f"Selected: {file_path}", fg="green")
        btn_run.config(state="normal")
    else:
        messagebox.showerror("Invalid File", "Please select valid Excel file")

def browse_file():
    #global selected_file_path

    # Open file browser
    file_path = filedialog.askopenfilename(
        title = "Select input Schedule File",
        filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", ".")]
    )
    if file_path:
        update_file_selection(file_path)

def handle_drop(event):
    if event.data:
        update_file_selection(event.data)

def process_file():
    if selected_file_path:

        label_status.config(text="Processing your file.... Please wait.", fg="orange")
        btn_run.config(state="disabled")
        btn_browse.config(state="disabled")

        progress_bar.pack(pady=10, fill="x", padx=50)
        progress_bar.start(10)

        threading.Thread(target=bg_worker, args=(selected_file_path,), daemon=True).start()
    else:
        messagebox.showwarning("Error", "Please select a file first.")

def bg_worker(file_path):
    try:
        run_excel_stuff(file_path)
        root.after(0, completion_success)

    except Exception as e:
        root.after(0, completion_error, str(e))

def completion_success():
    progress_bar.stop()
    progress_bar.pack_forget()

    label_status.config(text="Done! Schedule is ready.", fg="green")
    btn_run.config(state="normal")
    btn_browse.config(state="normal")
    messagebox.showinfo("Success", "Process finished without errors.")

def completion_error(error_msg):
    progress_bar.stop()
    progress_bar.pack_forget()

    label_status.config(text="Error occurred during processing, please check.", fg="red")
    btn_run.config(state="normal")
    btn_browse.config(state="normal")
    messagebox.showerror("Error", f"Something went wrong:\n{error_msg}")

# initialise
root = TkinterDnD.Tk()
root.title("Schedule Automation System")
root.geometry("500x420")

drop_canvas = tk.Canvas(root, bg="#f8f9fa", highlightthickness=0)
drop_canvas.pack(fill="both", expand=True, padx=25, pady=15)

def draw_dashed_border(event):
    drop_canvas.delete("border")
    w, h = event.width, event.height
    drop_canvas.create_rectangle(5, 5, w-5, h-5, outline="#999999", width=2, dash=(4, 4), tags="border")

drop_canvas.bind("<Configure>", draw_dashed_border)

drop_canvas.create_text(
    225, 60,
    text="Drag & Drop Your Excel File Here\n\n— OR —",
    justify="center",
    font=("Arial", 11),
    fill="#333333"
)

btn_browse = tk.Button(drop_canvas, text="Browse File", command=browse_file, bg="#e0e0e0")
btn_browse.pack(pady=25)

drop_canvas.create_window(225, 130, window=btn_browse)

drop_canvas.drop_target_register(DND_FILES)
drop_canvas.dnd_bind('<<Drop>>', handle_drop)

label_status = tk.Label(root, text="No file selected", fg="red", wraplength=400)
label_status.pack(pady=5)

progress_bar = ttk.Progressbar(root, orient="horizontal", mode="indeterminate", length="300")

btn_run = tk.Button(root, text="Run Automation", command=process_file, state="disabled", fg="green", font=("Arial", 11, "bold") )
btn_run.pack(pady=15)

root.mainloop()

