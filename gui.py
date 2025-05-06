import threading
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from main import main


def select_input_folder():
    folder_selected = filedialog.askdirectory()
    input_var.set(folder_selected)


def select_output_folder():
    folder_selected = filedialog.askdirectory()
    output_var.set(folder_selected)


def run_processing():
    input_path = input_var.get()
    output_folder = output_var.get()

    if not input_path or not output_folder:
        messagebox.showerror("Error", "Please select both input and output folders.")
        return

    processing_label.config(text="Processing... Please wait.", foreground="red")
    threading.Thread(target=process_pdfs, args=(input_path, output_folder), daemon=True).start()


def process_pdfs(input_path, output_folder):
    try:
        main(input_path, output_folder)
        messagebox.showinfo("Success", "Processing complete!")
    except Exception as e:
        messagebox.showerror("Error", f"Processing failed: {e}")
    finally:
        processing_label.config(text="")


# Create the main window with a Bootstrap theme
root = ttk.Window(themename="darklyC")  # Choose a Bootstrap theme # Solar, Cyborg, morph, vapor, darkly, journal
root.title("OCR Statement Process")
root.geometry("340x270")

# UI Elements
ttk.Label(root, text="Select Input Folder:", font=("Arial", 10)).pack(pady=5)
input_var = ttk.StringVar()
ttk.Entry(root, textvariable=input_var, width=50).pack(pady=2)
ttk.Button(root, text="Browse", command=select_input_folder).pack(pady=10)

ttk.Label(root, text="Select Output Folder:", font=("Arial", 10)).pack(pady=5)
output_var = ttk.StringVar()
ttk.Entry(root, textvariable=output_var, width=50).pack(pady=2)
ttk.Button(root, text="Browse", command=select_output_folder).pack(pady=5)

ttk.Button(root, text="Start Processing", command=run_processing).pack(pady=10)
processing_label = ttk.Label(root, text="", font=("Arial", 10))
processing_label.pack()

root.mainloop()
