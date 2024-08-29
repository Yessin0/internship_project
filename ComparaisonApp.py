from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import comparatif
import os
import platform


def print_author_name():
    python_version = platform.python_version()
    print("********************************************")
    print("*                                          *")
    print("*     Application Name: Comparatif Tool     *")
    print("*                                          *")
    print(f"*     Author: [YESSIN TOUMI]                   *")
    print(f"*     Date: 2024                        *")
    print(f"*     Python Version: {python_version}                  *")
    print("*     All rights reserved.                  *")
    print("*                                          *")
    print("********************************************")
    print("")


class ComparaisonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparison App")

        # DA File Selection
        self.da_file_path = tk.StringVar()
        ttk.Label(root, text="Select DA File:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        self.da_entry = ttk.Entry(root, textvariable=self.da_file_path)
        self.da_entry.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(root, text="Browse DA File", command=self.browse_da_file).grid(row=0, column=2, padx=10, pady=5)

        # Folder Selection
        ttk.Label(root, text="Select Folder:").grid(row=1, column=0)
        self.folder_list = ttk.Combobox(root, state="readonly")
        self.folder_list.grid(row=1, column=1, padx=10, pady=5)
        ttk.Button(root, text="Browse Folder", command=self.browse_folder).grid(row=1, column=2, padx=10, pady=5)

        # Compare Button
        ttk.Button(root, text="Compare", command=self.compare_files).grid(row=2, column=1, padx=10, pady=20)

        # Add Direction Achat text at the bottom left
        self.direction_label = ttk.Label(root, text="Direction Achat", font=("Helvetica", 10, "bold"))
        self.direction_label.grid(row=3, column=0, padx=10, pady=20, sticky=tk.W)

        # Add Ooredoo image at the bottom right
        self.load_image()

    def browse_da_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.da_file_path.set(file_path)

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            subfolders = [d for d in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, d))]
            self.folder_list['values'] = subfolders
            if subfolders:
                self.folder_list.current(0)
            self.selected_folder_path = folder_path

    def compare_files(self):
        da_file_path = self.da_file_path.get()
        selected_subfolder = self.folder_list.get()

        if not da_file_path:
            messagebox.showerror("Error", "Please select a DA file")
            return
        if not selected_subfolder:
            messagebox.showerror("Error", "Please select a folder")
            return

        selected_subfolder = os.path.join(self.selected_folder_path, selected_subfolder)
        if not os.path.exists(da_file_path):
            messagebox.showerror("Error", f"DA file does not exist: {da_file_path}")
            return
        if not os.path.isdir(selected_subfolder):
            messagebox.showerror("Error", f"Selected folder does not exist: {selected_subfolder}")
            return

        try:
            comparatif.process_files(da_file_path, selected_subfolder)
            messagebox.showinfo("Success", "Comparison complete, report generated.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during comparison: {e}")

    def load_image(self):  # load the ooredoo logo
        try :
            image = Image.open("ooredoo.png")  # Load the image
            resized_image = image.resize((100, 100), Image.LANCZOS)  # Resize the image
            self.image = ImageTk.PhotoImage(resized_image)
            self.image_label = tk.Label(self.root, image=self.image)
            self.image_label.grid(row=3, column=2, padx=10, pady=20, sticky=tk.E)
        except Exception as e:
            messagebox.showerror("Error",f"An error occurred during the loading: {e}")


if __name__ == "__main__":
    print_author_name()
    root = tk.Tk()
    app = ComparaisonApp(root)
    root.mainloop()
