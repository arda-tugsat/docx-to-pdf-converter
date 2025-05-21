import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk # Import ttk for themed widgets
import os
from converter_logic import convert_file_to_pdf

class DocxToPdfConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("File to PDF Converter")
        self.root.geometry("1280x720") # Changed window size
        self.root.configure(bg="#f0f0f0")

        # Style configuration for ttk widgets
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", font=("Helvetica", 10))
        style.configure("TLabel", padding=5, background="#f0f0f0", font=("Helvetica", 10))
        style.configure("Title.TLabel", font=("Helvetica", 16, "bold"), padding=10)
        style.configure("Status.TLabel", font=("Helvetica", 10, "italic"))

        # Create main frame with ttk style
        self.main_frame = ttk.Frame(root, padding="20 20 20 20")
        self.main_frame.pack(fill="both", expand=True)

        # Title
        self.title_label = ttk.Label(
            self.main_frame,
            text="File to PDF Converter",
            style="Title.TLabel"
        )
        self.title_label.pack(pady=(0, 20)) # More padding below title

        # File selection button
        self.select_button = ttk.Button(
            self.main_frame,
            text="Select DOCX or PPTX File",
            command=self.select_file,
            style="TButton"
        )
        self.select_button.pack(pady=10)

        # Selected file label
        self.file_label = ttk.Label(
            self.main_frame,
            text="No file selected",
            wraplength=550, # Adjust wraplength for new window size
            style="TLabel"
        )
        self.file_label.pack(pady=10)

        # Convert button
        self.convert_button = ttk.Button(
            self.main_frame,
            text="Convert to PDF",
            command=self.start_conversion,
            state="disabled",
            style="TButton"
        )
        self.convert_button.pack(pady=20)
        
        # Progress bar (indeterminate)
        self.progress_bar = ttk.Progressbar(
            self.main_frame,
            orient="horizontal",
            length=300,
            mode="indeterminate"
        )
        # self.progress_bar.pack_forget() # Initially hidden, will be shown during conversion
        self.progress_bar.pack(pady=(0,10))
        self.progress_bar.pack_forget() # Hide it initially

        # Status label
        self.status_label = ttk.Label(
            self.main_frame,
            text="",
            style="Status.TLabel"
        )
        self.status_label.pack(pady=10)

        self.selected_file = None

    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Office Documents", "*.docx *.pptx"),
                ("Word Documents", "*.docx"),
                ("PowerPoint Presentations", "*.pptx"),
                ("All Files", "*.*")
            ]
        )
        if file_path:
            self.selected_file = file_path
            self.file_label.config(text=f"Selected: {os.path.basename(file_path)}") # Show only filename
            self.convert_button.config(state="normal")
            self.status_label.config(text="File selected. Ready to convert.")
        else:
            self.selected_file = None
            self.file_label.config(text="No file selected")
            self.convert_button.config(state="disabled")
            self.status_label.config(text="")

    def start_conversion(self):
        if not self.selected_file:
            messagebox.showerror("Error", "Please select a file first!")
            return

        self.convert_button.config(state="disabled")
        self.select_button.config(state="disabled") # Disable select button during conversion
        self.status_label.config(text="Converting... Please wait.")
        self.progress_bar.pack(pady=(5,10)) # Show progress bar
        self.progress_bar.start(10) # Start indeterminate animation
        
        convert_file_to_pdf(
            self.selected_file,
            self._conversion_complete,
            self._conversion_error,
            self._update_status # This callback might not be strictly needed if progress bar handles "Converting..."
        )

    def _update_status(self, message):
        # This will primarily be used by convert_file_to_pdf if it sends intermediate statuses
        # For now, it sets the initial "Converting..." message.
        self.status_label.config(text=message)

    def _conversion_complete(self, output_path):
        self.root.after(0, self.__conversion_complete_ui_updates, output_path)

    def __conversion_complete_ui_updates(self, output_path):
        self.progress_bar.stop()
        self.progress_bar.pack_forget() # Hide progress bar
        self.status_label.config(text="Conversion completed successfully!")
        self.convert_button.config(state="normal")
        self.select_button.config(state="normal") # Re-enable select button
        messagebox.showinfo("Success", f"PDF saved as:\n{output_path}")
        # Reset for next conversion
        self.file_label.config(text="Select another file or close.")
        self.selected_file = None 
        # self.convert_button.config(state="disabled") # Optionally disable convert until new file selected

    def _conversion_error(self, error_message):
        self.root.after(0, self.__conversion_error_ui_updates, error_message)

    def __conversion_error_ui_updates(self, error_message):
        self.progress_bar.stop()
        self.progress_bar.pack_forget() # Hide progress bar
        self.status_label.config(text=error_message) 
        self.convert_button.config(state="normal")
        self.select_button.config(state="normal") # Re-enable select button
        messagebox.showerror("Error", error_message) 