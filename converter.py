import tkinter as tk
from gui import DocxToPdfConverter # Import from the new gui.py

if __name__ == "__main__":
    root = tk.Tk()
    app = DocxToPdfConverter(root)
    root.mainloop() 