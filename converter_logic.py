import os
import threading
from docx2pdf import convert as convert_docx
import comtypes.client # Import for direct PPTX conversion

def _convert_pptx_to_pdf_direct(input_ppt_path, output_pdf_path):
    """Converts a single PPTX file to PDF using COM objects."""
    powerpoint = None
    presentation = None
    try:
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        # Making it visible might help with some COM issues, can be set to 0 for invisible
        # powerpoint.Visible = 1 
        presentation = powerpoint.Presentations.Open(os.path.abspath(input_ppt_path))
        presentation.SaveAs(os.path.abspath(output_pdf_path), 32) # 32 is the ppSaveAsPDF format
    finally:
        if presentation:
            presentation.Close()
        if powerpoint:
            powerpoint.Quit()

def _convert_file_thread_logic(selected_file, output_path, file_type, success_callback, error_callback):
    try:
        if file_type == ".docx":
            convert_docx(selected_file, output_path)
        elif file_type == ".pptx":
            _convert_pptx_to_pdf_direct(selected_file, output_path)
        success_callback(output_path)
    except Exception as e:
        # Ensure COM objects are released if an error occurs during their use
        # This is partly handled by the finally block in _convert_pptx_to_pdf_direct
        error_callback(f"Conversion failed: {str(e)}")

def convert_file_to_pdf(selected_file, success_callback, error_callback, status_update_callback):
    """
    Converts a DOCX or PPTX file to PDF in a separate thread.

    Args:
        selected_file (str): The path to the DOCX or PPTX file.
        success_callback (function): Callback on successful conversion.
        error_callback (function): Callback on conversion error.
        status_update_callback (function): Callback to update status message.
    """
    if not selected_file:
        error_callback("Please select a file first!")
        return

    file_name, file_extension = os.path.splitext(selected_file)
    file_extension = file_extension.lower()

    if file_extension not in (".docx", ".pptx"):
        error_callback(f"Unsupported file type: {file_extension}. Please select a DOCX or PPTX file.")
        return

    status_update_callback("Converting... Please wait.")
    output_path = file_name + ".pdf"

    thread = threading.Thread(target=_convert_file_thread_logic, args=(
        selected_file, output_path, file_extension, success_callback, error_callback
    ))
    thread.start()

# Keep the old function for now, or decide if it should be removed/deprecated
# For this refactor, we'll assume the new function replaces the old one in terms of usage from gui.py
# def convert_docx_to_pdf(selected_file, success_callback, error_callback, status_update_callback):
#     convert_file_to_pdf(selected_file, success_callback, error_callback, status_update_callback) 