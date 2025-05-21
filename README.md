# File to PDF Converter

A simple and reliable application to convert DOCX and PPTX files to PDF format while maintaining formatting and layout.

## Prerequisites

1. Python 3.7 or higher
2. Microsoft Office (Word and PowerPoint must be installed on your system)
3. Required Python packages (see `requirements.txt`)

## Installation

1. Ensure you have Python installed on your system.
2. Clone this repository or download the source files.
3. Install the required packages by navigating to the project directory in your terminal and running:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the converter by executing the following command in the project directory:
   ```bash
   python converter.py
   ```
2. The application window will appear (default size 1280x720).
3. Click "Select DOCX or PPTX File" to choose your input file.
4. The selected filename will be displayed.
5. Click "Convert to PDF" to start the conversion.
6. A progress bar will indicate that the conversion is in process.
7. Upon completion, a success message will be shown, and the PDF will be saved in the same directory as the source file.
8. Error messages will be displayed if any issues occur during the conversion.

## Features

- User-friendly graphical interface with themed widgets (`tkinter.ttk`).
- Supports conversion of both DOCX (Microsoft Word) and PPTX (Microsoft PowerPoint) files to PDF.
- Aims to maintain original formatting and layout by utilizing Microsoft Office's COM automation.
- Indeterminate progress bar during conversion.
- Clear status messages and error handling.
- Background conversion process (non-blocking UI) thanks to threading.
- Modular code structure (`gui.py`, `converter_logic.py`, `converter.py`).

## Important Note

- This converter relies on **Microsoft Office (Word and PowerPoint) being installed** on your Windows system. It uses COM automation to interact with these applications for high-fidelity conversion.
- The `comtypes` library is used for PowerPoint conversion, and `docx2pdf` is used for Word conversion. 