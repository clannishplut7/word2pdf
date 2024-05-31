import os
from tkinter import Tk, filedialog

def select_files():
    root = Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(title="Select Word Documents", filetypes=[("Word files", "*.docx")])
    root.destroy()
    return file_paths

def convert_to_pdf(file_paths):
    for file_path in file_paths:
        # Construct the PDF output path
        filename, _ = os.path.splitext(file_path)
        pdf_path = f"{filename}.pdf"
        
        # Execute the conversion process using os.system
        command = f"docx2pdf \"{file_path}\""
        print(f"Executing command: {command}")
        result = os.system(command)
        
        # Check if the PDF file was created
        if os.path.exists(pdf_path):
            print(f"{file_path} converted to {pdf_path}")
        else:
            print(f"Failed to convert {file_path} to PDF")
            print(f"os.system command returned: {result}")

if __name__ == "__main__":
    print("Welcome to Word to PDF Converter!")
    files = select_files()
    if files:
        convert_to_pdf(files)
    else:
        print("No files selected. Exiting.")
