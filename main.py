import os
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox, Text, Scrollbar, RIGHT, Y, END, BOTH
from docx2pdf import convert

def read_excel(file_path):
    df = pd.read_excel(file_path)
    log_text.insert(END, f"Columns in the Excel file: {df.columns.tolist()}\n")
    return df

def find_coa_file(code, source_folder):
    for root, _, files in os.walk(source_folder):
        for file in files:
            if code.lower() in file.lower() and file.startswith("COA+") and file.endswith(".docx"):
                return os.path.join(root, file)
    return None

def set_font(paragraph, font_name, font_size, bold=False):
    for run in paragraph.runs:
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(font_size)
        run.font.bold = bold

def update_docx(input_path, output_path, product_code, product_name, lot_no, sap_code, date, re_assay_date, storage):
    try:
        doc = Document(input_path)
        
        # Update first page
        for paragraph in doc.paragraphs:
            if "Lot No. :" in paragraph.text:
                paragraph.text = f"Lot No. : {lot_no}"
                set_font(paragraph, font_name="Arial", font_size=11)
                break
        
        # Update second page
        for paragraph in doc.paragraphs:
            if "The following is a statement for material:" in paragraph.text:
                paragraph.text = f"The following is a statement for material: {product_name}"
                set_font(paragraph, font_name="Calibri", font_size=12, bold=True)
            elif "Product code:" in paragraph.text:
                paragraph.text = f"Product code: {product_code} ({sap_code})"
                set_font(paragraph, font_name="Calibri", font_size=12, bold=True)
            elif "Lot#" in paragraph.text:
                paragraph.text = f"Lot# {lot_no}"
                set_font(paragraph, font_name="Calibri", font_size=12, bold=True)
                ############################ 
            elif "Re-Assay Date:" in paragraph.text and pd.notna(re_assay_date):
                paragraph.text = f"Re-Assay Date: {str(re_assay_date)}"
                set_font(paragraph, font_name="Calibri", font_size=12)

            elif "Storage:" in paragraph.text and pd.notna(storage):
                paragraph.text = f"Storage: {str(storage)}"
                set_font(paragraph, font_name="Calibri", font_size=12)

            elif "Date:" in paragraph.text and pd.notna(date):
                paragraph.text = f"Date: {str(date)}"
                set_font(paragraph, font_name="Calibri", font_size=12)

        
        doc.save(output_path)
        return True
    except Exception as e:
        log_text.insert(END, f"Error updating Word document {input_path}: {str(e)}\n", 'error')
        return False

def process_files():
    excel_path = excel_path_entry.get()
    source_folder = source_folder_entry.get()
    destination_folder = destination_folder_entry.get()

    if not excel_path or not source_folder or not destination_folder:
        messagebox.showerror("Error", "All fields must be filled!")
        return

    try:
        df = read_excel(excel_path)
        log_text.insert(END, f"Columns in the Excel file: {df.columns.tolist()}\n")

        # Define the expected columns
        expected_columns = ['Product Code', 'Description', 'Batch', 'Material', 'Date', 'Re-Assay Date', 'Storage']

        # Check if all required columns exist
        missing_columns = [col for col in expected_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing columns in Excel file: {', '.join(missing_columns)}")

        for index, row in df.iterrows():
            try:
                code = str(row['Product Code'])
                description = str(row['Description'])
                lot_no = str(row['Batch'])
                sap_code = str(row['Material'])
  
                date = row['Date']
                re_assay_date = row['Re-Assay Date']
                storage = row['Storage']

                print('DEBUG', date, re_assay_date, storage)

                log_text.insert(END, f"Processing: Code={code}, Description={description}, Lot={lot_no}, SAP Code={sap_code}\n")

                coa_file = find_coa_file(code, source_folder)
                if coa_file:
                    new_file_name = f"COA+{code} ({sap_code})+{lot_no}.docx"
                    output_path = os.path.join(destination_folder, new_file_name)
                    
                    if update_docx(coa_file, output_path, code, description, lot_no, sap_code, date, re_assay_date, storage):
                        pdf_output_path = output_path.replace(".docx", ".pdf")
                        convert(output_path, pdf_output_path)
                        log_text.insert(END, f"Updated and converted COA for {code}\n")
                    else:
                        log_text.insert(END, f"Failed to update COA for {code}\n", 'error')
                else:
                    log_text.insert(END, f"COA file not found for {code}\n", 'error')
            except Exception as e:
                log_text.insert(END, f"Error processing row {index}: {row.to_dict()}\n", 'error')
                log_text.insert(END, f"Error details: {str(e)}\n", 'error')
                log_text.insert(END, f"Failed to generate COA for product code: {code}\n", 'error')

        messagebox.showinfo("Success", "Processing completed.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
        log_text.insert(END, f"Critical error: {str(e)}\n", 'error')
        
def browse_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    entry.delete(0, 'end')
    entry.insert(0, file_path)

def browse_folder(entry):
    folder_path = filedialog.askdirectory()
    entry.delete(0, 'end')
    entry.insert(0, folder_path)

# GUI setup
root = Tk()
root.title("COA Updater")

# Excel file selection
Label(root, text="Excel File Path:").grid(row=0, column=0, padx=10, pady=5)
excel_path_entry = Entry(root, width=50)
excel_path_entry.grid(row=0, column=1, padx=10, pady=5)
Button(root, text="Browse...", command=lambda: browse_file(excel_path_entry)).grid(row=0, column=2, padx=10, pady=5)

# Source folder selection
Label(root, text="Source Folder Path:").grid(row=1, column=0, padx=10, pady=5)
source_folder_entry = Entry(root, width=50)
source_folder_entry.grid(row=1, column=1, padx=10, pady=5)
Button(root, text="Browse...", command=lambda: browse_folder(source_folder_entry)).grid(row=1, column=2, padx=10, pady=5)

# Destination folder selection
Label(root, text="Destination Folder Path:").grid(row=2, column=0, padx=10, pady=5)
destination_folder_entry = Entry(root, width=50)
destination_folder_entry.grid(row=2, column=1, padx=10, pady=5)
Button(root, text="Browse...", command=lambda: browse_folder(destination_folder_entry)).grid(row=2, column=2, padx=10, pady=5)

# Log window
log_text = Text(root, height=15, width=80)
log_text.grid(row=3, column=0, columnspan=3, padx=10, pady=10)
scrollbar = Scrollbar(root, command=log_text.yview)
log_text.config(yscrollcommand=scrollbar.set)
scrollbar.grid(row=3, column=3, sticky='nsew')

# Create a tag to style error messages in red
log_text.tag_config('error', foreground='red')

# Process button
Button(root, text="Process Files", command=process_files, width=20).grid(row=4, column=1, pady=20)

root.mainloop()