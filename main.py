import os
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox, Text, Scrollbar, RIGHT, Y, END, BOTH
from docx2pdf import convert
from docx.shared import Pt
import moment


def read_excel(file_path):
    df = pd.read_excel(file_path)
    log_text.insert(END, f"Columns in the Excel file: {df.columns.tolist()}\n")
    return df


def find_coa_file(code, source_folder):
    # code = Product Code
    for root, _, files in os.walk(source_folder):
        for file in files:
            if code.lower() in file.lower() and file.endswith(".docx"):
                return os.path.join(root, file)
    return None


def set_font(paragraph, font_name, font_size, bold=False):
    for run in paragraph.runs:
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(font_size)
        run.font.bold = bold


# if update_docx(coa_file, output_path, product_code, product_description, lot_no, sap_code, manufacturing_date, expiry_re_assay_date_date):
def update_docx(input_path, output_path, product_code, lot_no, date, re_assay_date):
    try:
        doc = Document(input_path)

        for row_index, row in enumerate(doc.tables[0].rows):
            for col_index, cell in enumerate(row.cells):
                if row_index == 0 and col_index == 1:
                    # text = cell.text + f'{lot_no}\n{str(moment.date(re_assay_date).format('YYYY-MM-DD'))}\n'
                    # cell.text = ''

                    for paragraph in cell.paragraphs:
                        p = paragraph._element
                        p.getparent().remove(p)

                    paragraph = cell.add_paragraph()
                    paragraph.text = text
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

                    set_font(paragraph, font_name="Arial", font_size=11, bold=False)
                    paragraph_format = paragraph.paragraph_format
                    paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

                elif row_index == 0 and col_index == 0:
                    for paragraph in cell.paragraphs:
                        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

        for paragraph in doc.paragraphs:
            # print('paragraph.text', paragraph.text)
            # if "Lot No." in paragraph.text:
            #     paragraph.text = f"Lot# {lot_no}"
            #     set_font(paragraph, font_name="Calibri", font_size=12, bold=True)
            #     ############################
            # elif "Re-assay Date" in paragraph.text:
            #     paragraph.text = f"Re-Assay Date: {str(re_assay_date)}"
            #     set_font(paragraph, font_name="Arial", font_size=11)

            if "Date" in paragraph.text:
                print('date', date)
                paragraph.text = f"Date: {str(moment.date(date).format('MMM D, YYYY'))}"

                set_font(paragraph, font_name="Arial", font_size=10.5)

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
        expected_columns = ['Product Code', 'Lot / Batch', 'Manufacturing Date', 'Expiry Date/ Re-Assay Date']

        # Check if all required columns exist
        missing_columns = [col for col in expected_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing columns in Excel file: {', '.join(missing_columns)}")

        for index, row in df.iterrows():
            try:
                product_code = str(row['Product Code'])
                # sap_code = str(row['SAP/ Material ID']).strip()
                # product_description = str(row['Product Description'])
                lot_no = str(row['Lot / Batch'])
                # manufacturing_date = row['Manufacturing Date']
                # expiry_re_assay_date_date = row['Expiry Date/ Re-Assay Date']

                log_text.insert(END,
                                f"Processing: Product Code={product_code}, Lot / Batch={lot_no}, Manufacturing Date={manufacturing_date}, Expiry Date/ Re-Assay Date= \n")

                coa_file = find_coa_file(product_code, source_folder)

                if coa_file:
                    new_file_name = f"{product_code}-{lot_no}-with ED-C6).docx"

                    # if not os.path.exists(os.path.join(destination_folder, customer_po)):
                    #     os.makedirs(os.path.join(destination_folder, customer_po))

                    output_path = os.path.join(destination_folder, new_file_name)

                    if update_docx(coa_file, output_path, product_code, lot_no):
                        pdf_output_path = output_path.replace(".docx", ".pdf")
                        convert(output_path, pdf_output_path)
                        log_text.insert(END, f"Updated and converted COA for {product_code}\n")
                    else:
                        log_text.insert(END, f"Failed to update COA for {product_code}\n", 'error')
                else:
                    log_text.insert(END, f"COA file not found for {product_code}\n", 'error')
            except Exception as e:
                log_text.insert(END, f"Error processing row {index}: {row.to_dict()}\n", 'error')
                log_text.insert(END, f"Error details: {str(e)}\n", 'error')
                log_text.insert(END, f"Failed to generate COA for product code: {product_code}\n", 'error')

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
root.title("Buffer COA Generator")

# Excel file selection
Label(root, text="Excel File Path:").grid(row=0, column=0, padx=10, pady=5)
excel_path_entry = Entry(root, width=50)
excel_path_entry.grid(row=0, column=1, padx=10, pady=5)
Button(root, text="Browse...", command=lambda: browse_file(excel_path_entry)).grid(row=0, column=2, padx=10, pady=5)

# Source folder selection
Label(root, text="Source Folder Path:").grid(row=1, column=0, padx=10, pady=5)
source_folder_entry = Entry(root, width=50)
source_folder_entry.grid(row=1, column=1, padx=10, pady=5)
Button(root, text="Browse...", command=lambda: browse_folder(source_folder_entry)).grid(row=1, column=2, padx=10,
                                                                                        pady=5)

# Destination folder selection
Label(root, text="Destination Folder Path:").grid(row=2, column=0, padx=10, pady=5)
destination_folder_entry = Entry(root, width=50)
destination_folder_entry.grid(row=2, column=1, padx=10, pady=5)
Button(root, text="Browse...", command=lambda: browse_folder(destination_folder_entry)).grid(row=2, column=2, padx=10,
                                                                                             pady=5)

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