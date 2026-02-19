import os
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox, Text, END, Scrollbar
from docx2pdf import convert
import pythoncom
import sys

# 关键修复：重定向标准输出，防止 docx2pdf 在无控制台环境下崩溃
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")


def copy_paragraph_format(src_para, tgt_para):
    """Deep copy paragraph formatting for perfect alignment."""
    tgt_para.paragraph_format.alignment = src_para.paragraph_format.alignment
    tgt_para.paragraph_format.space_before = src_para.paragraph_format.space_before
    tgt_para.paragraph_format.space_after = src_para.paragraph_format.space_after
    tgt_para.paragraph_format.line_spacing = src_para.paragraph_format.line_spacing
    tgt_para.paragraph_format.line_spacing_rule = src_para.paragraph_format.line_spacing_rule
    tgt_para.paragraph_format.left_indent = src_para.paragraph_format.left_indent
    tgt_para.paragraph_format.right_indent = src_para.paragraph_format.right_indent
    tgt_para.paragraph_format.first_line_indent = src_para.paragraph_format.first_line_indent


def set_run_font(run, font_name="Arial", font_size=11):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = Pt(font_size)


def update_docx(input_path, output_path, lot_no, manufacturing_date):
    try:
        doc = Document(input_path)
        for table in doc.tables:
            if len(table.rows) == 1 and len(table.columns) >= 2:
                left_cell = table.cell(0, 0)
                right_cell = table.cell(0, 1)
                left_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
                right_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

                reassay_idx = -1
                for i, p in enumerate(left_cell.paragraphs):
                    if "re-assay" in p.text.lower() or "reassay" in p.text.lower():
                        reassay_idx = i
                        break
                if reassay_idx != -1:
                    p_to_del = left_cell.paragraphs[reassay_idx]
                    p_to_del._element.getparent().remove(p_to_del._element)

                original_text = right_cell.text.replace('\r', '\n')
                lines = [line.strip() for line in original_text.split('\n') if line.strip()]
                target_data = ["", "", "", str(lot_no)]
                if len(lines) > 0: target_data[0] = lines[0]
                if len(lines) > 1: target_data[1] = lines[1]
                if len(lines) > 2: target_data[2] = lines[2]

                formats = [p for p in left_cell.paragraphs]
                right_cell.text = ""
                for i in range(len(formats)):
                    new_p = right_cell.paragraphs[0] if i == 0 else right_cell.add_paragraph()
                    copy_paragraph_format(formats[i], new_p)
                    val = target_data[i] if i < len(target_data) else ""
                    run = new_p.add_run(val)
                    set_run_font(run)

        try:
            dtn = pd.to_datetime(manufacturing_date)
            formatted_date = dtn.strftime("%b %d, %Y")
        except:
            formatted_date = str(manufacturing_date)

        for p in doc.paragraphs:
            if "date:" in p.text.lower() and "re-assay" not in p.text.lower():
                p.text = f"Date: {formatted_date}"
                for run in p.runs:
                    set_run_font(run, font_size=10.5)

        doc.save(output_path)
        return True
    except Exception as e:
        return f"Error: {str(e)}"


def process_files():
    e_path = excel_entry.get()
    s_folder = src_entry.get()
    d_folder = dst_entry.get()

    if not all([e_path, s_folder, d_folder]):
        messagebox.showwarning("Warning", "Please select all required paths.")
        return

    log_text.delete(1.0, END)
    log_text.insert(END, ">>> Starting Process...\n")
    root.update()

    try:
        # 强制初始化 COM
        pythoncom.CoInitialize()

        df = pd.read_excel(e_path)
        total_rows = len(df)
        success_count = 0

        for index, row in df.iterrows():
            p_code = str(row["Product Code"]).strip()
            l_no = str(row["Lot / Batch"]).strip()
            m_date = row["Manufacturing Date"]

            log_text.insert(END, f"[{index + 1}/{total_rows}] {p_code}: ")
            log_text.see(END)
            root.update()

            template = None
            for root_dir, _, files in os.walk(s_folder):
                for f in files:
                    if p_code.lower() in f.lower() and f.endswith(".docx"):
                        template = os.path.join(root_dir, f)
                        break

            if not template:
                log_text.insert(END, "Template Missing\n")
                continue

            new_name = f"{p_code}-{l_no.replace('/', '-')}-C6.docx"
            out_path = os.path.join(d_folder, new_name)

            result = update_docx(template, out_path, l_no, m_date)
            if result is True:
                try:
                    # 调用转换
                    convert(out_path, out_path.replace(".docx", ".pdf"))
                    log_text.insert(END, "SUCCESS (DOCX+PDF)\n")
                except Exception as pdf_err:
                    log_text.insert(END, f"DOCX ONLY (PDF Failed: {str(pdf_err)})\n")
                success_count += 1
            else:
                log_text.insert(END, f"FAILED ({result})\n")
            root.update()

        log_text.insert(END,
                        f"\n{'=' * 40}\nAll Tasks Completed!\nSuccessfully Generated: {success_count} files\n{'=' * 40}\n")
        log_text.see(END)
        messagebox.showinfo("Done", "Process Completed Successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"Fatal error: {str(e)}")
    finally:
        pythoncom.CoUninitialize()


# UI Setup
root = Tk()
root.title("COA Auto-Generator Pro V15")
root.geometry("850x600")  # 宽度增加到 850
root.resizable(False, False)

# Labels & Entries - 优化间距，防止文字遮挡
label_x = 20
entry_x = 220  # 将输入框起始点进一步后移
entry_width = 55

Label(root, text="Excel Data File:", font=("Arial", 10, "bold")).place(x=label_x, y=20)
excel_entry = Entry(root, width=entry_width, font=("Arial", 10))
excel_entry.place(x=entry_x, y=20)
Button(root, text="Browse", width=10, command=lambda: excel_entry.insert(0, filedialog.askopenfilename())).place(x=730,
                                                                                                                 y=17)

Label(root, text="Source Folder Path:", font=("Arial", 10, "bold")).place(x=label_x, y=65)
src_entry = Entry(root, width=entry_width, font=("Arial", 10))
src_entry.place(x=entry_x, y=65)
Button(root, text="Browse", width=10, command=lambda: src_entry.insert(0, filedialog.askdirectory())).place(x=730, y=62)

Label(root, text="Destination Folder Path:", font=("Arial", 10, "bold")).place(x=label_x, y=110)
dst_entry = Entry(root, width=entry_width, font=("Arial", 10))
dst_entry.place(x=entry_x, y=110)
Button(root, text="Browse", width=10, command=lambda: dst_entry.insert(0, filedialog.askdirectory())).place(x=730,
                                                                                                            y=107)

# Log Area
log_frame = Label(root)
log_frame.place(x=20, y=160)
log_text = Text(log_frame, height=21, width=113, font=("Consolas", 9), bg="#F5F5F5")
log_text.pack(side="left", fill="both")
scrollbar = Scrollbar(log_frame, command=log_text.yview)
scrollbar.pack(side="right", fill="y")
log_text.config(yscrollcommand=scrollbar.set)

Button(root, text="Process Files", command=process_files, bg="#1B5E20", fg="white", font=("Arial", 12, "bold"),
       width=30, height=2).place(x=280, y=515)

root.mainloop()