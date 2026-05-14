import tkinter as tk
import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox
import os
import tempfile
import win32com.client
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter as openpyxl_get_col_letter
from copy import copy

# --- SETTINGS ---
ctk.set_appearance_mode("System")  # Use system light/dark mode
ctk.set_default_color_theme("blue")

# --- HELPER FUNCTIONS ---
def get_column_letter(n):
    """Convert a zero-indexed column number to Excel-style letter (0 -> A, 25 -> Z)."""
    result = ""
    while n >= 0:
        result = chr(65 + (n % 26)) + result
        n = (n // 26) - 1
    return result

def get_column_index(letter):
    """Convert Excel-style letter to zero-indexed column number (A -> 0)."""
    n = 0
    for char in letter.upper():
        n = n * 26 + (ord(char) - 64)
    return n - 1

def ensure_xlsx(file_path):
    """Chuyển đổi .xls sang .xlsx bằng win32com (nếu cần)."""
    if str(file_path).lower().endswith('.xls'):
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            abs_path = os.path.abspath(file_path)
            wb = excel.Workbooks.Open(abs_path)
            temp_dir = tempfile.gettempdir()
            base_name = os.path.basename(file_path)
            xlsx_path = os.path.join(temp_dir, base_name + "x")
            if os.path.exists(xlsx_path):
                os.remove(xlsx_path)
            wb.SaveAs(xlsx_path, FileFormat=51)
            wb.Close()
            excel.Quit()
            return xlsx_path
        except Exception as e:
            raise Exception(f"Không thể tự động chuyển đổi file .xls sang .xlsx: {str(e)}\nVui lòng mở file bằng Excel và lưu lại dưới dạng .xlsx.")
    return file_path

def copy_cell_style(source_cell, target_cell):
    """Copies style from source_cell to target_cell."""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

def get_real_max_row(ws):
    """Lấy số dòng thực sự có chứa dữ liệu thay vì max_row (để tránh trường hợp người dùng tô màu cả cột xuống 1 triệu dòng)."""
    real_max = 0
    # 1. Kiểm tra các ô đã được merge (vì ô merge chỉ chứa dữ liệu ở ô đầu tiên)
    for merged_range in ws.merged_cells.ranges:
        if merged_range.max_row > real_max:
            real_max = merged_range.max_row
            
    # 2. Quét nhanh tất cả các cell thực sự tồn tại trong bộ nhớ
    if hasattr(ws, '_cells'):
        for (r, c), cell in ws._cells.items():
            if cell.value is not None and str(cell.value).strip() != "":
                if r > real_max:
                    real_max = r
    else:
        real_max = ws.max_row
        
    return max(1, real_max)

def copy_sheet_data(source_ws, target_ws, start_row, include_header=True, header_row_count=1):
    row_offset = start_row - 1
    for col_letter, col_dimension in source_ws.column_dimensions.items():
        target_ws.column_dimensions[col_letter].width = col_dimension.width
    source_start_row = 1 if include_header else (header_row_count + 1)
    
    for row_idx, row_dimension in source_ws.row_dimensions.items():
        if row_idx >= source_start_row:
            new_row_idx = row_idx + row_offset - (0 if include_header else header_row_count)
            target_ws.row_dimensions[new_row_idx].height = row_dimension.height

    real_max_row = get_real_max_row(source_ws)
    for row in source_ws.iter_rows(min_row=source_start_row, max_row=real_max_row):
        for cell in row:
            if cell.value is None and not cell.has_style:
                continue # Bỏ qua các ô hoàn toàn trống và không có style để tối ưu hiệu suất
            new_cell = target_ws.cell(row=cell.row + row_offset - (0 if include_header else header_row_count), 
                                      column=cell.column, value=cell.value)
            copy_cell_style(cell, new_cell)
    for merged_range in source_ws.merged_cells.ranges:
        if merged_range.min_row >= source_start_row:
            new_min_row = merged_range.min_row + row_offset - (0 if include_header else header_row_count)
            new_max_row = merged_range.max_row + row_offset - (0 if include_header else header_row_count)
            new_range = f"{openpyxl_get_col_letter(merged_range.min_col)}{new_min_row}:" \
                        f"{openpyxl_get_col_letter(merged_range.max_col)}{new_max_row}"
            try: target_ws.merge_cells(new_range)
            except Exception: pass
    return get_real_max_row(target_ws)

# --- CORE LOGIC ---
def merge_selected_sheets(input_path, output_path, selected_sheet_names):
    if not selected_sheet_names:
        return 0
        
    converted_path = ensure_xlsx(input_path)
    src_wb = load_workbook(converted_path, data_only=False)
    
    # Use the first selected sheet as base
    base_sheet_name = selected_sheet_names[0]
    dst_wb = Workbook()
    # Remove default sheet
    del dst_wb["Sheet"]
    
    # Create base sheet in new workbook by copying
    src_ws_base = src_wb[base_sheet_name]
    dst_ws = dst_wb.create_sheet(title="Merged_Sheets")
    
    # Copy data from the first sheet
    current_row = copy_sheet_data(src_ws_base, dst_ws, 1, include_header=True)
    
    # Merge subsequent sheets
    for i in range(1, len(selected_sheet_names)):
        sheet_name = selected_sheet_names[i]
        ws = src_wb[sheet_name]
        current_row = copy_sheet_data(ws, dst_ws, current_row + 1, include_header=True)
        
    dst_wb.save(output_path)
    return len(selected_sheet_names)

def merge_multiple_files(file_list, output_path):
    if not file_list: return 0
    
    # Lấy file đầu tiên làm base workbook
    converted_path_0 = ensure_xlsx(file_list[0])
    dst_wb = load_workbook(converted_path_0, data_only=False)
    dst_ws = dst_wb.worksheets[0]
    dst_ws.title = "Merged_Files"
    
    current_row = get_real_max_row(dst_ws)
    sheet_names_0 = dst_wb.sheetnames.copy()
    
    # Gộp các sheet còn lại trong file đầu tiên
    for j in range(1, len(sheet_names_0)):
        ws = dst_wb[sheet_names_0[j]]
        current_row = copy_sheet_data(ws, dst_ws, current_row + 1, include_header=True)
    
    # Xóa các sheet thừa trong file đầu tiên
    for sheet_name in sheet_names_0[1:]:
        del dst_wb[sheet_name]

    # Gộp các file tiếp theo
    for i in range(1, len(file_list)):
        converted_path = ensure_xlsx(file_list[i])
        src_wb = load_workbook(converted_path, data_only=False)
        for j, sheet_name in enumerate(src_wb.sheetnames):
            ws = src_wb[sheet_name]
            current_row = copy_sheet_data(ws, dst_ws, current_row + 1, include_header=True)
            
    dst_wb.save(output_path)
    return len(file_list)

def copy_header_and_format(source_ws, target_ws, header_start, header_end):
    # Copy column width
    for col_letter, col_dimension in source_ws.column_dimensions.items():
        target_ws.column_dimensions[col_letter].width = col_dimension.width
    
    # Copy header rows (from source row header_start to target row 1)
    target_row_idx = 1
    for row_idx in range(header_start, header_end + 1):
        if row_idx in source_ws.row_dimensions:
            target_ws.row_dimensions[target_row_idx].height = source_ws.row_dimensions[row_idx].height
        
        for cell in source_ws[row_idx]:
            new_cell = target_ws.cell(row=target_row_idx, column=cell.column, value=cell.value)
            copy_cell_style(cell, new_cell)
        target_row_idx += 1
            
    # Copy merged cells for header
    for merged_range in source_ws.merged_cells.ranges:
        if merged_range.min_row >= header_start and merged_range.max_row <= header_end:
            new_min_row = merged_range.min_row - header_start + 1
            new_max_row = merged_range.max_row - header_start + 1
            new_range = f"{openpyxl_get_col_letter(merged_range.min_col)}{new_min_row}:" \
                        f"{openpyxl_get_col_letter(merged_range.max_col)}{new_max_row}"
            try: target_ws.merge_cells(new_range)
            except: pass

def copy_data_chunk(source_ws, target_ws, rows_chunk, header_count):
    start_target_row = header_count + 1
    current_target_row = start_target_row
    row_mapping = {} # maps source_row -> target_row for merged cells
    
    for row in rows_chunk:
        source_row_idx = row[0].row
        row_mapping[source_row_idx] = current_target_row
        
        if source_row_idx in source_ws.row_dimensions:
            target_ws.row_dimensions[current_target_row].height = source_ws.row_dimensions[source_row_idx].height
            
        for cell in row:
            if cell.value is None and not cell.has_style:
                continue
            new_cell = target_ws.cell(row=current_target_row, column=cell.column, value=cell.value)
            copy_cell_style(cell, new_cell)
        current_target_row += 1
        
    # Copy merged cells within this chunk
    for merged_range in source_ws.merged_cells.ranges:
        min_r, max_r = merged_range.min_row, merged_range.max_row
        if min_r in row_mapping and max_r in row_mapping:
            new_range = f"{openpyxl_get_col_letter(merged_range.min_col)}{row_mapping[min_r]}:" \
                        f"{openpyxl_get_col_letter(merged_range.max_col)}{row_mapping[max_r]}"
            try: target_ws.merge_cells(new_range)
            except: pass

def clean_sheet_name(name):
    s_name = str(name)[:30].strip()
    for char in ['/', '\\', '?', '*', ':', '[', ']']: s_name = s_name.replace(char, '')
    return s_name if s_name else "Sheet"

def split_excel(input_path, output_path_base, mode, output_mode, header_start, header_end, sheet_name=None, column_index=None, row_count=None):
    converted_path = ensure_xlsx(input_path)
    wb = load_workbook(converted_path, data_only=False)
    if not wb.sheetnames: return 0
    
    if sheet_name and sheet_name in wb.sheetnames:
        source_ws = wb[sheet_name]
    else:
        source_ws = wb.active
    real_max_row = get_real_max_row(source_ws)
    header_count = header_end - header_start + 1
    chunks = {} # name -> list of rows
    
    # Data starts from header_end + 1
    if mode == "column":
        for row in source_ws.iter_rows(min_row=header_end + 1, max_row=real_max_row):
            val = row[column_index].value
            key = str(val) if val is not None else "Blank"
            if key not in chunks: chunks[key] = []
            chunks[key].append(row)
    else: # row_count
        current_chunk, count = 1, 0
        current_rows = []
        for row in source_ws.iter_rows(min_row=header_end + 1, max_row=real_max_row):
            current_rows.append(row)
            count += 1
            if count >= row_count:
                chunks[f"Phần_{current_chunk}"] = current_rows
                current_chunk += 1
                current_rows, count = [], 0
        if current_rows: chunks[f"Phần_{current_chunk}"] = current_rows
            
    if not chunks: return 0
    
    if output_mode == "sheets":
        for name, rows in chunks.items():
            safe_name = clean_sheet_name(name)
            base_name, counter = safe_name, 1
            while safe_name in wb.sheetnames:
                safe_name = f"{base_name}_{counter}"
                counter += 1
            new_ws = wb.create_sheet(title=safe_name)
            copy_header_and_format(source_ws, new_ws, header_start, header_end)
            copy_data_chunk(source_ws, new_ws, rows, header_count)
        wb.save(output_path_base)
        return len(chunks)
    else: # multiple files
        base_dir = os.path.dirname(output_path_base)
        base_name_file, ext = os.path.splitext(os.path.basename(output_path_base))
        for name, rows in chunks.items():
            new_wb = Workbook()
            new_ws = new_wb.active
            new_ws.title = clean_sheet_name(name)
            copy_header_and_format(source_ws, new_ws, header_start, header_end)
            copy_data_chunk(source_ws, new_ws, rows, header_count)
            out_file = os.path.join(base_dir, f"{base_name_file}_{clean_sheet_name(name)}{ext}")
            new_wb.save(out_file)
        return len(chunks)

# --- UI APPLICATION ---
class ExcelMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Excel Tool - Duy Hieu - Professional")
        self.geometry("1100x700")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Sidebar
        self.navigation_frame = ctk.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(4, weight=1)
        
        # Explicit text color for the sidebar header
        self.label_header = ctk.CTkLabel(self.navigation_frame, text="  EXCEL TOOLBOX", 
                                         font=ctk.CTkFont(size=15, weight="bold"),
                                         text_color=("black", "white"))
        self.label_header.grid(row=0, column=0, padx=20, pady=20)

        # Sidebar buttons with explicit text colors for unselected state
        self.home_button = ctk.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, 
                                         text="Ghép Sheet (1 File)", fg_color="transparent", 
                                         text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                         anchor="w", command=lambda: self.select_frame_by_name("ghép_sheet"))
        self.home_button.grid(row=1, column=0, sticky="ew")

        self.frame_2_button = ctk.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, 
                                            text="Ghép Nhiều File", fg_color="transparent", 
                                            text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                            anchor="w", command=lambda: self.select_frame_by_name("ghép_nhiều_file"))
        self.frame_2_button.grid(row=2, column=0, sticky="ew")

        self.frame_3_button = ctk.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, 
                                            text="Tách File Excel", fg_color="transparent", 
                                            text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                            anchor="w", command=lambda: self.select_frame_by_name("tách_file"))
        self.frame_3_button.grid(row=3, column=0, sticky="ew")

        # Frames
        self.home_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent"); self.setup_home_frame()
        self.second_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent"); self.setup_second_frame()
        self.third_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent"); self.setup_third_frame()
        
        self.select_frame_by_name("ghép_sheet")

    def setup_home_frame(self):
        ctk.CTkLabel(self.home_frame, text="Ghép Sheet (1 File) - Giữ nguyên định dạng & Merge", font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, padx=20, pady=20)
        self.input_file_path = ctk.StringVar()
        
        file_frame = ctk.CTkFrame(self.home_frame, fg_color="transparent")
        file_frame.grid(row=1, column=0, padx=20, pady=10)
        ctk.CTkEntry(file_frame, textvariable=self.input_file_path, width=400).grid(row=0, column=0, padx=5)
        ctk.CTkButton(file_frame, text="Chọn File", command=self.browse_single_file).grid(row=0, column=1, padx=5)
        
        self.sheet_info_label = ctk.CTkLabel(self.home_frame, text="Chọn các sheet cần ghép:", text_color=("gray40", "gray60"))
        self.sheet_info_label.grid(row=2, column=0, pady=(10, 0))
        
        # Listbox for sheet selection
        list_frame = ctk.CTkFrame(self.home_frame)
        list_frame.grid(row=3, column=0, padx=20, pady=5)
        
        self.home_sheet_listbox = tk.Listbox(list_frame, selectmode="multiple", width=60, height=8, bg="#333333", fg="white", selectbackground="#2ecc71")
        self.home_sheet_listbox.pack(side="left", fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.home_sheet_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.home_sheet_listbox.config(yscrollcommand=scrollbar.set)
        
        ctk.CTkButton(self.home_frame, text="Chọn tất cả Sheet", command=self.select_all_home_sheets, width=150).grid(row=4, column=0, pady=5)
        
        ctk.CTkButton(self.home_frame, text="Bắt đầu Ghép Sheet", command=self.merge_sheets_process, height=50, fg_color="#2ecc71").grid(row=5, column=0, pady=20)

    def select_all_home_sheets(self):
        self.home_sheet_listbox.select_set(0, tk.END)

    def setup_second_frame(self):
        ctk.CTkLabel(self.second_frame, text="Ghép nhiều file Excel (Giữ nguyên định dạng & Merge)", font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, padx=20, pady=20)
        self.file_list = []
        self.file_listbox = tk.Listbox(self.second_frame, width=80, height=10, bg="#333333", fg="white")
        self.file_listbox.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        file_ops = ctk.CTkFrame(self.second_frame, fg_color="transparent")
        file_ops.grid(row=1, column=1, padx=20, pady=10, sticky="n")
        ctk.CTkButton(file_ops, text="Thêm Files", command=self.browse_multiple_files).grid(row=0, column=0, pady=5)
        ctk.CTkButton(file_ops, text="Xóa", fg_color="#e74c3c", command=self.clear_file_list).grid(row=1, column=0, pady=5)
        ctk.CTkButton(self.second_frame, text="Bắt đầu Ghép File", command=self.merge_files_process, height=50, fg_color="#3498db").grid(row=2, column=0, columnspan=2, pady=30)

    def setup_third_frame(self):
        ctk.CTkLabel(self.third_frame, text="Tách File Excel (Giữ nguyên định dạng)", font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, columnspan=2, padx=20, pady=20)
        self.split_input_path = ctk.StringVar()
        file_f = ctk.CTkFrame(self.third_frame, fg_color="transparent")
        file_f.grid(row=1, column=0, columnspan=2, padx=20, pady=10)
        ctk.CTkEntry(file_f, textvariable=self.split_input_path, width=400).grid(row=0, column=0, padx=5)
        ctk.CTkButton(file_f, text="Chọn File", command=self.browse_split_file).grid(row=0, column=1, padx=5)
        
        # Sheet selection for splitting
        sheet_f = ctk.CTkFrame(self.third_frame, fg_color="transparent")
        sheet_f.grid(row=2, column=0, columnspan=2, padx=20, pady=5)
        ctk.CTkLabel(sheet_f, text="Chọn Sheet cần tách:").grid(row=0, column=0, padx=5)
        self.split_sheet_var = ctk.StringVar(value="Active Sheet")
        self.split_sheet_menu = ctk.CTkOptionMenu(sheet_f, variable=self.split_sheet_var, values=["Active Sheet"], command=lambda _: self.reload_columns())
        self.split_sheet_menu.grid(row=0, column=1, padx=5)
        
        # Split options frame
        options_frame = ctk.CTkFrame(self.third_frame, fg_color="transparent")
        options_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        # Mode selection
        self.split_mode_var = ctk.StringVar(value="column")
        mode_frame = ctk.CTkFrame(options_frame)
        mode_frame.grid(row=0, column=0, padx=10, sticky="n")
        ctk.CTkLabel(mode_frame, text="Phương thức tách:", font=ctk.CTkFont(weight="bold")).pack(pady=5)
        ctk.CTkRadioButton(mode_frame, text="Theo giá trị cột", variable=self.split_mode_var, value="column", command=self.toggle_split_inputs).pack(anchor="w", padx=10, pady=5)
        ctk.CTkRadioButton(mode_frame, text="Theo số lượng dòng", variable=self.split_mode_var, value="row_count", command=self.toggle_split_inputs).pack(anchor="w", padx=10, pady=5)
        
        # Output selection
        self.split_output_var = ctk.StringVar(value="sheets")
        output_frame = ctk.CTkFrame(options_frame)
        output_frame.grid(row=0, column=1, padx=10, sticky="n")
        ctk.CTkLabel(output_frame, text="Hình thức xuất:", font=ctk.CTkFont(weight="bold")).pack(pady=5)
        ctk.CTkRadioButton(output_frame, text="Nhiều Sheet trong 1 File", variable=self.split_output_var, value="sheets").pack(anchor="w", padx=10, pady=5)
        ctk.CTkRadioButton(output_frame, text="Nhiều File riêng biệt", variable=self.split_output_var, value="files").pack(anchor="w", padx=10, pady=5)
        
        # Inputs frame
        self.inputs_frame = ctk.CTkFrame(self.third_frame, fg_color="transparent")
        self.inputs_frame.grid(row=4, column=0, columnspan=2, pady=5)
        
        ctk.CTkLabel(self.inputs_frame, text="Tiêu đề từ dòng:").grid(row=0, column=0, padx=2)
        self.header_start_entry = ctk.CTkEntry(self.inputs_frame, width=40); self.header_start_entry.insert(0, "1"); self.header_start_entry.grid(row=0, column=1, padx=2)
        ctk.CTkLabel(self.inputs_frame, text="đến:").grid(row=0, column=2, padx=2)
        self.header_end_entry = ctk.CTkEntry(self.inputs_frame, width=40); self.header_end_entry.insert(0, "1"); self.header_end_entry.grid(row=0, column=3, padx=2)
        
        # Dynamic inputs
        self.split_header_menu = ctk.CTkOptionMenu(self.inputs_frame, values=["Chưa có dữ liệu"], width=180)
        self.split_header_menu.grid(row=0, column=4, padx=5)
        
        self.row_count_entry = ctk.CTkEntry(self.inputs_frame, width=150, placeholder_text="Số dòng mỗi phần...")
        self.row_count_entry.grid_remove() 
        
        ctk.CTkButton(self.inputs_frame, text="Tải lại cột", width=80, command=self.reload_columns).grid(row=0, column=5, padx=5)
        
        ctk.CTkButton(self.third_frame, text="Bắt đầu Tách File", command=self.split_process, height=50, fg_color="#9b59b6").grid(row=5, column=0, columnspan=2, pady=30)

    def toggle_split_inputs(self):
        mode = self.split_mode_var.get()
        if mode == "column":
            self.split_header_menu.grid()
            self.row_count_entry.grid_remove()
        else:
            self.split_header_menu.grid_remove()
            self.row_count_entry.grid(row=0, column=4, padx=5)

    def select_frame_by_name(self, name):
        # Configure button colors with high contrast for selected and unselected states
        self.home_button.configure(fg_color=("gray75", "gray25") if name == "ghép_sheet" else "transparent",
                                   text_color=("black", "white") if name == "ghép_sheet" else ("gray10", "gray90"))
        self.frame_2_button.configure(fg_color=("gray75", "gray25") if name == "ghép_nhiều_file" else "transparent",
                                      text_color=("black", "white") if name == "ghép_nhiều_file" else ("gray10", "gray90"))
        self.frame_3_button.configure(fg_color=("gray100", "gray0") if name == "tách_file" else "transparent", # Use explicit contrast
                                      text_color=("black", "white") if name == "tách_file" else ("gray10", "gray90"))
        
        self.home_frame.grid_forget(); self.second_frame.grid_forget(); self.third_frame.grid_forget()
        if name == "ghép_sheet": self.home_frame.grid(row=0, column=1, sticky="nsew")
        elif name == "ghép_nhiều_file": self.second_frame.grid(row=0, column=1, sticky="nsew")
        elif name == "tách_file": self.third_frame.grid(row=0, column=1, sticky="nsew")

    def browse_single_file(self):
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if f: 
            self.input_file_path.set(f)
            try:
                wb_info = pd.ExcelFile(f)
                sheet_names = wb_info.sheet_names
                self.home_sheet_listbox.delete(0, tk.END)
                for name in sheet_names:
                    self.home_sheet_listbox.insert(tk.END, name)
                # Select all by default
                self.home_sheet_listbox.select_set(0, tk.END)
            except Exception as e:
                messagebox.showerror("Lỗi", str(e))

    def browse_split_file(self):
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if f: 
            self.split_input_path.set(f)
            try:
                wb_info = pd.ExcelFile(f)
                sheet_names = wb_info.sheet_names
                self.split_sheet_menu.configure(values=sheet_names)
                if sheet_names:
                    self.split_sheet_var.set(sheet_names[0])
                self.reload_columns()
            except Exception as e:
                messagebox.showerror("Lỗi", str(e))

    def reload_columns(self):
        f = self.split_input_path.get()
        if not f or not os.path.exists(f): return
        try:
            h_end = int(self.header_end_entry.get())
            selected_sheet = self.split_sheet_var.get()
            # Nếu chưa có sheet nào được chọn hoặc đang để mặc định
            s_name = selected_sheet if selected_sheet != "Active Sheet" else 0
            
            df = pd.read_excel(f, sheet_name=s_name, header=h_end-1, nrows=1, dtype=str)
            self.split_header_menu.configure(values=[f"{get_column_letter(i)} | {c}" for i, c in enumerate(df.columns)])
            if not df.columns.empty:
                self.split_header_menu.set(f"{get_column_letter(0)} | {df.columns[0]}")
        except Exception as e: messagebox.showerror("Lỗi", str(e))

    def split_process(self):
        input_file = self.split_input_path.get()
        if not input_file:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file cần tách!")
            return

        try:
            h_start = int(self.header_start_entry.get())
            h_end = int(self.header_end_entry.get())
            if h_start < 1 or h_end < h_start:
                messagebox.showwarning("Cảnh báo", "Dải dòng tiêu đề không hợp lệ!")
                return
                
            mode = self.split_mode_var.get()
            output_mode = self.split_output_var.get()
            selected_sheet = self.split_sheet_var.get()
            
            row_count = None
            col_index = None
            
            if mode == "column":
                selected_val = self.split_header_menu.get()
                if selected_val == "Chưa có dữ liệu" or not selected_val:
                    messagebox.showwarning("Cảnh báo", "Vui lòng tải lại và chọn cột để tách!")
                    return
                col_letter = selected_val.split(" | ")[0]
                col_index = get_column_index(col_letter)
            else:
                try:
                    row_count = int(self.row_count_entry.get())
                    if row_count <= 0: raise ValueError
                except:
                    messagebox.showwarning("Cảnh báo", "Số dòng mỗi phần phải là số nguyên dương!")
                    return

            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if not save_path: return

            num_parts = split_excel(input_file, save_path, mode, output_mode, h_start, h_end, 
                                    sheet_name=selected_sheet, column_index=col_index, row_count=row_count)
            messagebox.showinfo("Thành công", f"Đã tách xong thành {num_parts} phần!")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {str(e)}")

    def merge_sheets_process(self):
        input_path = self.input_file_path.get()
        if not input_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file!")
            return
            
        selected_indices = self.home_sheet_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất 1 sheet để ghép!")
            return
            
        selected_sheets = [self.home_sheet_listbox.get(i) for i in selected_indices]
        
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            try:
                count = merge_selected_sheets(input_path, output_path, selected_sheets)
                messagebox.showinfo("Thành công", f"Đã ghép xong {count} sheet!")
            except Exception as e: messagebox.showerror("Lỗi", str(e))

    def browse_multiple_files(self):
        for f in filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")]):
            if f not in self.file_list: self.file_list.append(f); self.file_listbox.insert(tk.END, os.path.basename(f))

    def clear_file_list(self): self.file_list = []; self.file_listbox.delete(0, tk.END)

    def merge_files_process(self):
        if not self.file_list: return
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if output_path:
            try:
                count = merge_multiple_files(self.file_list, output_path)
                messagebox.showinfo("Thành công", f"Đã ghép xong {count} file!")
            except Exception as e: messagebox.showerror("Lỗi", str(e))

if __name__ == "__main__":
    app = ExcelMergerApp()
    app.mainloop()
