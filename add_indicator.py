# add_indicator.py
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from database import DB_NAME
from utils import wrap_text

class AddIndicatorWindow:
    def __init__(self, root, refresh_callback):
        self.root = root
        self.refresh_callback = refresh_callback
        self.root.title("Thêm khung chỉ tiêu cơ bản mới")
        self.root.state('zoomed')
        
        # Nhập tên loại sản phẩm mới
        tk.Label(self.root, text="Tên loại sản phẩm mới:").grid(row=0, column=0, padx=5, pady=5, sticky="nw")
        self.type_text = tk.Text(self.root, width=50, height=2, wrap=tk.WORD)
        self.type_text.grid(row=0, column=1, padx=5, pady=5)
        
        self.tree_frame = tk.Frame(self.root)
        self.tree_frame.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        
        v_scrollbar = tk.Scrollbar(self.tree_frame, orient="vertical")
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree = ttk.Treeview(self.tree_frame, columns=("Mã chỉ tiêu", "Chỉ tiêu", "Giá trị", "Đơn vị", "Hành động"), show="headings",
                                 yscrollcommand=v_scrollbar.set)
        v_scrollbar.config(command=self.tree.yview)
        
        self.tree.heading("Mã chỉ tiêu", text="Mã chỉ tiêu")
        self.tree.heading("Chỉ tiêu", text="Chỉ tiêu")
        self.tree.heading("Giá trị", text="Giá trị")
        self.tree.heading("Đơn vị", text="Đơn vị")
        self.tree.heading("Hành động", text="Hành động")
        
        self.tree.column("Mã chỉ tiêu", width=100, anchor="center")
        self.tree.column("Chỉ tiêu", width=240, anchor="w")
        self.tree.column("Giá trị", width=440, anchor="w")
        self.tree.column("Đơn vị", width=150, anchor="w")
        self.tree.column("Hành động", width=100, anchor="center")
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Áp dụng style để set rowheight=60 (3 dòng) cho Treeview
        self.setup_treeview_style()
        
        self.tree.bind("<Double-1>", self.edit_inline)
        self.tree.bind("<Button-1>", self.on_click)
        
        button_frame = tk.Frame(self.root)
        button_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        tk.Button(button_frame, text="+", command=self.add_new_row, width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Lưu", command=self.save_indicators, width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Hủy", command=self.root.destroy, width=10).pack(side=tk.LEFT, padx=5)
    
    def get_units(self):
        """Lấy danh sách ký hiệu đơn vị từ bảng units"""
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        units = c.execute("SELECT ki_hieu_don_vi FROM units ORDER BY ki_hieu_don_vi").fetchall()
        conn.close()
        return [u[0] for u in units]
    
    def setup_treeview_style(self):
        """Cấu hình style cho Treeview với chiều cao hàng 3 dòng (60px)"""
        style = ttk.Style()
        style.configure("Treeview", rowheight=60)
    
    def on_click(self, event):
        """Xử lý click vào cột Hành động"""
        column = self.tree.identify_column(event.x)
        if column == "#5":  # Cột Hành động
            item = self.tree.identify_row(event.y)
            if item:
                self.tree.delete(item)
    
    def add_new_row(self):
        """Thêm hàng mới"""
        self.tree.insert("", "end", values=("", "", "", "", "Xóa"))
    
    def edit_inline(self, event):
        """SỬA: Chỉnh sửa inline - ô nhập cao bằng hàng hiển thị (60px, hỗ trợ multi-line)"""
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if not item or not column or column == "#5":  # Không chỉnh sửa Hành động
            return
        col_index = int(column[1:]) - 1
        current_value = self.tree.item(item)["values"][col_index]
        
        if column == "#4":  # Cột Đơn vị (#4)
            combo = ttk.Combobox(self.tree, values=self.get_units(), height=9)
            combo.set(current_value)
            bbox = self.tree.bbox(item, column)
            if bbox:
                x, y, width, height = bbox
                combo.place(x=x, y=y, width=width, height=60, anchor="nw")
            
            def save_combo(e=None):
                new_value = combo.get().strip()
                values = list(self.tree.item(item)["values"])
                values[col_index] = new_value
                self.tree.item(item, values=values)
                combo.destroy()
            
            combo.bind("<FocusOut>", save_combo)
            combo.bind("<Return>", save_combo)
            combo.focus_set()
        else:
            # SỬA: Sử dụng tk.Text với height=3 (khớp rowheight=60px)
            text = tk.Text(self.tree, height=3, wrap=tk.WORD, font=("Arial", 10))
            text.insert(tk.END, str(current_value))  # SỬA: Chuyển sang str() để tránh lỗi
            bbox = self.tree.bbox(item, column)
            if bbox:
                x, y, width, height = bbox
                text.place(x=x, y=y, width=width, height=60, anchor="nw")
        
            def save_text(e=None):
                new_value = text.get("1.0", tk.END).strip()
                values = list(self.tree.item(item)["values"])
                values[col_index] = new_value
                self.tree.item(item, values=values)
                text.destroy()
        
            # Bindings hỗ trợ multi-line
            text.bind("<FocusOut>", save_text)
            text.bind("<Control-Return>", save_text)
            text.focus_set()
            # SỬA: Dùng tag_add thay vì select_range
            text.tag_add("sel", "1.0", "end")
            text.mark_set("insert", "1.0")
    
    def save_indicators(self):
        """Lưu loại sản phẩm mới và chỉ tiêu"""
        type_name = self.type_text.get("1.0", tk.END).strip()
        if not type_name:
            messagebox.showerror("Lỗi", "Vui lòng nhập tên loại sản phẩm")
            return
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        existing = c.execute("SELECT id FROM product_types WHERE name=?", (type_name,)).fetchone()
        if existing:
            messagebox.showerror("Lỗi", "Tên loại sản phẩm đã tồn tại")
            conn.close()
            return
        try:
            c.execute("INSERT INTO product_types (name) VALUES (?)", (type_name,))
            type_id = c.lastrowid
            for child in self.tree.get_children():
                values = self.tree.item(child)["values"]
                # SỬA: Lấy 4 giá trị đầu, bỏ qua cột "Hành động" (cột thứ 5)
                indicator_code = str(values[0]).strip()
                indc = str(values[1]).strip()
                value = str(values[2]).strip()
                unit = str(values[3]).strip()
                
                if not indicator_code:
                    messagebox.showerror("Lỗi", "Mã chỉ tiêu bắt buộc phải có")
                    conn.rollback()
                    conn.close()
                    return
                if not indc:
                    messagebox.showerror("Lỗi", "Chỉ tiêu bắt buộc phải có")
                    conn.rollback()
                    conn.close()
                    return
                c.execute('''INSERT INTO indicators (type_id, indicator_code, indicator, value, unit)
                             VALUES (?, ?, ?, ?, ?)''', (type_id, indicator_code, indc, value, unit))
            conn.commit()
            messagebox.showinfo("Thành công", "Đã thêm chỉ tiêu mới")
            self.refresh_callback()
            self.root.destroy()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("Lỗi", f"Lỗi khi thêm: {str(e)}")
        finally:
            conn.close()