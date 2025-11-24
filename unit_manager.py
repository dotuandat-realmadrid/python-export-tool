# unit_manager.py
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from database import DB_NAME
from utils import wrap_text

class UnitManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Quản lý đơn vị")
        self.root.state('zoomed')
        
        # Tạo main frame
        self.main_frame = tk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        title_label = tk.Label(self.main_frame, text="Danh sách đơn vị", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Treeview frame
        tree_frame = tk.LabelFrame(self.main_frame, padx=5, pady=5)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        v_scrollbar = tk.Scrollbar(tree_frame, orient="vertical")
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        h_scrollbar = tk.Scrollbar(tree_frame, orient="horizontal")
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.tree = ttk.Treeview(tree_frame, columns=("Đại lượng", "Tên đơn vị", "Ký hiệu đơn vị", "Hành động"), 
                                 show="headings", yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)
        
        self.tree.heading("Đại lượng", text="Đại lượng")
        self.tree.heading("Tên đơn vị", text="Tên đơn vị")
        self.tree.heading("Ký hiệu đơn vị", text="Ký hiệu đơn vị")
        self.tree.heading("Hành động", text="Hành động")
        
        self.tree.column("Đại lượng", width=300, anchor="w")
        self.tree.column("Tên đơn vị", width=300, anchor="center")
        self.tree.column("Ký hiệu đơn vị", width=150, anchor="center")
        self.tree.column("Hành động", width=100, anchor="center")
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Bind events
        self.tree.bind("<Double-1>", self.on_double_click_edit)
        self.tree.bind("<Button-1>", self.on_click)
        
        # Buttons frame
        button_frame = tk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        tk.Button(button_frame, text="+", command=self.add_new_row).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Lưu", command=self.save_new_units).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Import từ Excel", command=self.import_excel).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Thoát", command=self.root.destroy).pack(side=tk.RIGHT, padx=5)
        
        # Load units initially
        self.load_units()
    
    def load_units(self):
        """Tải danh sách đơn vị từ DB"""
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        units = c.execute("SELECT id, dai_luong, ten_don_vi, ki_hieu_don_vi FROM units ORDER BY dai_luong").fetchall()
        self.tree.delete(*self.tree.get_children())
        max_height = 25
        for unit in units:
            ind_id = str(unit[0])
            dai_luong_wrapped = wrap_text(unit[1], 40)
            ten_don_vi_wrapped = wrap_text(unit[2], 40)
            ki_hieu_wrapped = wrap_text(unit[3], 20)
            line_count = max(dai_luong_wrapped.count('\n') + 1, ten_don_vi_wrapped.count('\n') + 1, ki_hieu_wrapped.count('\n') + 1)
            self.tree.insert("", "end", iid=ind_id, values=(dai_luong_wrapped, ten_don_vi_wrapped, ki_hieu_wrapped, "Xóa"))
            height = line_count * 60
            max_height = max(max_height, height)
        
        style = ttk.Style()
        style.configure("Treeview", rowheight=max_height)
        conn.close()
    
    def on_click(self, event):
        """Xử lý click vào cột Hành động"""
        column = self.tree.identify_column(event.x)
        if column == "#4":  # Cột Hành động
            item = self.tree.identify_row(event.y)
            if item:
                if item.startswith("new_"):
                    self.tree.delete(item)
                else:
                    self.delete_unit(item)
    
    def delete_unit(self, item):
        """Xóa đơn vị từ DB với xác nhận"""
        unit_id = int(item)
        if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa đơn vị này?"):
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            try:
                c.execute("DELETE FROM units WHERE id=?", (unit_id,))
                conn.commit()
                self.load_units()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi xóa: {str(e)}")
            finally:
                conn.close()
    
    def add_new_row(self):
        """Thêm hàng mới để thêm đơn vị"""
        current_children = self.tree.get_children()
        stt = len(current_children) + 1
        new_iid = f"new_{stt}"
        self.tree.insert("", "end", iid=new_iid, values=("", "", "", "Hủy"))
    
    def save_new_units(self):
        """Lưu các đơn vị mới thêm"""
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        try:
            for child in list(self.tree.get_children()):
                if child.startswith("new_"):
                    values = self.tree.item(child)['values']
                    dai_luong_wrapped = values[0]
                    ten_don_vi_wrapped = values[1]
                    ki_hieu_wrapped = values[2]
                    dai_luong = dai_luong_wrapped.replace('\n', ' ')
                    ten_don_vi = ten_don_vi_wrapped.replace('\n', ' ')
                    ki_hieu = ki_hieu_wrapped.replace('\n', ' ')
                    if not ki_hieu:
                        messagebox.showerror("Lỗi", "Ký hiệu đơn vị bắt buộc phải có")
                        return
                    c.execute('''INSERT OR IGNORE INTO units (dai_luong, ten_don_vi, ki_hieu_don_vi)
                                 VALUES (?, ?, ?)''', (dai_luong, ten_don_vi, ki_hieu))
                    self.tree.delete(child)
            conn.commit()
            messagebox.showinfo("Thành công", "Cập nhật thành công!")
            self.load_units()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("Lỗi", f"Lỗi khi lưu: {str(e)}")
        finally:
            conn.close()
    
    def on_double_click_edit(self, event):
        """Bắt đầu chỉnh sửa khi double-click vào ô"""
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if item and column != "#4":  # Không chỉnh sửa Hành động
            self.start_inline_edit(item, column, event.x, event.y)
    
    def start_inline_edit(self, item, column, x, y):
        """Bắt đầu inline editing"""
        column_index = int(column[1:]) - 1  # 0-based
        current_value = self.tree.set(item, column).replace('\n', ' ')
        self.edit_entry = tk.Entry(self.tree)
        self.edit_entry.insert(0, current_value)
        bbox = self.tree.bbox(item, column)
        if bbox:
            entry_x, entry_y, entry_width, entry_height = bbox
            self.edit_entry.place(x=entry_x, y=entry_y, width=entry_width, height=entry_height, anchor="nw")
        self.edit_entry.focus_set()
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.bind("<Return>", lambda e: self.save_inline_edit(item, column, column_index))
        self.edit_entry.bind("<FocusOut>", lambda e: self.save_inline_edit(item, column, column_index))
        self.edit_entry.bind("<Escape>", lambda e: self.cancel_inline_edit())
        self.current_edit_item = item
        self.current_edit_column = column
        self.current_edit_column_index = column_index
    
    def save_inline_edit(self, item, column, column_index):
        """Lưu thay đổi inline editing"""
        if hasattr(self, 'edit_entry') and self.edit_entry:
            new_value = self.edit_entry.get().strip()
            wrapped_value = wrap_text(new_value, 40 if column_index < 2 else 20)
            self.tree.set(item, column, wrapped_value)
            original_value = new_value
            if not item.startswith("new_"):  # Chỉ cập nhật DB cho hàng hiện có
                conn = sqlite3.connect(DB_NAME)
                c = conn.cursor()
                try:
                    unit_id = int(item)
                    if column_index == 0:  # Đại lượng
                        c.execute("UPDATE units SET dai_luong=? WHERE id=?", (original_value, unit_id))
                    elif column_index == 1:  # Tên đơn vị
                        c.execute("UPDATE units SET ten_don_vi=? WHERE id=?", (original_value, unit_id))
                    elif column_index == 2:  # Ký hiệu đơn vị
                        c.execute("UPDATE units SET ki_hieu_don_vi=? WHERE id=?", (original_value, unit_id))
                    conn.commit()
                    self.load_units()
                except Exception as e:
                    messagebox.showerror("Lỗi", f"Lỗi khi cập nhật: {str(e)}")
                finally:
                    conn.close()
            self.edit_entry.destroy()
            self.edit_entry = None
        self.current_edit_item = None
        self.current_edit_column = None
        self.current_edit_column_index = None
    
    def cancel_inline_edit(self):
        """Hủy chỉnh sửa inline"""
        if hasattr(self, 'edit_entry') and self.edit_entry:
            self.edit_entry.destroy()
            self.edit_entry = None
        self.current_edit_item = None
        self.current_edit_column = None
        self.current_edit_column_index = None
    
    def import_excel(self):
        """Import dữ liệu từ file Excel"""
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file:
            return
        try:
            df = pd.read_excel(file)
            # Bỏ cột TT nếu tồn tại
            if 'TT' in df.columns:
                df = df.drop('TT', axis=1)
            # Giả sử cột còn lại: Đại lượng, Tên đơn vị, Ký hiệu đơn vị
            required_columns = ['Đại lượng', 'Tên đơn vị', 'Ký hiệu đơn vị']
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("Lỗi", "File Excel phải có cột: Đại lượng, Tên đơn vị, Ký hiệu đơn vị")
                return
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            for _, row in df.iterrows():
                dai_luong = row['Đại lượng']
                ten_don_vi = row['Tên đơn vị']
                ki_hieu = row['Ký hiệu đơn vị']
                c.execute("INSERT OR IGNORE INTO units (dai_luong, ten_don_vi, ki_hieu_don_vi) VALUES (?, ?, ?)",
                          (dai_luong, ten_don_vi, ki_hieu))
            conn.commit()
            conn.close()
            messagebox.showinfo("Thành công", "Đã import dữ liệu từ Excel!")
            self.load_units()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi import: {str(e)}")