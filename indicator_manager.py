import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from database import DB_NAME
from utils import wrap_text
from add_indicator import AddIndicatorWindow

class IndicatorManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Quản lý khung chỉ tiêu cơ bản")
        self.root.state('zoomed')
        
        # Biến để lưu trạng thái hiện tại
        self.current_screen = "types"  # "types" hoặc "indicators"
        self.current_type_id = None
        self.current_type_name = None
        
        # Tạo main frame
        self.main_frame = tk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Cấu hình grid cho main_frame để frame con mở rộng
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # Tạo các frame cho từng màn hình
        self.create_type_screen()
        self.create_indicator_screen()
        
        # Hiển thị màn hình loại sản phẩm đầu tiên
        self.show_type_screen()
    
    def create_type_screen(self):
        """Tạo giao diện màn hình danh sách loại sản phẩm"""
        # Frame cho màn hình loại sản phẩm
        self.type_frame = tk.Frame(self.main_frame)
        self.type_frame.grid_rowconfigure(1, weight=1)  # Cho tree_frame mở rộng
        self.type_frame.grid_columnconfigure(0, weight=1)
        
        # Title
        title_label = tk.Label(self.type_frame, text="Danh sách loại sản phẩm", 
                              font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 20), sticky="ew")
        
        # Treeview frame for types - CHỈ 2 CỘT
        tree_frame = tk.LabelFrame(self.type_frame, padx=5, pady=5)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        v_scrollbar = tk.Scrollbar(tree_frame, orient="vertical")
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Thêm style riêng cho type_tree
        type_style = ttk.Style()
        type_style.configure("TypeTree.Treeview", rowheight=100)  # Cao cố định
        
        self.type_tree = ttk.Treeview(tree_frame, columns=("STT", "Loại sản phẩm"), show="headings",
                                     yscrollcommand=v_scrollbar.set, style="TypeTree.Treeview")
        v_scrollbar.config(command=self.type_tree.yview)
        self.type_tree.heading("STT", text="STT")
        self.type_tree.heading("Loại sản phẩm", text="Loại sản phẩm")
        self.type_tree.column("STT", width=50, anchor="center")
        self.type_tree.column("Loại sản phẩm", width=600, anchor="center")
        self.type_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Bind double-click event for types
        self.type_tree.bind("<Double-1>", self.on_type_selected)
        
        # Buttons frame - THÊM NÚT "Thêm mới=", "Sửa", "Xóa"
        button_frame = tk.Frame(self.type_frame)
        button_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        
        # Thêm nút "Thêm mới" ở bên trái
        tk.Button(button_frame, text="Thêm mới", command=self.add_new_indicator).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Sửa", command=self.rename_product_type).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Xóa", command=self.delete_product_type).pack(side=tk.LEFT, padx=5)
        # Nút "Thoát" ở bên phải
        tk.Button(button_frame, text="Thoát", command=self.root.destroy).pack(side=tk.RIGHT, padx=5)
        
        # Load types initially
        self.load_types()
    
    def create_indicator_screen(self):
        """Tạo giao diện màn hình danh sách chỉ tiêu"""
        # Frame cho màn hình chỉ tiêu
        self.indicator_frame = tk.Frame(self.main_frame)
        self.indicator_frame.grid_rowconfigure(1, weight=1)  # Cho tree_frame mở rộng
        self.indicator_frame.grid_columnconfigure(0, weight=1)
        
        # Title with back button
        title_frame = tk.Frame(self.indicator_frame)
        title_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        self.back_button = tk.Button(title_frame, text="◀ Quay lại danh sách loại", 
                                    command=self.show_type_screen, bg="lightgray", 
                                    font=("Arial", 10))
        self.back_button.pack(side=tk.LEFT)
        
        self.title_label = tk.Label(title_frame, text="Chỉ tiêu của loại: [Tên loại sản phẩm]", 
                                   font=("Arial", 16, "bold"))
        self.title_label.pack(side=tk.LEFT, padx=(20, 0))
        
        # Treeview frame for indicators
        tree_frame = tk.LabelFrame(self.indicator_frame, text="Danh sách chỉ tiêu", padx=5, pady=5)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        v_scrollbar = tk.Scrollbar(tree_frame, orient="vertical")
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        h_scrollbar = tk.Scrollbar(tree_frame, orient="horizontal")
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Style riêng cho indicator tree với rowheight cố định 100
        indicator_style = ttk.Style()
        indicator_style.configure("IndicatorTree.Treeview", rowheight=100)  # Cao cố định 100
        
        self.tree = ttk.Treeview(tree_frame, columns=("Mã chỉ tiêu", "Chỉ tiêu", "Giá trị", "Đơn vị", "Hành động"), 
                                 show="headings", yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set,
                                 style="IndicatorTree.Treeview")
        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)
        
        self.tree.heading("Mã chỉ tiêu", text="Mã chỉ tiêu")
        self.tree.heading("Chỉ tiêu", text="Chỉ tiêu")
        self.tree.heading("Giá trị", text="Giá trị")
        self.tree.heading("Đơn vị", text="Đơn vị")
        self.tree.heading("Hành động", text="Hành động")
        
        # Adjust column widths
        self.tree.column("Mã chỉ tiêu", width=100, anchor="center")
        self.tree.column("Chỉ tiêu", width=240, anchor="w")
        self.tree.column("Giá trị", width=460, anchor="center")
        self.tree.column("Đơn vị", width=150, anchor="center")
        self.tree.column("Hành động", width=100, anchor="center")
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Bind events
        self.tree.bind("<Double-1>", self.on_double_click_edit)
        self.tree.bind("<Button-1>", self.on_click)
        
        # Buttons frame
        button_frame = tk.Frame(self.indicator_frame)
        button_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        
        tk.Button(button_frame, text="+", command=self.add_new_row).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Lưu", command=self.save_new_indicators).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Quay lại", command=self.show_type_screen).pack(side=tk.RIGHT, padx=5)
    
    def show_type_screen(self):
        """Hiển thị màn hình loại sản phẩm"""
        self.indicator_frame.grid_remove()  # Ẩn indicator_frame mà không mất cấu hình
        self.type_frame.grid(row=0, column=0, sticky="nsew")  # Hiển thị type_frame đầy khung
        self.current_screen = "types"
        if hasattr(self, 'back_button'):
            self.back_button.pack_forget()  # Ẩn nút back nếu cần
        
        # Đảm bảo type_tree luôn có rowheight nhỏ
        style = ttk.Style()
        style.configure("TypeTree.Treeview", rowheight=100)
        
        self.load_types()
    
    def show_indicator_screen(self, type_id, type_name):
        """Hiển thị màn hình danh sách chỉ tiêu"""
        self.type_frame.grid_remove()  # Ẩn type_frame mà không mất cấu hình
        self.indicator_frame.grid(row=0, column=0, sticky="nsew")  # Hiển thị indicator_frame đầy khung
        
        self.current_type_id = type_id
        self.current_type_name = type_name
        self.title_label.config(text=f"Chỉ tiêu của loại: {type_name}")
        
        self.current_screen = "indicators"
        self.back_button.pack(side=tk.LEFT)
        self.load_indicators()
    
    def load_types(self):
        """Tải danh sách loại sản phẩm - lưu type_id trong iid"""
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        types = c.execute("SELECT id, name FROM product_types ORDER BY name").fetchall()
        
        self.type_tree.delete(*self.type_tree.get_children())
        for index, (type_id, name) in enumerate(types, 1):
            self.type_tree.insert("", "end", iid=str(type_id), values=(index, name))
        
        conn.close()
    
    def on_type_selected(self, event):
        """Xử lý khi chọn loại sản phẩm - chuyển sang màn hình chỉ tiêu và load ngay"""
        selected = self.type_tree.selection()
        if not selected:
            return
        
        item = self.type_tree.item(selected[0])
        values = item['values']
        type_id = selected[0]
        type_name = values[1]
        
        if type_id and type_name:
            self.show_indicator_screen(type_id, type_name)
        else:
            messagebox.showerror("Lỗi", "Không thể xác định loại sản phẩm")
    
    def rename_product_type(self):
        selected = self.type_tree.selection()
        if not selected:
            messagebox.showerror("Lỗi", "Vui lòng chọn loại sản phẩm để sửa")
            return

        item = self.type_tree.item(selected[0])
        old_name = item['values'][1]

        popup = tk.Toplevel(self.root)
        popup.title("Sửa tên loại sản phẩm")

        text_box = tk.Text(popup, width=40, height=4, font=("Arial", 14))
        text_box.insert("1.0", old_name)
        text_box.grid(row=0, column=0, padx=10, pady=20, sticky="ew")

        popup.grid_columnconfigure(1, weight=1)

        def save_new_name():
            new_name = text_box.get("1.0", "end-1c").strip()
            if not new_name:
                messagebox.showerror("Lỗi", "Tên loại sản phẩm mới không được để trống")
                return
            if new_name == old_name:
                messagebox.showinfo("Thông báo", "Tên loại sản phẩm không thay đổi")
                popup.destroy()
                return

            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            try:
                existing = c.execute(
                    "SELECT id FROM product_types WHERE name=?", (new_name,)
                ).fetchone()
                if existing:
                    messagebox.showerror("Lỗi", "Tên loại sản phẩm mới đã tồn tại!")
                    return

                c.execute("UPDATE product_types SET name=? WHERE name=?", (new_name, old_name))
                conn.commit()
                messagebox.showinfo("Thành công", f"Đã sửa '{old_name}' thành '{new_name}'!")
                self.load_types()
                popup.destroy()
            except Exception as e:
                conn.rollback()
                messagebox.showerror("Lỗi", f"Lỗi khi sửa: {str(e)}")
            finally:
                conn.close()

        btn_frame = tk.Frame(popup)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=20)

        tk.Button(btn_frame, text="Lưu", width=10, command=save_new_name).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Hủy", width=10, command=popup.destroy).pack(side="left", padx=10)

    def delete_product_type(self):
        """Xóa loại sản phẩm được chọn, các chỉ tiêu liên quan, và các nhà sản xuất liên quan"""
        selected = self.type_tree.selection()
        if not selected:
            messagebox.showerror("Lỗi", "Vui lòng chọn loại sản phẩm để xóa")
            return
        
        item = self.type_tree.item(selected[0])
        type_name = item['values'][1]
        type_id = selected[0]
        
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        
        # Kiểm tra ràng buộc trong product_type_mapping_products
        mapping_count = c.execute("SELECT COUNT(*) FROM product_type_mapping_products WHERE type_id=?", (type_id,)).fetchone()[0]
        if mapping_count > 0:
            messagebox.showerror("Lỗi", "Không thể xóa loại sản phẩm đang được sử dụng trong ánh xạ với sản phẩm!")
            conn.close()
            return
        
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa loại sản phẩm '{type_name}', tất cả chỉ tiêu liên quan, và các nhà sản xuất liên quan?"):
            try:
                # Lấy danh sách manufacturer_id liên quan đến type_id
                manufacturer_ids = c.execute("SELECT manufacturer_id FROM product_type_mapping WHERE type_id=?", (type_id,)).fetchall()
                
                # Xóa các bản ghi liên quan trong product_specifications
                indicator_ids = c.execute("SELECT id FROM indicators WHERE type_id=?", (type_id,)).fetchall()
                for ind_id in indicator_ids:
                    c.execute("DELETE FROM product_specifications WHERE indicator_id=?", (ind_id[0],))
                
                # Xóa các chỉ tiêu liên quan
                c.execute("DELETE FROM indicators WHERE type_id=?", (type_id,))
                
                # Xóa các bản ghi trong product_type_mapping
                c.execute("DELETE FROM product_type_mapping WHERE type_id=?", (type_id,))
                
                # Xóa các nhà sản xuất liên quan
                for man_id in manufacturer_ids:
                    c.execute("DELETE FROM product_specifications WHERE manufacturer_id=?", (man_id[0],))
                    c.execute("DELETE FROM reference_products WHERE manufacturer_id=?", (man_id[0],))
                    c.execute("DELETE FROM manufacturers WHERE id=?", (man_id[0],))
                
                # Xóa loại sản phẩm
                c.execute("DELETE FROM product_types WHERE id=?", (type_id,))
                
                conn.commit()
                messagebox.showinfo("Thành công", "Đã xóa loại sản phẩm, các chỉ tiêu, và các nhà sản xuất liên quan!")
                self.load_types()
            except Exception as e:
                conn.rollback()
                messagebox.showerror("Lỗi", f"Lỗi khi xóa: {str(e)}")
            finally:
                conn.close()
    
    def add_new_indicator(self):
        """Mở cửa sổ thêm loại sản phẩm và chỉ tiêu mới"""
        add_win = tk.Toplevel(self.root)
        add_win.geometry("1080x600")
        AddIndicatorWindow(add_win, self.load_types)
    
    def on_click(self, event):
        """Xử lý click vào cột Hành động"""
        column = self.tree.identify_column(event.x)
        if column == "#5":
            item = self.tree.identify_row(event.y)
            if item:
                if item.startswith("new_"):
                    self.tree.delete(item)
                else:
                    self.delete_indicator(item)
    
    def delete_indicator(self, item):
        """Xóa chỉ tiêu từ DB"""
        ind_id = int(item)
        if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa chỉ tiêu này?"):
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            try:
                c.execute("DELETE FROM product_specifications WHERE indicator_id=?", (ind_id,))
                c.execute("DELETE FROM indicators WHERE id=?", (ind_id,))
                conn.commit()
                self.load_indicators()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi xóa: {str(e)}")
            finally:
                conn.close()
    
    def add_new_row(self):
        """Thêm hàng mới để thêm chỉ tiêu"""
        if not self.current_type_id:
            messagebox.showerror("Lỗi", "Vui lòng chọn loại sản phẩm trước khi thêm")
            return
        current_children = self.tree.get_children()
        new_iid = f"new_{len(current_children) + 1}"
        self.tree.insert("", "end", iid=new_iid, values=("", "", "", "", "Hủy"))
    
    def save_new_indicators(self):
        """Lưu các chỉ tiêu mới thêm"""
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        try:
            for child in list(self.tree.get_children()):
                if child.startswith("new_"):
                    values = self.tree.item(child)['values']
                    indicator_code_wrapped = values[0]
                    ind_wrapped = values[1]
                    value_wrapped = values[2]
                    unit_wrapped = values[3]
                    indicator_code = indicator_code_wrapped.replace('\n', ' ')
                    ind = ind_wrapped.replace('\n', ' ')
                    value = value_wrapped.replace('\n', ' ')
                    unit = unit_wrapped.replace('\n', ' ')
                    if not indicator_code:
                        messagebox.showerror("Lỗi", "Mã chỉ tiêu bắt buộc phải có")
                        return
                    if not ind:
                        messagebox.showerror("Lỗi", "Chỉ tiêu bắt buộc phải có")
                        return
                    type_id = self.current_type_id
                    c.execute('''INSERT INTO indicators (type_id, indicator_code, indicator, value, unit)
                                 VALUES (?, ?, ?, ?, ?)''', (type_id, indicator_code, ind, value, unit))
                    ind_id = c.lastrowid
                    man_ids = c.execute("SELECT manufacturer_id FROM product_type_mapping WHERE type_id=?", (type_id,)).fetchall()
                    for m in man_ids:
                        c.execute("INSERT OR IGNORE INTO product_specifications (manufacturer_id, indicator_id, specification_value) VALUES (?, ?, ?)", 
                                  (m[0], ind_id, ""))
                    self.tree.delete(child)
            conn.commit()
            messagebox.showinfo("Thành công", "Cập nhật thành công!")
            self.load_indicators()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("Lỗi", f"Lỗi khi lưu: {str(e)}")
        finally:
            conn.close()
    
    def on_double_click_edit(self, event):
        """Bắt đầu chỉnh sửa khi double-click vào ô"""
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if item and column not in ("#5"):
            self.start_inline_edit(item, column, event.x, event.y)
    
    def get_units(self):
        """Lấy danh sách ký hiệu đơn vị từ bảng units"""
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        units = c.execute("SELECT ki_hieu_don_vi FROM units ORDER BY ki_hieu_don_vi").fetchall()
        conn.close()
        return [u[0] for u in units]
    
    def start_inline_edit(self, item, column, x, y):
        """Bắt đầu inline editing"""
        column_index = int(column[1:]) - 1
        current_value = self.tree.set(item, column).replace('\n', ' ')
        
        bbox = self.tree.bbox(item, column)
        if not bbox:
            return
        entry_x, entry_y, entry_width, entry_height = bbox
        
        if column == "#4":  # Đơn vị - Combobox
            self.edit_combo = ttk.Combobox(self.tree, values=self.get_units())
            self.edit_combo.set(current_value)
            self.edit_combo.place(x=entry_x, y=entry_y, width=entry_width, height=entry_height, anchor="nw")
            self.edit_combo.focus_set()
            self.edit_combo.bind("<Return>", lambda e: self.save_inline_edit(item, column, column_index))
            self.edit_combo.bind("<FocusOut>", lambda e: self.save_inline_edit(item, column, column_index))
            self.edit_combo.bind("<Escape>", lambda e: self.cancel_inline_edit())
        elif column in ("#1", "#2", "#3"):  # Các cột multiline: Mã chỉ tiêu, Chỉ tiêu, Giá trị - Text với scrollbar
            self.edit_text = tk.Text(self.tree, wrap='word')
            self.edit_text.insert("1.0", current_value)
            text_width = entry_width - 20  # Để chỗ cho scrollbar
            self.edit_text.place(x=entry_x, y=entry_y, width=text_width, height=entry_height, anchor="nw")
            
            self.edit_scrollbar = tk.Scrollbar(self.tree, orient="vertical", command=self.edit_text.yview)
            self.edit_scrollbar.place(x=entry_x + text_width, y=entry_y, width=20, height=entry_height, anchor="nw")
            
            self.edit_text.config(yscrollcommand=self.edit_scrollbar.set)
            self.edit_text.focus_set()
            self.edit_text.bind("<Control-Return>", lambda e: self.save_inline_edit(item, column, column_index))
            self.edit_text.bind("<FocusOut>", lambda e: self.save_inline_edit(item, column, column_index))
            self.edit_text.bind("<Escape>", lambda e: self.cancel_inline_edit())
        else:  # Các cột khác - Entry
            self.edit_entry = tk.Entry(self.tree)
            self.edit_entry.insert(0, current_value)
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
        if hasattr(self, 'edit_combo') and self.edit_combo:
            new_value = self.edit_combo.get().strip()
            self.edit_combo.destroy()
            self.edit_combo = None
        elif hasattr(self, 'edit_text') and self.edit_text:
            new_value = self.edit_text.get("1.0", "end-1c").strip()
            self.edit_text.destroy()
            if hasattr(self, 'edit_scrollbar') and self.edit_scrollbar:
                self.edit_scrollbar.destroy()
                self.edit_scrollbar = None
            self.edit_text = None
        elif hasattr(self, 'edit_entry') and self.edit_entry:
            new_value = self.edit_entry.get().strip()
            self.edit_entry.destroy()
            self.edit_entry = None
        else:
            return
        
        if column_index == 0:
            wrapped_value = wrap_text(new_value, 40)
        elif column_index == 1:
            wrapped_value = wrap_text(new_value, 40)
        elif column_index == 2:
            wrapped_value = wrap_text(new_value, 80)
        elif column_index == 3:
            wrapped_value = wrap_text(new_value, 20)
        else:
            wrapped_value = new_value
        
        self.tree.set(item, column, wrapped_value)
        original_value = new_value
        
        if not item.startswith("new_"):
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            try:
                ind_id = int(item)
                if column_index == 0:
                    c.execute("UPDATE indicators SET indicator_code=? WHERE id=?", (original_value, ind_id))
                elif column_index == 1:
                    c.execute("UPDATE indicators SET indicator=? WHERE id=?", (original_value, ind_id))
                elif column_index == 2:
                    c.execute("UPDATE indicators SET value=? WHERE id=?", (original_value, ind_id))
                elif column_index == 3:
                    c.execute("UPDATE indicators SET unit=? WHERE id=?", (original_value, ind_id))
                conn.commit()
                self.load_indicators()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi cập nhật: {str(e)}")
            finally:
                conn.close()
        
        self.current_edit_item = None
        self.current_edit_column = None
        self.current_edit_column_index = None
    
    def cancel_inline_edit(self):
        """Hủy chỉnh sửa inline"""
        if hasattr(self, 'edit_entry') and self.edit_entry:
            self.edit_entry.destroy()
            self.edit_entry = None
        if hasattr(self, 'edit_combo') and self.edit_combo:
            self.edit_combo.destroy()
            self.edit_combo = None
        if hasattr(self, 'edit_text') and self.edit_text:
            self.edit_text.destroy()
            self.edit_text = None
        if hasattr(self, 'edit_scrollbar') and self.edit_scrollbar:
            self.edit_scrollbar.destroy()
            self.edit_scrollbar = None
        self.current_edit_item = None
        self.current_edit_column = None
        self.current_edit_column_index = None
    
    def load_indicators(self):
        """Tải chỉ tiêu theo loại sản phẩm đã chọn"""
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        
        if self.current_type_id:
            indicators = c.execute('''SELECT id, indicator_code, indicator, value, unit 
                                    FROM indicators 
                                    WHERE type_id = ?''', (self.current_type_id,)).fetchall()
        else:
            indicators = []
        
        self.tree.delete(*self.tree.get_children())
        for ind in indicators:
            indicator_code_wrapped = wrap_text(ind[1], 40)
            ind_wrapped = wrap_text(ind[2], 40)
            value_wrapped = wrap_text(ind[3] or "", 80)
            unit_wrapped = wrap_text(ind[4] or "", 20)
            self.tree.insert("", "end", iid=str(ind[0]), values=(indicator_code_wrapped, ind_wrapped, value_wrapped, unit_wrapped, "Xóa"))
        
        conn.close()