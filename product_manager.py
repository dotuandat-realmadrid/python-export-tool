# product_manager.py
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from database import DB_NAME
from utils import wrap_text
from add_product import AddProduct
from detail_product import DetailProduct
import openpyxl
from tkinter import filedialog
import re
from collections import defaultdict

class ProductManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Quản lý thiết bị")
        self.root.state('zoomed')
        
        # Tùy chỉnh style cho Treeview với chiều cao cố định đủ lớn
        self.style = ttk.Style()
        self.style.configure("Treeview", rowheight=60)

        # ===== FRAME CHÍNH =====
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ===== TREEVIEW =====
        columns = ("STT", "Dự án", "Danh sách thiết bị", "Chủng loại", "Đơn vị", "Số lượng", "Ghi chú", "Hành động")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings")
        
        # ===== ĐẶT TÊN CỘT =====
        for col in columns:
            self.tree.heading(col, text=col)
        
        # ===== CÀI KÍCH THƯỚC CỘT =====
        self.tree.column("STT", width=50, anchor="center")
        self.tree.column("Dự án", width=200, anchor="w")
        self.tree.column("Danh sách thiết bị", width=230, anchor="w")
        self.tree.column("Chủng loại", width=180, anchor="w")
        self.tree.column("Đơn vị", width=80, anchor="center")
        self.tree.column("Số lượng", width=80, anchor="center")
        self.tree.column("Ghi chú", width=150, anchor="w")
        self.tree.column("Hành động", width=100, anchor="center")

        # ===== SCROLLBAR =====
        scrollbar_y = tk.Scrollbar(main_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ===== BUTTONS =====
        button_frame = tk.Frame(root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Button(button_frame, text="Thêm mới", command=self.add_new_product, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Sửa", command=self.edit_product, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Import Excel", command=self.import_excel, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Thoát", command=self.root.destroy, width=12).pack(side=tk.RIGHT, padx=5)

        # ===== SỰ KIỆN =====
        self.tree.bind("<Double-1>", self.open_detail_product)
        self.tree.bind("<Button-1>", self.on_action_click)

        # ===== TẢI DỮ LIỆU =====
        self.load_products()

    def load_products(self):
        """Load danh sách sản phẩm - mỗi sản phẩm một dòng, hiển thị tất cả dự án trong ô 'Dự án'"""
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        
        # Query lấy thông tin sản phẩm và các dự án liên quan
        query = '''
            SELECT 
                p.id,
                p.name,
                p.note,
                p.unit,
                p.quantity,
                pt.name as product_type,
                prj.name as project_name,
                prj.ma_du_an as project_code
            FROM products p
            LEFT JOIN product_type_mapping_products ptmp ON p.id = ptmp.product_id
            LEFT JOIN product_types pt ON ptmp.type_id = pt.id
            LEFT JOIN product_projects pp ON p.id = pp.product_id
            LEFT JOIN projects prj ON pp.project_id = prj.id
            ORDER BY p.name
        '''
        
        products = c.execute(query).fetchall()
        conn.close()

        self.tree.delete(*self.tree.get_children())
        
        # Dictionary để nhóm các dự án theo product_id
        product_projects = {}
        for prod_id, product_name, note, unit, quantity, product_type, project_name, project_code in products:
            if prod_id not in product_projects:
                product_projects[prod_id] = {
                    "name": product_name,
                    "note": note,
                    "unit": unit if unit else "Bộ",
                    "quantity": quantity if quantity is not None else 1,
                    "product_type": product_type if product_type else "",
                    "projects": []
                }
            if project_name:
                display_project = project_name
                if project_code:
                    display_project += f" ({project_code})"
                product_projects[prod_id]["projects"].append(display_project)

        stt = 1
        for prod_id, data in product_projects.items():
            # Hiển thị tất cả dự án trong ô 'Dự án', mỗi dự án trên một dòng
            projects_str = "\n".join(data["projects"]) if data["projects"] else "Chưa có dự án"
            
            self.tree.insert("", "end", iid=str(prod_id), values=(
                stt,
                projects_str,
                wrap_text(data["name"], 40),
                data["product_type"],
                data["unit"],
                data["quantity"],
                wrap_text(data["note"] if data["note"] else "", 25),
                "Xóa"
            ))
            stt += 1

    def add_new_product(self):
        """Mở cửa sổ thêm mới"""
        add_window = tk.Toplevel(self.root)
        AddProduct(add_window, None, self)

    def edit_product(self):
        """Mở cửa sổ sửa thông tin sản phẩm"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một sản phẩm để sửa!")
            return
        
        prod_id = int(selected[0])
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        product = c.execute("SELECT name, note, unit, quantity FROM products WHERE id=?", (prod_id,)).fetchone()
        if not product:
            conn.close()
            return
        
        product_name, note, unit, quantity = product
        
        # Tạo cửa sổ sửa
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Sửa thông tin sản phẩm")
        edit_window.geometry("400x280")
        edit_window.transient(self.root)
        edit_window.grab_set()

        frame = tk.Frame(edit_window)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Hiển thị tên sản phẩm (không cho sửa)
        tk.Label(frame, text="Tên sản phẩm:").grid(row=0, column=0, sticky="w", pady=5)
        tk.Label(frame, text=product_name, state="disabled").grid(row=0, column=1, pady=5)

        # Trường Đơn vị (cho sửa)
        tk.Label(frame, text="Đơn vị:").grid(row=1, column=0, sticky="w", pady=5)
        unit_entry = tk.Entry(frame, width=30)
        unit_entry.insert(0, unit if unit else "Bộ")
        unit_entry.grid(row=1, column=1, pady=5)

        # Trường Số lượng (cho sửa)
        tk.Label(frame, text="Số lượng:").grid(row=2, column=0, sticky="w", pady=5)
        quantity_entry = tk.Entry(frame, width=30)
        quantity_entry.insert(0, str(quantity) if quantity is not None else "1")
        quantity_entry.grid(row=2, column=1, pady=5)

        # Trường Ghi chú (cho sửa)
        tk.Label(frame, text="Ghi chú:").grid(row=3, column=0, sticky="w", pady=5)
        note_entry = tk.Text(frame, width=30, height=5)
        note_entry.insert("1.0", note if note else "")
        note_entry.grid(row=3, column=1, pady=5)

        def save_changes():
            new_unit = unit_entry.get().strip()
            new_quantity = quantity_entry.get().strip()
            new_note = note_entry.get("1.0", tk.END).strip()
            
            if not new_unit or not new_quantity:
                messagebox.showerror("Lỗi", "Đơn vị và Số lượng không được để trống!")
                return
            
            try:
                new_quantity = int(new_quantity)  # Kiểm tra số lượng là số nguyên
                if new_quantity <= 0:
                    messagebox.showerror("Lỗi", "Số lượng phải lớn hơn 0!")
                    return
            except ValueError:
                messagebox.showerror("Lỗi", "Số lượng phải là một số nguyên!")
                return
            
            try:
                c.execute("UPDATE products SET note=?, unit=?, quantity=? WHERE id=?", 
                          (new_note, new_unit, new_quantity, prod_id))
                conn.commit()
                messagebox.showinfo("Thành công", "Đã cập nhật thông tin sản phẩm!")
                edit_window.destroy()
                self.load_products()
            except sqlite3.Error as e:
                conn.rollback()
                messagebox.showerror("Lỗi", f"Lỗi khi cập nhật: {str(e)}")
            finally:
                conn.close()

        tk.Button(frame, text="Lưu", command=save_changes).grid(row=4, column=1, pady=10)
        tk.Button(frame, text="Hủy", command=edit_window.destroy).grid(row=4, column=0, pady=10)

    def open_detail_product(self, event):
        """Double-click để xem chi tiết"""
        column = self.tree.identify_column(event.x)
        if column == "#8":  # Cột Hành động
            return
            
        selected = self.tree.selection()
        if not selected:
            return
        
        prod_id = int(selected[0])
        
        detail_win = tk.Toplevel(self.root)
        DetailProduct(detail_win, prod_id, self)

    def on_action_click(self, event):
        """Click cột Hành động để xóa"""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        column = self.tree.identify_column(event.x)
        if column == "#8":  # cột "Hành động"
            item = self.tree.identify_row(event.y)
            if item:
                prod_id = int(item)
                self.delete_product(prod_id)

    def delete_product(self, prod_id):
        """Xóa sản phẩm chính"""
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        product = c.execute("SELECT name FROM products WHERE id=?", (prod_id,)).fetchone()
        if not product:
            conn.close()
            return

        if messagebox.askyesno("Xác nhận", f"Xóa sản phẩm '{product[0]}'?\n\nLưu ý: Sản phẩm sẽ bị xóa khỏi tất cả dự án!"):
            try:
                c.execute("DELETE FROM product_type_mapping_products WHERE product_id=?", (prod_id,))
                c.execute("DELETE FROM product_projects WHERE product_id=?", (prod_id,))
                c.execute("DELETE FROM product_hidden_indicators WHERE product_id=?", (prod_id,))
                c.execute("DELETE FROM product_custom_indicators WHERE product_id=?", (prod_id,))
                c.execute("DELETE FROM reference_products WHERE product_id=?", (prod_id,))
                c.execute("DELETE FROM products WHERE id=?", (prod_id,))
                conn.commit()
                messagebox.showinfo("Thành công", "Đã xóa sản phẩm khỏi tất cả dự án")
                self.load_products()
            except Exception as e:
                conn.rollback()
                messagebox.showerror("Lỗi", f"Lỗi khi xóa: {str(e)}")
        conn.close()

    def refresh_data(self):
        """Refresh sau khi thêm"""
        self.load_products()

    def import_excel(self):
        """
        ĐÃ SỬA: Xử lý "Chủng loại" thành một indicator
        - Hàng "Chủng loại" VỪA lưu type_name và mapping, VỪA tạo indicator
        - Indicator "Chủng loại" có: requirement="Chủng loại", indicator=NULL, value=type_name
        - THÊM: Kiểm tra cột TT để phân biệt yêu cầu kỹ thuật cha và chỉ tiêu con

        ĐÃ SỬA THEO YÊU CẦU MỚI (18/11/2025):
        - Parse cột "Chỉ tiêu kỹ thuật hãng" theo format: "Mã: [tên SP] Hãng: [tên hãng]"
        - Lưu ref_value_{man_id}_product_name = tên sản phẩm tham khảo (cho hàng đầu)
        - Lưu ref_value_{man_id}_{ind_id} = nội dung cột tham chiếu (cho từng indicator)
        """
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active
            rows = list(sheet.iter_rows(values_only=True))

            if not rows or len(rows[0]) < 6:
                messagebox.showerror("Lỗi", "File Excel không đúng định dạng!")
                return

            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()

            conn.execute("BEGIN TRANSACTION")

            current_product_id = None
            current_type_id = None
            manufacturer_ids = []
            indicator_ids = []

            product_specs = defaultdict(list)
            reference_values = defaultdict(list)
            custom_indicators = {
                "three_brands": {},
                "bom": {},
                "dmkt": {},
                "ctkt_bo": {},
                "ctkt_mua_sam": {}
            }

            for row_idx, row in enumerate(rows[1:], start=2):
                tt = str(row[0]).strip() if row[0] else ""
                noi_dung = str(row[1]).strip() if row[1] else ""
                chi_tieu_chung_loai = str(row[2]).strip() if len(row) > 2 and row[2] else ""
                chi_tieu_ky_thuat = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                tieu_chi_danh_gia = str(row[4]) if len(row) > 4 and row[4] and str(row[4]).strip() else ""
                loai_chi_tieu = str(row[5]).strip() if len(row) > 5 and row[5] else "CTCB"
                
                # Lấy type_name từ cột M (index 12)
                type_name = str(row[12]).strip() if len(row) > 12 and row[12] else ""
                
                # Đọc cặp chi_tieu_hang và tham_chieu: G(6)-H(7), I(8)-J(9), K(10)-L(11)
                chi_tieu_hangs = []
                tham_chieus = []
                for i in range(6, 12, 2):  # 6(G),8(I),10(K)
                    if i < len(row):
                        chi_tieu_hang = str(row[i]).strip() if row[i] else ""
                        tham_chieu = str(row[i+1]).strip() if i+1 < len(row) and row[i+1] else ""
                        chi_tieu_hangs.append(chi_tieu_hang)
                        tham_chieus.append(tham_chieu)

                # Xử lý hàng số La Mã (sản phẩm)
                if re.match(r'^[IVXLCDM]+$', tt):
                    if current_product_id:
                        self.save_product_data(conn, c, current_product_id, current_type_id, manufacturer_ids, product_specs, reference_values, custom_indicators, indicator_ids)

                    product_name = noi_dung
                    c.execute("INSERT INTO products (name, ma_san_pham, note) VALUES (?, ?, ?)", (product_name, None, None))
                    current_product_id = c.lastrowid
                    manufacturer_ids = []
                    custom_indicators = {
                        "three_brands": {},
                        "bom": {},
                        "dmkt": {},
                        "ctkt_bo": {},
                        "ctkt_mua_sam": {}
                    }
                    indicator_ids = []

                    # Lưu type từ cột M nếu có
                    if type_name:
                        existing_type = c.execute("SELECT id FROM product_types WHERE name=?", (type_name,)).fetchone()
                        if existing_type:
                            current_type_id = existing_type[0]
                        else:
                            c.execute("INSERT INTO product_types (name) VALUES (?)", (type_name,))
                            current_type_id = c.lastrowid

                        # Kiểm tra trùng tên sản phẩm trong cùng loại
                        if current_product_id:
                            current_product_name = c.execute("SELECT name FROM products WHERE id=?", (current_product_id,)).fetchone()[0]
                            duplicate_check = c.execute("""
                                SELECT p.id, p.name 
                                FROM products p
                                JOIN product_type_mapping_products ptmp ON p.id = ptmp.product_id
                                WHERE ptmp.type_id = ? AND p.name = ? AND p.id != ?
                            """, (current_type_id, current_product_name, current_product_id)).fetchall()
                            
                            if duplicate_check:
                                conn.rollback()
                                messagebox.showerror("Lỗi Import", 
                                    f"Dòng {row_idx}: Sản phẩm '{current_product_name}' đã tồn tại trong loại '{type_name}'!\n"
                                    f"Không thể import 2 sản phẩm trùng tên trong cùng loại sản phẩm.")
                                conn.close()
                                return
                            
                            c.execute("INSERT INTO product_type_mapping_products (product_id, type_id) VALUES (?, ?)", (current_product_id, current_type_id))

                    # Xử lý thông tin hãng từ hàng sản phẩm
                    for idx, (chi_tieu_hang, tham_chieu) in enumerate(zip(chi_tieu_hangs, tham_chieus), 1):
                        # Bỏ qua nếu không có dữ liệu hãng
                        if not chi_tieu_hang:
                            continue
                            
                        chi_tieu_hang_str = str(chi_tieu_hang).strip()
                        tham_chieu_str = str(tham_chieu).strip() if tham_chieu else ""

                        # Parse format "Mã: ... Hãng: ..."
                        match = re.search(r'Mã\s*:\s*(.*?)\s*Hãng\s*:\s*(.*)', chi_tieu_hang_str, re.IGNORECASE)
                        
                        if match:
                            # Format mới: "Mã: Pro Rugged 14 Laptop Hãng: Dell"
                            product_name_ref = match.group(1).strip()  # "Pro Rugged 14 Laptop" (tên SP hiển thị ở cột chính)
                            manufacturer_name = match.group(2).strip()  # "Dell" (tên hãng)
                        else:
                            # Tương thích ngược: toàn bộ nội dung là tên hãng
                            manufacturer_name = chi_tieu_hang_str
                            product_name_ref = chi_tieu_hang_str

                        # Lưu vào bảng manufacturers
                        c.execute("INSERT INTO manufacturers (name, product_name) VALUES (?, ?)", (manufacturer_name, product_name_ref))
                        man_id = c.lastrowid
                        manufacturer_ids.append(man_id)

                        # SỬA: Lưu THAM CHIẾU (tham_chieu_str từ cột lẻ H2, J2, L2...) vào ref_value_{man_id}_product_name
                        # để hiển thị đúng ở cột “Tham chiếu …” của hàng “Tên sản phẩm tham khảo”
                        # Chỉ lấy nội dung cột tham chiếu, nếu trống → để rỗng luôn
                        ref_product_name = tham_chieu_str.strip()   # Không có fallback nữa

                        custom_indicators["three_brands"][f"ref_value_{man_id}_product_name"] = ref_product_name
                        custom_indicators["bom"][f"ref_value_{man_id}_product_name"] = ref_product_name

                        # Lưu vào product_specs và reference_values
                        product_specs[man_id].append(chi_tieu_hang_str)
                        reference_values[man_id].append(tham_chieu_str)

                        # Lưu mapping cho manufacturer với type_id nếu có
                        if current_type_id:
                            c.execute("INSERT INTO product_type_mapping (manufacturer_id, type_id) VALUES (?, ?)", (man_id, current_type_id))

                    continue

                # Xử lý các indicator khác
                has_indicator_data = chi_tieu_ky_thuat or tieu_chi_danh_gia or loai_chi_tieu != "CTCB" or any(chi_tieu_hangs) or any(tham_chieus) or (tt and noi_dung)

                if not has_indicator_data:
                    continue

                # Tạo indicator
                danh_gia, gia_tri, don_vi = self.parse_chi_tieu_ky_thuat(chi_tieu_ky_thuat)
                c.execute("INSERT INTO indicators (type_id, indicator_code, indicator, value, unit) VALUES (?, ?, ?, ?, ?)",
                        (current_type_id, tt, noi_dung, gia_tri, don_vi))
                ind_id = c.lastrowid
                indicator_ids.append(ind_id)

                # Lưu custom indicators
                if danh_gia:
                    custom_indicators["three_brands"][f"danh_gia_{ind_id}"] = danh_gia
                if gia_tri:
                    custom_indicators["three_brands"][f"so_sanh_{ind_id}"] = gia_tri
                    custom_indicators["bom"][f"so_sanh_{ind_id}"] = gia_tri
                    custom_indicators["dmkt"][f"so_sanh_{ind_id}"] = gia_tri
                    custom_indicators["ctkt_bo"][f"gia_tri_{ind_id}"] = gia_tri
                    custom_indicators["ctkt_mua_sam"][f"so_sanh_{ind_id}"] = gia_tri
                if don_vi:
                    custom_indicators["ctkt_mua_sam"][f"don_vi_{ind_id}"] = don_vi
                if tieu_chi_danh_gia:
                    formatted_tieu_chi = self.format_tieu_chi_danh_gia(tieu_chi_danh_gia)
                    custom_indicators["ctkt_mua_sam"][f"tieu_chi_{ind_id}"] = formatted_tieu_chi
                if loai_chi_tieu:
                    custom_indicators["three_brands"][f"crit_type_{ind_id}"] = loai_chi_tieu
                    custom_indicators["bom"][f"crit_type_{ind_id}"] = loai_chi_tieu
                    custom_indicators["dmkt"][f"crit_type_{ind_id}"] = loai_chi_tieu
                    custom_indicators["ctkt_mua_sam"][f"crit_type_{ind_id}"] = loai_chi_tieu

                # Xử lý giá trị hãng cho indicator này
                for man_idx, (chi_tieu_hang, tham_chieu) in enumerate(zip(chi_tieu_hangs, tham_chieus)):
                    if man_idx < len(manufacturer_ids):
                        man_id = manufacturer_ids[man_idx]
                        
                        # Lưu specification_value
                        if chi_tieu_hang:
                            product_specs[man_id].append(chi_tieu_hang)
                            c.execute("INSERT INTO product_specifications (manufacturer_id, indicator_id, specification_value) VALUES (?, ?, ?)",
                                    (man_id, ind_id, chi_tieu_hang))
                        
                        # Lưu tham chiếu với key ref_value_{man_id}_{ind_id}
                        if tham_chieu:
                            reference_values[man_id].append(tham_chieu)
                            custom_indicators["three_brands"][f"ref_value_{man_id}_{ind_id}"] = tham_chieu
                            custom_indicators["bom"][f"ref_value_{man_id}_{ind_id}"] = tham_chieu

            # Lưu sản phẩm cuối cùng
            if current_product_id:
                self.save_product_data(conn, c, current_product_id, current_type_id, manufacturer_ids, product_specs, reference_values, custom_indicators, indicator_ids)

            conn.commit()
            messagebox.showinfo("Thành công", "Đã import dữ liệu từ file Excel thành công!")
            self.load_products()

        except Exception as e:
            conn.rollback()
            messagebox.showerror("Lỗi", f"Lỗi khi import file: {str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            conn.close()

    def format_tieu_chi_danh_gia(self, tieu_chi_raw):
        if not tieu_chi_raw or str(tieu_chi_raw).strip() == "" or str(tieu_chi_raw) == "None":
            return ""
        
        content = str(tieu_chi_raw)
        content = content.replace('\r\n', '\n').replace('\r', '\n')
        content = content.strip()
        
        if "- Đạt:" not in content and "- Không đạt:" not in content:
            print(f"CẢNH BÁO: Tiêu chí không có '- Đạt:' hoặc '- Không đạt:': {content}")
            return ""
        
        # SỬA: In đầy đủ nội dung log
        print(f"Lưu tiêu chí ({len(content)} ký tự): {content}")
        
        return content

    def save_product_data(self, conn, c, product_id, type_id, manufacturer_ids, product_specs, reference_values, custom_indicators, indicator_ids):
        """
        GIẢI THÍCH: Lưu dữ liệu sản phẩm vào database
        - Lưu reference_products
        - Lưu product_custom_indicators (các custom đã được chuẩn bị sẵn trong custom_indicators)
        - ĐÃ SỬA: Bỏ phần lưu value từ indicators vào product_custom_indicators vì đã được lưu sẵn trong custom_indicators
        """
        # Lưu reference_products
        for sort_order, man_id in enumerate(manufacturer_ids):
            c.execute("INSERT INTO reference_products (product_id, manufacturer_id, sort_order) VALUES (?, ?, ?)",
                    (product_id, man_id, sort_order))

        # Lưu product_custom_indicators (các custom đã được chuẩn bị sẵn trong custom_indicators)
        for tab_name, customs in custom_indicators.items():
            for key, value in customs.items():
                if key and value:
                    try:
                        parts = key.split('_')
                        if len(parts) >= 2:
                            prefix = parts[0]
                            ind_id_str = '_'.join(parts[1:])
                            if ind_id_str.lstrip('-').isdigit():
                                ind_id_save = int(ind_id_str)
                            else:
                                ind_id_save = ind_id_str
                            custom_value = f"{prefix}_{value}"
                            unique_indicator_id = f"{ind_id_save}_{prefix}"

                            c.execute("""
                                INSERT INTO product_custom_indicators (product_id, tab_name, indicator_id, custom_value)
                                VALUES (?, ?, ?, ?)
                            """, (product_id, tab_name, unique_indicator_id, custom_value))
                    except Exception as e:
                        print(f"Lỗi khi lưu {key}: {str(e)}")
                        continue

    def parse_chi_tieu_ky_thuat(self, chi_tieu_ky_thuat):
        """
        GIẢI THÍCH: Phân tích cột Chỉ tiêu kỹ thuật để tách Đánh giá, Giá trị, Đơn vị
        - Đánh giá có 6 giá trị: >, <, =, <=, >=, not
        - Nếu không bắt đầu bằng ≥, ≤, <, >, =, lưu toàn bộ vào Giá trị, Đánh giá là 'not', Đơn vị rỗng
        - Nếu bắt đầu bằng các dấu trên:
          - Trường hợp 1: Dấu + số nguyên hoặc số nguyên x số nguyên → Đánh giá là dấu, Giá trị là số, Đơn vị rỗng
          - Trường hợp 2: Dấu + số nguyên + chữ → Đánh giá là dấu, Giá trị là số, Đơn vị là chữ
          - Trường hợp 3: Dấu + chữ → Đánh giá là dấu, Giá trị là chữ, Đơn vị rỗng
        """
        chi_tieu_ky_thuat = chi_tieu_ky_thuat.strip()
        if not chi_tieu_ky_thuat:
            return "not", "", ""

        # Các dấu đánh giá hợp lệ
        operators = {
            "≥": ">=", "≤": "<=", ">": ">", "<": "<", "=": "="
        }

        for op, danh_gia in operators.items():
            if chi_tieu_ky_thuat.startswith(op):
                content = chi_tieu_ky_thuat[len(op):].strip()
                # Trường hợp 1: Số nguyên hoặc số nguyên x số nguyên
                if re.match(r'^\d+(\s*x\s*\d+)?$', content):
                    return danh_gia, content, ""
                # Trường hợp 2: Số nguyên + chữ
                match = re.match(r'^(\d+)\s*(\S+)$', content)
                if match:
                    return danh_gia, match.group(1), match.group(2)
                # Trường hợp 3: Chỉ có chữ
                if content:
                    return danh_gia, content, ""
                return "not", chi_tieu_ky_thuat, ""
        
        # Trường hợp không có dấu
        return "not", chi_tieu_ky_thuat, ""