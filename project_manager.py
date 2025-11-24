import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from database import DB_NAME
from utils import wrap_text
from detail_project import DetailProject


class ProjectManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Quản lý dự án")
        self.root.geometry("900x600")

        # Tùy chỉnh style cho Treeview để đặt chiều cao hàng
        style = ttk.Style()
        style.configure("Treeview", rowheight=30)

        # Main frame
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Title
        title_label = tk.Label(main_frame, text="Danh sách dự án", font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 10))

        # Treeview frame
        tree_frame = ttk.LabelFrame(main_frame, text="Danh sách dự án", padding=5)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbar
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Treeview
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("STT", "Tên dự án", "Mã dự án", "Ghi chú", "Hành động"),
            show="headings",
            yscrollcommand=v_scrollbar.set
        )
        v_scrollbar.config(command=self.tree.yview)

        # Cấu hình cột
        self.tree.heading("STT", text="STT")
        self.tree.heading("Tên dự án", text="Tên dự án")
        self.tree.heading("Mã dự án", text="Mã dự án")
        self.tree.heading("Ghi chú", text="Ghi chú")
        self.tree.heading("Hành động", text="Hành động")

        self.tree.column("STT", width=50, anchor="center")
        self.tree.column("Tên dự án", width=200, anchor="center")
        self.tree.column("Mã dự án", width=100, anchor="center")
        self.tree.column("Ghi chú", width=250, anchor="w")
        self.tree.column("Hành động", width=100, anchor="center")

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Bind double-click
        self.tree.bind("<Double-1>", self.on_project_selected)
        self.tree.bind("<Button-1>", self.on_tree_click)

        # Button frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        tk.Button(button_frame, text="Thêm mới", command=self.add_new_project, font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Sửa", command=self.edit_project, font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Thoát", command=root.destroy, font=("Arial", 12)).pack(side=tk.RIGHT, padx=5)

        # Load dữ liệu
        self.load_projects()

    def load_projects(self):
        """Tải danh sách dự án từ DB."""
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        projects = c.execute('''SELECT id, name, ma_du_an, ghi_chu 
                                FROM projects 
                                ORDER BY name''').fetchall()
        self.tree.delete(*self.tree.get_children())

        for index, (proj_id, name, ma_du_an, ghi_chu) in enumerate(projects, 1):
            self.tree.insert(
                "",
                "end",
                iid=str(proj_id),
                values=(
                    index,
                    wrap_text(name, 30),
                    ma_du_an,
                    wrap_text(ghi_chu, 20),
                    "Xóa"
                ),
            )

        conn.close()

    def on_tree_click(self, event):
        """Bắt sự kiện click chuột vào cột hành động để xóa."""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        column = self.tree.identify_column(event.x)
        if column != "#5":
            return

        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return

        proj_id = int(row_id)
        self.delete_project(proj_id)

    def delete_project(self, proj_id):
        """Xóa dự án khỏi DB (có xác nhận)."""
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        project_name = c.execute("SELECT name FROM projects WHERE id=?", (proj_id,)).fetchone()
        if not project_name:
            messagebox.showerror("Lỗi", "Dự án không tồn tại!")
            conn.close()
            return
        project_name = project_name[0]

        confirm = messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa dự án '{project_name}' không?")
        if not confirm:
            conn.close()
            return

        try:
            # Bắt đầu transaction
            conn.execute("BEGIN")

            # Xóa mapping sản phẩm - dự án
            c.execute("DELETE FROM product_projects WHERE project_id=?", (proj_id,))
            
            # Xóa dự án
            c.execute("DELETE FROM projects WHERE id=?", (proj_id,))
            
            conn.commit()
            messagebox.showinfo("Thành công", f"Đã xóa dự án '{project_name}'!")
            self.load_projects()
        except sqlite3.IntegrityError as e:
            conn.rollback()
            messagebox.showerror("Lỗi", f"Không thể xóa dự án do ràng buộc dữ liệu.\nChi tiết: {e}")
        except sqlite3.Error as e:
            conn.rollback()
            messagebox.showerror("Lỗi", f"Không thể xóa dự án.\nChi tiết: {e}")
        finally:
            conn.close()

    def on_project_selected(self, event):
        """Double-click mở chi tiết dự án."""
        # Kiểm tra xem có click vào cột Hành động không
        column = self.tree.identify_column(event.x)
        if column == "#5":  # Cột Hành động
            return
        
        selected = self.tree.selection()
        if not selected:
            return
        proj_id = int(selected[0])
        detail_win = tk.Toplevel(self.root)
        DetailProject(detail_win, proj_id, self.refresh_list)

    def add_new_project(self):
        """Mở dialog thêm dự án mới."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Thêm dự án mới")
        dialog.geometry("400x280")
        dialog.transient(self.root)
        dialog.grab_set()

        frame = tk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(frame, text="Tên dự án:").grid(row=0, column=0, sticky="w", pady=5)
        name_entry = tk.Entry(frame, width=30)
        name_entry.grid(row=0, column=1, pady=5)

        tk.Label(frame, text="Mã dự án:").grid(row=1, column=0, sticky="w", pady=5)
        ma_entry = tk.Entry(frame, width=30)
        ma_entry.grid(row=1, column=1, pady=5)

        tk.Label(frame, text="Ghi chú:").grid(row=2, column=0, sticky="w", pady=5)
        ghi_entry = tk.Text(frame, width=30, height=8)
        ghi_entry.grid(row=2, column=1, pady=5)

        def save_project():
            name = name_entry.get().strip()
            ma = ma_entry.get().strip()
            ghi = ghi_entry.get("1.0", tk.END).strip()
            if not all([name, ma]):
                messagebox.showerror("Lỗi", "Điền đầy đủ Tên dự án và Mã dự án!")
                return
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            try:
                c.execute("INSERT INTO projects (name, ma_du_an, ghi_chu) VALUES (?, ?, ?)",
                          (name, ma, ghi))
                conn.commit()
                messagebox.showinfo("Thành công", "Đã thêm dự án!")
                dialog.destroy()
                self.load_projects()
            except sqlite3.IntegrityError:
                messagebox.showerror("Lỗi", "Mã dự án đã tồn tại!")
            finally:
                conn.close()

        tk.Button(frame, text="Thêm", command=save_project).grid(row=3, column=1, pady=10)
        tk.Button(frame, text="Hủy", command=dialog.destroy).grid(row=3, column=0, pady=10)

        name_entry.focus_set()

    def edit_project(self):
        """Mở dialog sửa dự án."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một dự án để sửa!")
            return
        
        proj_id = int(selected[0])
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        project = c.execute("SELECT name, ma_du_an, ghi_chu FROM projects WHERE id=?", (proj_id,)).fetchone()
        if not project:
            conn.close()
            return
        
        name, ma_du_an, ghi_chu = project
        
        # Tạo cửa sổ sửa
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Sửa dự án")
        edit_window.geometry("400x280")
        edit_window.transient(self.root)
        edit_window.grab_set()

        frame = tk.Frame(edit_window)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(frame, text="Tên dự án:").grid(row=0, column=0, sticky="w", pady=5)
        name_entry = tk.Entry(frame, width=30)
        name_entry.insert(0, name)
        name_entry.grid(row=0, column=1, pady=5)

        tk.Label(frame, text="Mã dự án:").grid(row=1, column=0, sticky="w", pady=5)
        ma_entry = tk.Entry(frame, width=30)
        ma_entry.insert(0, ma_du_an)
        ma_entry.grid(row=1, column=1, pady=5)

        tk.Label(frame, text="Ghi chú:").grid(row=2, column=0, sticky="w", pady=5)
        ghi_entry = tk.Text(frame, width=30, height=8)
        ghi_entry.insert("1.0", ghi_chu if ghi_chu else "")
        ghi_entry.grid(row=2, column=1, pady=5)

        def save_changes():
            new_name = name_entry.get().strip()
            new_ma = ma_entry.get().strip()
            new_ghi = ghi_entry.get("1.0", tk.END).strip()
            if not all([new_name, new_ma]):
                messagebox.showerror("Lỗi", "Điền đầy đủ Tên dự án và Mã dự án!")
                return
            
            try:
                c.execute("UPDATE projects SET name=?, ma_du_an=?, ghi_chu=? WHERE id=?", (new_name, new_ma, new_ghi, proj_id))
                conn.commit()
                messagebox.showinfo("Thành công", "Đã cập nhật dự án!")
                edit_window.destroy()
                self.load_projects()
            except sqlite3.IntegrityError:
                conn.rollback()
                messagebox.showerror("Lỗi", "Mã dự án đã tồn tại!")
            except sqlite3.Error as e:
                conn.rollback()
                messagebox.showerror("Lỗi", f"Lỗi khi cập nhật: {str(e)}")
            finally:
                conn.close()

        tk.Button(frame, text="Lưu", command=save_changes).grid(row=3, column=1, pady=10)
        tk.Button(frame, text="Hủy", command=edit_window.destroy).grid(row=3, column=0, pady=10)

        # Không đóng kết nối ngay, để save_changes xử lý

    def refresh_list(self):
        """Callback để refresh danh sách sau khi chỉnh sửa chi tiết."""
        self.load_projects()