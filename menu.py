import tkinter as tk
from product_manager import ProductManager
from indicator_manager import IndicatorManager
from project_manager import ProjectManager
from unit_manager import UnitManager

class MainMenu:
    def __init__(self, root):
        self.root = root
        self.root.title("Quản lý hệ thống")
        self.root.geometry("540x200")

        # Cấu hình khoảng cách giữa các cột và hàng
        for i in range(2):
            self.root.columnconfigure(i, weight=1)
        for j in range(2):
            self.root.rowconfigure(j, weight=1)

        # Hàng 1
        tk.Button(root, text="Quản lý đơn vị", command=self.manage_units, width=20, height=2).grid(row=0, column=0, padx=20, pady=20)
        tk.Button(root, text="Quản lý khung chỉ tiêu cơ bản", command=self.manage_indicators, width=25, height=2).grid(row=0, column=1, padx=20, pady=20)

        # Hàng 2
        tk.Button(root, text="Quản lý sản phẩm", command=self.manage_products, width=20, height=2).grid(row=1, column=0, padx=20, pady=20)
        tk.Button(root, text="Quản lý dự án", command=self.manage_projects, width=20, height=2).grid(row=1, column=1, padx=20, pady=20)

    def manage_units(self):
        unit_win = tk.Toplevel(self.root)
        UnitManager(unit_win)

    def manage_indicators(self):
        indicator_win = tk.Toplevel(self.root)
        IndicatorManager(indicator_win)

    def manage_products(self):
        product_win = tk.Toplevel(self.root)
        ProductManager(product_win)

    def manage_projects(self):
        project_win = tk.Toplevel(self.root)
        ProjectManager(project_win)