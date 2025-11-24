# database.py
import sqlite3

DB_NAME = "product_db.sqlite"

def init_db():
    """
    Khởi tạo cơ sở dữ liệu với các bảng cần thiết.
    """
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    # Kiểm tra schema hiện tại của bảng products
    c.execute("PRAGMA table_info(products)")
    columns = [info[1] for info in c.fetchall()]
    
    # Nếu thiếu cột ma_san_pham, note, unit hoặc quantity, tạo bảng mới và sao chép dữ liệu
    if not all(col in columns for col in ["ma_san_pham", "note", "unit", "quantity"]):
        available_columns = [col for col in ["id", "name", "ma_san_pham", "note"] if col in columns]
        select_clause = ", ".join(available_columns)
        insert_clause = ", ".join(available_columns)
        
        c.execute('''CREATE TABLE products_new
                     (id INTEGER PRIMARY KEY AUTOINCREMENT,
                      name TEXT,
                      ma_san_pham TEXT UNIQUE,
                      note TEXT,
                      unit TEXT,
                      quantity INTEGER)''')
        
        if select_clause:
            c.execute(f"INSERT INTO products_new ({insert_clause}) SELECT {select_clause} FROM products")
        
        c.execute("DROP TABLE IF EXISTS products")
        c.execute("ALTER TABLE products_new RENAME TO products")
        conn.commit()

    # product_types: loại sản phẩm
    c.execute('''CREATE TABLE IF NOT EXISTS product_types
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE)''')

    # indicators: chỉ tiêu chung
    c.execute('''CREATE TABLE IF NOT EXISTS indicators
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  type_id INTEGER,
                  requirement TEXT,
                  indicator TEXT,
                  value TEXT,
                  unit TEXT,
                  FOREIGN KEY(type_id) REFERENCES product_types(id))''')

    # Kiểm tra schema hiện tại của bảng indicators và thêm cột value nếu thiếu
    c.execute("PRAGMA table_info(indicators)")
    columns = [info[1] for info in c.fetchall()]
    if "value" not in columns:
        c.execute("ALTER TABLE indicators ADD COLUMN value TEXT")
        conn.commit()

    # manufacturers: catalog sản phẩm tham khảo
    c.execute('''CREATE TABLE IF NOT EXISTS manufacturers
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  name TEXT,
                  product_name TEXT)''')

    # product_type_mapping: mapping manufacturers <-> product_types
    c.execute('''CREATE TABLE IF NOT EXISTS product_type_mapping
                 (manufacturer_id INTEGER,
                  type_id INTEGER,
                  PRIMARY KEY(manufacturer_id, type_id),
                  FOREIGN KEY(manufacturer_id) REFERENCES manufacturers(id),
                  FOREIGN KEY(type_id) REFERENCES product_types(id))''')

    # product_specifications: thông số của manufacturers
    c.execute('''CREATE TABLE IF NOT EXISTS product_specifications
                 (manufacturer_id INTEGER,
                  indicator_id INTEGER,
                  specification_value TEXT,
                  PRIMARY KEY(manufacturer_id, indicator_id),
                  FOREIGN KEY(manufacturer_id) REFERENCES manufacturers(id),
                  FOREIGN KEY(indicator_id) REFERENCES indicators(id))''')

    # products: sản phẩm chính
    c.execute('''CREATE TABLE IF NOT EXISTS products
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  name TEXT,
                  ma_san_pham TEXT UNIQUE,
                  note TEXT,
                  unit TEXT,
                  quantity INTEGER)''')

    # product_type_mapping_products: mapping products <-> product_types
    c.execute('''CREATE TABLE IF NOT EXISTS product_type_mapping_products
                 (product_id INTEGER,
                  type_id INTEGER,
                  PRIMARY KEY(product_id, type_id),
                  FOREIGN KEY(product_id) REFERENCES products(id),
                  FOREIGN KEY(type_id) REFERENCES product_types(id))''')

    # reference_products: mapping sản phẩm chính -> manufacturers
    c.execute('''CREATE TABLE IF NOT EXISTS reference_products
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  product_id INTEGER,
                  manufacturer_id INTEGER,
                  sort_order INTEGER DEFAULT 0,
                  FOREIGN KEY(product_id) REFERENCES products(id),
                  FOREIGN KEY(manufacturer_id) REFERENCES manufacturers(id),
                  UNIQUE(product_id, manufacturer_id))''')

    # projects: giữ nếu vẫn quản lý dự án
    c.execute('''CREATE TABLE IF NOT EXISTS projects
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  name TEXT,
                  ma_du_an TEXT UNIQUE,
                  ghi_chu TEXT)''')

    # product_projects: mapping sản phẩm <-> dự án
    c.execute('''CREATE TABLE IF NOT EXISTS product_projects
                 (product_id INTEGER,
                  project_id INTEGER,
                  PRIMARY KEY(product_id, project_id),
                  FOREIGN KEY(product_id) REFERENCES products(id),
                  FOREIGN KEY(project_id) REFERENCES projects(id))''')

    # product_hidden_indicators: ẩn chỉ tiêu per sản phẩm
    c.execute('''CREATE TABLE IF NOT EXISTS product_hidden_indicators
                 (product_id INTEGER,
                  tab_name TEXT,
                  indicator_id INTEGER,
                  PRIMARY KEY(product_id, tab_name, indicator_id),
                  FOREIGN KEY(product_id) REFERENCES products(id),
                  FOREIGN KEY(indicator_id) REFERENCES indicators(id))''')

    # product_custom_indicators: tùy chỉnh per sản phẩm
    c.execute('''CREATE TABLE IF NOT EXISTS product_custom_indicators
                 (product_id INTEGER,
                  tab_name TEXT,
                  indicator_id TEXT,
                  custom_value TEXT,
                  PRIMARY KEY(product_id, tab_name, indicator_id),
                  FOREIGN KEY(product_id) REFERENCES products(id))''')

    # units
    c.execute('''CREATE TABLE IF NOT EXISTS units
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  dai_luong TEXT,
                  ten_don_vi TEXT,
                  ki_hieu_don_vi TEXT UNIQUE)''')

    # Sample data initialization
    c.execute("SELECT COUNT(*) FROM units")
    if c.fetchone()[0] == 0:
        sample_units = [
            ("Độ dài", "mét", "m"),
            ("Khối lượng", "kilôgam", "kg"),
            ("Thời gian", "giây", "s"),
            ("Số lượng", "bộ", "Bộ"),
        ]
        for dai_luong, ten_don_vi, ki_hieu in sample_units:
            c.execute("INSERT OR IGNORE INTO units (dai_luong, ten_don_vi, ki_hieu_don_vi) VALUES (?, ?, ?)",
                      (dai_luong, ten_don_vi, ki_hieu))

    c.execute("CREATE INDEX IF NOT EXISTS idx_indicators_type ON indicators(type_id)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_product_specifications_manufacturer ON product_specifications(manufacturer_id)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_reference_products_product ON reference_products(product_id)")

    conn.commit()
    conn.close()