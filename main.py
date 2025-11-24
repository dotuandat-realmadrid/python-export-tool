import tkinter as tk
from database import init_db
from menu import MainMenu

if __name__ == "__main__":
    init_db()
    root = tk.Tk()
    app = MainMenu(root)
    root.mainloop()