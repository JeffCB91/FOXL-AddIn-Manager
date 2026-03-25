import tkinter as tk
from ui.user_window import UserWindow

if __name__ == "__main__":
    root = tk.Tk()
    app = UserWindow(root)
    root.mainloop()
