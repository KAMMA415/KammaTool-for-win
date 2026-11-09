
import ttkbootstrap as ttk
from gui import DongToolGUI
from utils import enable_high_dpi

def main():
    """程序入口"""
    enable_high_dpi()
    # 使用lumen主题，更具现代感
    root = ttk.Window(themename="lumen")
    app = DongToolGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
