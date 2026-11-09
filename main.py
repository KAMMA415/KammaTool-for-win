
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

import ttkbootstrap as ttk
from gui import DongToolGUI
from utils import enable_high_dpi


def _smoke_test(app):
    app.root.update_idletasks()
    tabs = [app.notebook.tab(tab, "text") for tab in app.notebook.tabs()]
    if tabs != ["📁文件工具", "📊EXCEL工具", "📝WORD工具"] or len(app.file_icons) != 4:
        raise RuntimeError("The packaged application did not load all public tabs and icons.")

    with TemporaryDirectory(prefix="kammatool-smoke-") as directory:
        folder = Path(directory)
        (folder / "example.txt").write_text("KammaTool", encoding="utf-8")
        renamed = app.file_processor.batch_rename(folder, "checked_")
        if not renamed or not (folder / "checked_example.txt").is_file():
            raise RuntimeError("The packaged file rename test failed.")


def main():
    """程序入口"""
    enable_high_dpi()
    # 使用lumen主题，更具现代感
    root = ttk.Window(themename="lumen")
    smoke_test = sys.argv[1:] == ["--smoke-test"]
    if smoke_test:
        root.withdraw()
    app = DongToolGUI(root)
    if smoke_test:
        try:
            _smoke_test(app)
        finally:
            root.destroy()
        return
    root.mainloop()

if __name__ == "__main__":
    main()
