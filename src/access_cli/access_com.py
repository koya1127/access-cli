"""Access COM操作の共通処理"""
import os
from contextlib import contextmanager


@contextmanager
def open_access(db_path: str, exclusive: bool = False, visible: bool = False):
    """Accessを起動してDBを開くコンテキストマネージャ"""
    import win32com.client as win32

    db_path = os.path.abspath(db_path)
    if not os.path.exists(db_path):
        raise FileNotFoundError(f"ファイルが見つかりません: {db_path}")

    access = win32.DispatchEx("Access.Application")
    access.Visible = visible
    access.AutomationSecurity = 3  # スタートアップ・マクロを無効化
    access.OpenCurrentDatabase(db_path, exclusive)

    try:
        yield access
    finally:
        try:
            access.CloseCurrentDatabase()
            access.Quit()
        except Exception:
            pass


def close_startup_forms(access) -> list[str]:
    """起動時に自動で開いたフォームをすべて閉じる"""
    closed = []
    access.DoCmd.SetWarnings(False)
    while access.Forms.Count > 0:
        name = access.Forms(0).Name
        access.DoCmd.Close(2, name, 0)  # acForm, acSaveNo
        closed.append(name)
    access.DoCmd.SetWarnings(True)
    return closed
