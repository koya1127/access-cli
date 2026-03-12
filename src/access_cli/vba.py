"""VBA読み書き処理"""
from .access_com import open_access


def list_modules(db_path: str) -> list[dict]:
    """VBAモジュール一覧を返す"""
    with open_access(db_path) as access:
        proj = access.VBE.VBProjects.Item(1)
        result = []
        for i in range(1, proj.VBComponents.Count + 1):
            comp = proj.VBComponents.Item(i)
            type_name = {1: "Module", 2: "Class", 100: "Form/Report"}.get(comp.Type, str(comp.Type))
            lines = comp.CodeModule.CountOfLines
            result.append({
                "name": comp.Name,
                "type": type_name,
                "lines": lines,
            })
        return result


def read_vba(db_path: str, module_name: str) -> str:
    """指定モジュールのVBAコードを返す"""
    with open_access(db_path) as access:
        proj = access.VBE.VBProjects.Item(1)
        comp = proj.VBComponents(module_name)
        cm = comp.CodeModule
        if cm.CountOfLines == 0:
            return ""
        return cm.Lines(1, cm.CountOfLines)


def write_vba(db_path: str, module_name: str, code: str) -> None:
    """指定モジュールにVBAコードを書き込む"""
    with open_access(db_path) as access:
        proj = access.VBE.VBProjects.Item(1)
        comp = proj.VBComponents(module_name)
        cm = comp.CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        cm.InsertLines(1, code)
