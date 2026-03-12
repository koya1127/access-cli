"""フォーム操作処理"""
import os
import struct
from .access_com import open_access, close_startup_forms


def list_forms(db_path: str) -> list[str]:
    """フォーム一覧を返す"""
    with open_access(db_path) as access:
        proj = access.VBE.VBProjects.Item(1)
        forms = []
        for i in range(1, proj.VBComponents.Count + 1):
            comp = proj.VBComponents.Item(i)
            if comp.Type == 100 and comp.Name.startswith("Form_"):
                forms.append(comp.Name[5:])  # "Form_" を除去
        return forms


def list_controls(db_path: str, form_name: str) -> list[dict]:
    """フォームのコントロール一覧（Caption付き）を返す"""
    type_map = {
        100: "Label", 101: "Rectangle", 102: "Line",
        104: "Button", 106: "OptionGroup", 107: "OptionButton",
        108: "Toggle", 109: "TextBox", 110: "ComboBox",
        111: "ListBox", 112: "SubForm", 118: "PageBreak",
        122: "Image", 123: "Tab",
    }
    with open_access(db_path, visible=True) as access:
        close_startup_forms(access)
        access.DoCmd.OpenForm(form_name, 1)  # acDesign

        frm = access.Forms(form_name)
        result = []
        for i in range(frm.Controls.Count):
            ctl = frm.Controls(i)
            entry = {
                "name": ctl.Name,
                "type": type_map.get(ctl.ControlType, str(ctl.ControlType)),
                "caption": None,
            }
            try:
                entry["caption"] = ctl.Caption
            except Exception:
                pass
            result.append(entry)

        access.DoCmd.Close(2, form_name, 0)  # acSaveNo
        return result


def set_caption(db_path: str, old_caption: str, new_caption: str) -> int:
    """
    accdbバイナリ内のキャプション文字列を直接置換する。
    同じバイト長でも異なる長さでも対応（長さプレフィックスも更新）。
    戻り値: 置換した箇所数
    """
    old_bytes = old_caption.encode("utf-16-le")
    new_bytes = new_caption.encode("utf-16-le")
    old_len = struct.pack("<I", len(old_bytes))
    new_len = struct.pack("<I", len(new_bytes))

    old_pattern = old_len + old_bytes
    new_pattern = new_len + new_bytes

    with open(db_path, "rb") as f:
        data = f.read()

    count = data.count(old_pattern)
    if count == 0:
        return 0

    data = data.replace(old_pattern, new_pattern)
    with open(db_path, "wb") as f:
        f.write(data)

    return count


def export_form(db_path: str, form_name: str, output_path: str) -> None:
    """フォームをテキストファイルにエクスポート"""
    output_path = os.path.abspath(output_path)
    with open_access(db_path) as access:
        access.SaveAsText(2, form_name, output_path)
