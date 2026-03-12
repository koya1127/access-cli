"""access-cli エントリーポイント"""
import sys
import click
from . import vba as vba_ops
from . import forms as form_ops


@click.group()
def cli():
    """Microsoft Access VBA・フォーム編集CLI"""
    pass


# ─── VBA ──────────────────────────────────────────

@cli.command("list-modules")
@click.argument("db_path")
def list_modules(db_path):
    """VBAモジュール一覧を表示"""
    modules = vba_ops.list_modules(db_path)
    for m in modules:
        click.echo(f"{m['type']:12} {m['name']:40} ({m['lines']}行)")


@cli.command("read-vba")
@click.argument("db_path")
@click.argument("module_name")
@click.option("-o", "--output", default=None, help="出力ファイルパス（省略時は標準出力）")
def read_vba(db_path, module_name, output):
    """VBAコードを取得"""
    code = vba_ops.read_vba(db_path, module_name)
    if output:
        with open(output, "w", encoding="utf-8") as f:
            f.write(code)
        click.echo(f"書き出し完了: {output}")
    else:
        click.echo(code)


@cli.command("write-vba")
@click.argument("db_path")
@click.argument("module_name")
@click.argument("code_file")
def write_vba(db_path, module_name, code_file):
    """ファイルからVBAコードを書き込む"""
    with open(code_file, encoding="utf-8") as f:
        code = f.read()
    vba_ops.write_vba(db_path, module_name, code)
    click.echo(f"書き込み完了: {module_name}")


# ─── フォーム ──────────────────────────────────────

@cli.command("list-forms")
@click.argument("db_path")
def list_forms(db_path):
    """フォーム一覧を表示"""
    forms = form_ops.list_forms(db_path)
    for f in forms:
        click.echo(f)


@cli.command("list-controls")
@click.argument("db_path")
@click.argument("form_name")
def list_controls(db_path, form_name):
    """フォームのコントロール一覧とCaptionを表示"""
    controls = form_ops.list_controls(db_path, form_name)
    for c in controls:
        caption = f'  Caption="{c["caption"]}"' if c["caption"] else ""
        click.echo(f"{c['type']:12} {c['name']:40}{caption}")


@cli.command("set-caption")
@click.argument("db_path")
@click.argument("old_caption")
@click.argument("new_caption")
def set_caption(db_path, old_caption, new_caption):
    """
    フォーム内のCaptionを直接書き換える。
    OLD_CAPTION: 現在の文字列  NEW_CAPTION: 新しい文字列
    """
    count = form_ops.set_caption(db_path, old_caption, new_caption)
    if count == 0:
        click.echo(f'"{old_caption}" が見つかりませんでした', err=True)
        sys.exit(1)
    click.echo(f'"{old_caption}" → "{new_caption}" ({count}箇所) 完了')


@cli.command("export-form")
@click.argument("db_path")
@click.argument("form_name")
@click.argument("output_path")
def export_form(db_path, form_name, output_path):
    """フォーム定義をテキストファイルにエクスポート"""
    form_ops.export_form(db_path, form_name, output_path)
    click.echo(f"エクスポート完了: {output_path}")


if __name__ == "__main__":
    cli()
