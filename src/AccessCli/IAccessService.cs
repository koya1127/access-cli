namespace AccessCli;

/// <summary>
/// Microsoft Access操作の抽象インターフェース（テスト時はモックに差し替え）
/// </summary>
public interface IAccessService
{
    /// <summary>VBAモジュール一覧を返す</summary>
    IReadOnlyList<ModuleInfo> ListModules(string dbPath);

    /// <summary>指定モジュールのVBAコードを返す</summary>
    string ReadVba(string dbPath, string moduleName);

    /// <summary>指定モジュールにVBAコードを書き込む</summary>
    void WriteVba(string dbPath, string moduleName, string code);

    /// <summary>フォーム名一覧を返す</summary>
    IReadOnlyList<string> ListForms(string dbPath);

    /// <summary>フォームのコントロール一覧を返す</summary>
    IReadOnlyList<ControlInfo> ListControls(string dbPath, string formName);

    /// <summary>フォーム定義をテキストにエクスポート</summary>
    void ExportForm(string dbPath, string formName, string outputPath);

    /// <summary>テキストからフォーム定義をインポート</summary>
    void ImportForm(string dbPath, string formName, string inputPath);

    /// <summary>
    /// accdbバイナリ内のCaption文字列を直接置換する。
    /// 戻り値: 置換した箇所数
    /// </summary>
    int SetCaption(string dbPath, string oldCaption, string newCaption);
}

public record ModuleInfo(string Name, string Type, int Lines);
public record ControlInfo(string Name, string Type, string? Caption);
