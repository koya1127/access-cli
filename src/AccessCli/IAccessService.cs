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

    /// <summary>全モジュール・全フォームをディレクトリにテキスト出力する</summary>
    void ExportAll(string dbPath, string outputDir);

    /// <summary>ディレクトリから全モジュール・全フォームを一括インポートする</summary>
    void ImportAll(string dbPath, string inputDir);

    /// <summary>テーブル一覧を返す</summary>
    IReadOnlyList<string> ListTables(string dbPath);

    /// <summary>SELECT文を実行し結果を返す（1行目はカラム名）</summary>
    IReadOnlyList<string[]> QuerySql(string dbPath, string sql);

    /// <summary>INSERT/UPDATE/DELETE を実行する</summary>
    void ExecSql(string dbPath, string sql);

    /// <summary>保存クエリ（QueryDef）一覧と SQL を返す</summary>
    IReadOnlyList<QueryInfo> ListQueries(string dbPath);

    /// <summary>保存クエリ（QueryDef）の SQL を返す</summary>
    string GetQuerySql(string dbPath, string queryName);

    /// <summary>全クエリの SQL をディレクトリにファイル出力する</summary>
    void ExportAllQueries(string dbPath, string outputDir);

    /// <summary>保存クエリ（QueryDef）の SQL を書き換える</summary>
    void SetQuerySql(string dbPath, string queryName, string sql);
}

public record ModuleInfo(string Name, string Type, int Lines);
public record ControlInfo(string Name, string Type, string? Caption);
public record QueryInfo(string Name, string Sql);
