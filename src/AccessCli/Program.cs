using System.CommandLine;
using System.CommandLine.Invocation;
using AccessCli;

// 自動更新チェック（バックグラウンド）
var updateTask = Updater.CheckAsync();

var svc = new AccessService();

// ─── 共通引数 ──────────────────────────────────────────────────────
var dbArg    = new Argument<FileInfo>("db_path") { Description = "対象のAccdbファイルパス" }.AcceptExistingOnly();
var modArg   = new Argument<string>("module_name") { Description = "モジュール名" };
var formArg  = new Argument<string>("form_name") { Description = "フォーム名" };
var outputOpt = new Option<FileInfo?>("--output", "-o");
outputOpt.Description = "出力ファイルパス（省略時は標準出力）";

// ─── list-modules ─────────────────────────────────────────────────
var listModulesCmd = new Command("list-modules", "VBAモジュール一覧を表示");
listModulesCmd.Add(dbArg);
listModulesCmd.SetAction(r =>
{
    var db = r.GetValue(dbArg)!;
    foreach (var m in svc.ListModules(db.FullName))
        Console.WriteLine($"{m.Type,-12} {m.Name,-40} ({m.Lines}行)");
});

// ─── read-vba ─────────────────────────────────────────────────────
var readVbaCmd = new Command("read-vba", "VBAコードを取得");
readVbaCmd.Add(dbArg);
readVbaCmd.Add(modArg);
readVbaCmd.Add(outputOpt);
readVbaCmd.SetAction(r =>
{
    var db     = r.GetValue(dbArg)!;
    var mod    = r.GetValue(modArg)!;
    var output = r.GetValue(outputOpt);
    string code = svc.ReadVba(db.FullName, mod);
    if (output is not null)
    {
        File.WriteAllText(output.FullName, code, System.Text.Encoding.UTF8);
        Console.WriteLine($"書き出し完了: {output.FullName}");
    }
    else
    {
        Console.Write(code);
    }
});

// ─── write-vba ────────────────────────────────────────────────────
var writeVbaCmd = new Command("write-vba", "ファイルからVBAコードを書き込む");
var codeFileArg = new Argument<FileInfo>("code_file") { Description = "VBAコードを含むファイルパス" }.AcceptExistingOnly();
writeVbaCmd.Add(dbArg);
writeVbaCmd.Add(modArg);
writeVbaCmd.Add(codeFileArg);
writeVbaCmd.SetAction(r =>
{
    var db       = r.GetValue(dbArg)!;
    var mod      = r.GetValue(modArg)!;
    var codeFile = r.GetValue(codeFileArg)!;
    string code  = File.ReadAllText(codeFile.FullName, System.Text.Encoding.UTF8);
    svc.WriteVba(db.FullName, mod, code);
    Console.WriteLine($"書き込み完了: {mod}");
});

// ─── list-forms ───────────────────────────────────────────────────
var listFormsCmd = new Command("list-forms", "フォーム一覧を表示");
listFormsCmd.Add(dbArg);
listFormsCmd.SetAction(r =>
{
    var db = r.GetValue(dbArg)!;
    foreach (var f in svc.ListForms(db.FullName))
        Console.WriteLine(f);
});

// ─── list-controls ────────────────────────────────────────────────
var listControlsCmd = new Command("list-controls", "フォームのコントロール一覧を表示");
listControlsCmd.Add(dbArg);
listControlsCmd.Add(formArg);
listControlsCmd.SetAction(r =>
{
    var db   = r.GetValue(dbArg)!;
    var form = r.GetValue(formArg)!;
    foreach (var c in svc.ListControls(db.FullName, form))
    {
        string cap = c.Caption is not null ? $"  Caption=\"{c.Caption}\"" : "";
        Console.WriteLine($"{c.Type,-12} {c.Name,-40}{cap}");
    }
});

// ─── set-caption ──────────────────────────────────────────────────
var setCaptionCmd = new Command("set-caption", "CaptionをバイナリAPI直書き換え（同じバイト長のみ安全）");
var oldCapArg = new Argument<string>("old_caption") { Description = "現在のCaption文字列" };
var newCapArg = new Argument<string>("new_caption") { Description = "新しいCaption文字列" };
setCaptionCmd.Add(dbArg);
setCaptionCmd.Add(oldCapArg);
setCaptionCmd.Add(newCapArg);
setCaptionCmd.SetAction(r =>
{
    var db  = r.GetValue(dbArg)!;
    var old = r.GetValue(oldCapArg)!;
    var @new = r.GetValue(newCapArg)!;
    int count = svc.SetCaption(db.FullName, old, @new);
    if (count == 0)
    {
        Console.Error.WriteLine($"\"{old}\" が見つかりませんでした");
        Environment.Exit(1);
    }
    Console.WriteLine($"\"{old}\" → \"{@new}\" ({count}箇所) 完了");
});

// ─── export-form ──────────────────────────────────────────────────
var exportFormCmd = new Command("export-form", "フォーム定義をテキストにエクスポート（SaveAsText）");
var outPathArg = new Argument<FileInfo>("output_path") { Description = "出力先ファイルパス" };
exportFormCmd.Add(dbArg);
exportFormCmd.Add(formArg);
exportFormCmd.Add(outPathArg);
exportFormCmd.SetAction(r =>
{
    var db      = r.GetValue(dbArg)!;
    var form    = r.GetValue(formArg)!;
    var outPath = r.GetValue(outPathArg)!;
    svc.ExportForm(db.FullName, form, outPath.FullName);
    Console.WriteLine($"エクスポート完了: {outPath.FullName}");
});

// ─── import-form ──────────────────────────────────────────────────
var importFormCmd = new Command("import-form", "テキストからフォームをインポート（LoadFromText）");
var inPathArg = new Argument<FileInfo>("input_path") { Description = "インポート元ファイルパス" }.AcceptExistingOnly();
importFormCmd.Add(dbArg);
importFormCmd.Add(formArg);
importFormCmd.Add(inPathArg);
importFormCmd.SetAction(r =>
{
    var db     = r.GetValue(dbArg)!;
    var form   = r.GetValue(formArg)!;
    var inPath = r.GetValue(inPathArg)!;
    svc.ImportForm(db.FullName, form, inPath.FullName);
    Console.WriteLine($"インポート完了: {form}");
});

// ─── export-all ───────────────────────────────────────────────
var exportAllCmd = new Command("export-all", "全モジュール・全フォームをディレクトリにテキスト出力");
var outDirArg = new Argument<DirectoryInfo>("output_dir") { Description = "出力先ディレクトリ" };
exportAllCmd.Add(dbArg);
exportAllCmd.Add(outDirArg);
exportAllCmd.SetAction(r =>
{
    var db     = r.GetValue(dbArg)!;
    var outDir = r.GetValue(outDirArg)!;
    svc.ExportAll(db.FullName, outDir.FullName);
    Console.WriteLine($"エクスポート完了: {outDir.FullName}");
});

// ─── import-all ───────────────────────────────────────────────
var importAllCmd = new Command("import-all", "ディレクトリから全モジュール・全フォームを一括インポート");
var inDirArg = new Argument<DirectoryInfo>("input_dir") { Description = "インポート元ディレクトリ" };
importAllCmd.Add(dbArg);
importAllCmd.Add(inDirArg);
importAllCmd.SetAction(r =>
{
    var db    = r.GetValue(dbArg)!;
    var inDir = r.GetValue(inDirArg)!;
    svc.ImportAll(db.FullName, inDir.FullName);
    Console.WriteLine("インポート完了");
});

// ─── list-tables ──────────────────────────────────────────────
var listTablesCmd = new Command("list-tables", "テーブル一覧を表示");
listTablesCmd.Add(dbArg);
listTablesCmd.SetAction(r =>
{
    var db = r.GetValue(dbArg)!;
    foreach (var t in svc.ListTables(db.FullName))
        Console.WriteLine(t);
});

// ─── query-sql ────────────────────────────────────────────────
var querySqlCmd = new Command("query-sql", "SELECT文を実行して結果をTSV出力");
var sqlArg     = new Argument<string>("sql")      { Description = "SQL文（SELECT）" };
var sqlFileOpt = new Option<FileInfo?>("--sql-file", "-f");
sqlFileOpt.Description = "SQLをファイルから読み込む（日本語テーブル名などに使用）";
querySqlCmd.Add(dbArg);
querySqlCmd.Add(sqlArg);
querySqlCmd.Add(sqlFileOpt);
querySqlCmd.SetAction(r =>
{
    var db      = r.GetValue(dbArg)!;
    var sqlDirect = r.GetValue(sqlArg)!;
    var sqlFile = r.GetValue(sqlFileOpt);
    var sql = sqlFile is not null
        ? File.ReadAllText(sqlFile.FullName).Trim().TrimStart('\uFEFF') // BOM除去
        : sqlDirect;
    var rows = svc.QuerySql(db.FullName, sql);
    foreach (var row in rows)
        Console.WriteLine(string.Join("\t", row));
});

// ─── exec-sql ─────────────────────────────────────────────────
var execSqlCmd = new Command("exec-sql", "INSERT/UPDATE/DELETE を実行");
execSqlCmd.Add(dbArg);
execSqlCmd.Add(sqlArg);
execSqlCmd.Add(sqlFileOpt);
execSqlCmd.SetAction(r =>
{
    var db      = r.GetValue(dbArg)!;
    var sqlDirect = r.GetValue(sqlArg)!;
    var sqlFile = r.GetValue(sqlFileOpt);
    var sql = sqlFile is not null
        ? File.ReadAllText(sqlFile.FullName).Trim().TrimStart('\uFEFF') // BOM除去
        : sqlDirect;
    svc.ExecSql(db.FullName, sql);
    Console.WriteLine("実行完了");
});

// ─── ルートコマンド ────────────────────────────────────────────────
var rootCmd = new RootCommand("Microsoft Access VBA・フォーム編集CLI");
rootCmd.Add(listModulesCmd);
rootCmd.Add(readVbaCmd);
rootCmd.Add(writeVbaCmd);
rootCmd.Add(listFormsCmd);
rootCmd.Add(listControlsCmd);
rootCmd.Add(setCaptionCmd);
rootCmd.Add(exportFormCmd);
rootCmd.Add(importFormCmd);
rootCmd.Add(exportAllCmd);
rootCmd.Add(importAllCmd);
rootCmd.Add(listTablesCmd);
rootCmd.Add(querySqlCmd);
rootCmd.Add(execSqlCmd);

// 更新チェックをすぐタイムアウトさせて（起動遅延なし）
try { await updateTask.WaitAsync(TimeSpan.Zero); } catch { }

return await rootCmd.Parse(args).InvokeAsync();
