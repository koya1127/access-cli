using System.Buffers.Binary;
using System.Runtime.InteropServices;

namespace AccessCli;

/// <summary>
/// Microsoft Access COM自動化による実装。
/// Windows専用（win32com相当をdynamic COMで実現）。
/// </summary>
public sealed class AccessService : IAccessService
{
    // ControlType constants (Access)
    private static readonly Dictionary<int, string> ControlTypeMap = new()
    {
        [100] = "Label",    [101] = "Rectangle", [102] = "Line",
        [104] = "Button",   [106] = "OptionGroup",[107] = "OptionButton",
        [108] = "Toggle",   [109] = "TextBox",    [110] = "ComboBox",
        [111] = "ListBox",  [112] = "SubForm",    [118] = "PageBreak",
        [122] = "Image",    [123] = "Tab",
    };

    // VBComponentType constants
    private static readonly Dictionary<int, string> ModuleTypeMap = new()
    {
        [1]   = "Module",
        [2]   = "Class",
        [100] = "Form/Report",
    };

    public IReadOnlyList<ModuleInfo> ListModules(string dbPath)
    {
        dbPath = ResolveAbsPath(dbPath);
        dynamic access = CreateAccessApp();
        try
        {
            access.Visible = false;
            access.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
            access.OpenCurrentDatabase(dbPath, false);

            dynamic proj = access.VBE.VBProjects.Item(1);
            var result = new List<ModuleInfo>();
            int count = proj.VBComponents.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic comp = proj.VBComponents.Item(i);
                string typeName = ModuleTypeMap.TryGetValue((int)comp.Type, out var t) ? t : comp.Type.ToString();
                int lines = (int)comp.CodeModule.CountOfLines;
                result.Add(new ModuleInfo((string)comp.Name, typeName, lines));
            }
            return result;
        }
        finally
        {
            SafeQuit(access);
        }
    }

    public string ReadVba(string dbPath, string moduleName)
    {
        dbPath = ResolveAbsPath(dbPath);
        dynamic access = CreateAccessApp();
        try
        {
            access.Visible = false;
            access.AutomationSecurity = 3;
            access.OpenCurrentDatabase(dbPath, false);

            dynamic proj = access.VBE.VBProjects.Item(1);
            dynamic comp = proj.VBComponents(moduleName);
            dynamic cm = comp.CodeModule;
            int lines = (int)cm.CountOfLines;
            if (lines == 0) return string.Empty;
            return (string)cm.Lines(1, lines);
        }
        finally
        {
            SafeQuit(access);
        }
    }

    public void WriteVba(string dbPath, string moduleName, string code)
    {
        dbPath = ResolveAbsPath(dbPath);
        dynamic access = CreateAccessApp();
        try
        {
            access.Visible = false;
            access.AutomationSecurity = 3;
            access.OpenCurrentDatabase(dbPath, false);

            dynamic proj = access.VBE.VBProjects.Item(1);
            dynamic comp = proj.VBComponents(moduleName);
            dynamic cm = comp.CodeModule;
            int existing = (int)cm.CountOfLines;
            if (existing > 0)
                cm.DeleteLines(1, existing);
            cm.InsertLines(1, code);

            access.CloseCurrentDatabase();
        }
        finally
        {
            SafeQuit(access);
        }
    }

    public IReadOnlyList<string> ListForms(string dbPath)
    {
        dbPath = ResolveAbsPath(dbPath);
        dynamic access = CreateAccessApp();
        try
        {
            access.Visible = false;
            access.AutomationSecurity = 3;
            access.OpenCurrentDatabase(dbPath, false);

            dynamic proj = access.VBE.VBProjects.Item(1);
            var result = new List<string>();
            int count = proj.VBComponents.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic comp = proj.VBComponents.Item(i);
                if ((int)comp.Type == 100)
                {
                    string name = (string)comp.Name;
                    if (name.StartsWith("Form_"))
                        result.Add(name[5..]);
                }
            }
            return result;
        }
        finally
        {
            SafeQuit(access);
        }
    }

    public IReadOnlyList<ControlInfo> ListControls(string dbPath, string formName)
    {
        dbPath = ResolveAbsPath(dbPath);
        dynamic access = CreateAccessApp();
        try
        {
            access.Visible = true;
            access.AutomationSecurity = 3;
            access.OpenCurrentDatabase(dbPath, false);

            // Close any startup forms
            access.DoCmd.SetWarnings(false);
            while ((int)access.Forms.Count > 0)
            {
                string fn = (string)access.Forms(0).Name;
                access.DoCmd.Close(2, fn, 0);
            }
            access.DoCmd.SetWarnings(true);

            access.DoCmd.OpenForm(formName, 1); // acDesign = 1
            dynamic frm = access.Forms(formName);
            var result = new List<ControlInfo>();
            int count = (int)frm.Controls.Count;
            for (int i = 0; i < count; i++)
            {
                dynamic ctl = frm.Controls(i);
                string ctlName = (string)ctl.Name;
                int ctlType = (int)ctl.ControlType;
                string typeName = ControlTypeMap.TryGetValue(ctlType, out var tn) ? tn : ctlType.ToString();
                string? caption = null;
                try { caption = (string?)ctl.Caption; } catch { }
                result.Add(new ControlInfo(ctlName, typeName, caption));
            }
            access.DoCmd.Close(2, formName, 0); // acSaveNo
            return result;
        }
        finally
        {
            SafeQuit(access);
        }
    }

    public void ExportForm(string dbPath, string formName, string outputPath)
    {
        dbPath = ResolveAbsPath(dbPath);
        outputPath = Path.GetFullPath(outputPath);
        dynamic access = CreateAccessApp();
        try
        {
            access.Visible = false;
            access.AutomationSecurity = 3;
            access.OpenCurrentDatabase(dbPath, false);
            access.SaveAsText(2, formName, outputPath); // acForm = 2
        }
        finally
        {
            SafeQuit(access);
        }
    }

    public void ImportForm(string dbPath, string formName, string inputPath)
    {
        dbPath = ResolveAbsPath(dbPath);
        inputPath = Path.GetFullPath(inputPath);
        if (!File.Exists(inputPath))
            throw new FileNotFoundException($"インポートファイルが見つかりません: {inputPath}");

        dynamic access = CreateAccessApp();
        try
        {
            access.Visible = true;
            access.AutomationSecurity = 3;
            access.OpenCurrentDatabase(dbPath, true); // exclusive
            access.LoadFromText(2, formName, inputPath); // acForm = 2
        }
        finally
        {
            SafeQuit(access);
        }
    }

    public int SetCaption(string dbPath, string oldCaption, string newCaption)
    {
        dbPath = ResolveAbsPath(dbPath);

        byte[] oldBytes = System.Text.Encoding.Unicode.GetBytes(oldCaption);
        byte[] newBytes = System.Text.Encoding.Unicode.GetBytes(newCaption);

        // Format: [4-byte LE length][UTF-16-LE string]
        byte[] oldPattern = BuildPattern(oldBytes);
        byte[] newPattern = BuildPattern(newBytes);

        byte[] data = File.ReadAllBytes(dbPath);

        int count = 0;
        int idx = 0;
        while ((idx = IndexOf(data, oldPattern, idx)) >= 0)
        {
            count++;
            idx++;
        }
        if (count == 0) return 0;

        byte[] result = ReplaceAll(data, oldPattern, newPattern);
        File.WriteAllBytes(dbPath, result);
        return count;
    }

    public void ExportAll(string dbPath, string outputDir)
    {
        dbPath = ResolveAbsPath(dbPath);
        string modulesDir = Path.Combine(outputDir, "modules");
        string formsDir   = Path.Combine(outputDir, "forms");
        Directory.CreateDirectory(modulesDir);
        Directory.CreateDirectory(formsDir);

        dynamic access = CreateAccessApp();
        try
        {
            access.Visible = false;
            access.AutomationSecurity = 3;
            access.OpenCurrentDatabase(dbPath, false);

            dynamic proj = access.VBE.VBProjects.Item(1);
            int count = proj.VBComponents.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic comp = proj.VBComponents.Item(i);
                string name = (string)comp.Name;
                int type = (int)comp.Type;

                if (type == 1 || type == 2) // Module / Class
                {
                    int lines = (int)comp.CodeModule.CountOfLines;
                    string code = lines > 0 ? (string)comp.CodeModule.Lines(1, lines) : "";
                    File.WriteAllText(Path.Combine(modulesDir, name + ".bas"), code, System.Text.Encoding.UTF8);
                    Console.WriteLine($"[module] {name}");
                }
                else if (type == 100) // Form / Report
                {
                    string accessName;
                    int objectType;
                    if (name.StartsWith("Form_"))        { accessName = name[5..]; objectType = 2; }
                    else if (name.StartsWith("Report_")) { accessName = name[7..]; objectType = 3; }
                    else continue;

                    string outPath = Path.Combine(formsDir, name + ".form");
                    access.SaveAsText(objectType, accessName, outPath);
                    Console.WriteLine($"[form]   {name}");
                }
            }
        }
        finally
        {
            SafeQuit(access);
        }
    }

    public void ImportAll(string dbPath, string inputDir)
    {
        dbPath = ResolveAbsPath(dbPath);

        dynamic access = CreateAccessApp();
        try
        {
            access.Visible = true;
            access.AutomationSecurity = 3;
            access.OpenCurrentDatabase(dbPath, true); // exclusive

            // モジュール (.bas)
            string modulesDir = Path.Combine(inputDir, "modules");
            if (Directory.Exists(modulesDir))
            {
                dynamic proj = access.VBE.VBProjects.Item(1);
                foreach (string file in Directory.GetFiles(modulesDir, "*.bas"))
                {
                    string moduleName = Path.GetFileNameWithoutExtension(file);
                    string code = File.ReadAllText(file, System.Text.Encoding.UTF8);
                    dynamic comp = proj.VBComponents(moduleName);
                    dynamic cm = comp.CodeModule;
                    int existing = (int)cm.CountOfLines;
                    if (existing > 0) cm.DeleteLines(1, existing);
                    cm.InsertLines(1, code);
                    Console.WriteLine($"[module] {moduleName}");
                }
            }

            // フォーム・レポート (.form)
            string formsDir = Path.Combine(inputDir, "forms");
            if (Directory.Exists(formsDir))
            {
                foreach (string file in Directory.GetFiles(formsDir, "*.form"))
                {
                    string compName = Path.GetFileNameWithoutExtension(file);
                    string accessName;
                    int objectType;
                    if (compName.StartsWith("Form_"))        { accessName = compName[5..]; objectType = 2; }
                    else if (compName.StartsWith("Report_")) { accessName = compName[7..]; objectType = 3; }
                    else continue;

                    access.LoadFromText(objectType, accessName, Path.GetFullPath(file));
                    Console.WriteLine($"[form]   {compName}");
                }
            }

            access.CloseCurrentDatabase();
        }
        finally
        {
            SafeQuit(access);
        }
    }

    public IReadOnlyList<string> ListTables(string dbPath)
    {
        dbPath = ResolveAbsPath(dbPath);
        dynamic access = CreateAccessApp();
        try
        {
            access.Visible = false;
            access.AutomationSecurity = 3;
            access.OpenCurrentDatabase(dbPath, false);

            dynamic db = access.CurrentDb();
            var result = new List<string>();
            int count = (int)db.TableDefs.Count;
            for (int i = 0; i < count; i++)
            {
                dynamic td = db.TableDefs(i);
                string name = (string)td.Name;
                if (name.StartsWith("MSys")) continue;
                string connect = "";
                try { connect = (string)td.Connect; } catch { }
                string label = string.IsNullOrEmpty(connect) ? name : $"{name}  [linked: {connect}]";
                result.Add(label);
            }
            return result;
        }
        finally
        {
            SafeQuit(access);
        }
    }

    public IReadOnlyList<QueryInfo> ListQueries(string dbPath)
    {
        dbPath = ResolveAbsPath(dbPath);
        dynamic db = OpenDaoDatabase(dbPath);
        try
        {
            var result = new List<QueryInfo>();
            int count = (int)db.QueryDefs.Count;
            for (int i = 0; i < count; i++)
            {
                dynamic qd = db.QueryDefs(i);
                string name = (string)qd.Name;
                if (name.StartsWith("~")) continue; // 一時クエリ
                string sql = "";
                try { sql = (string)qd.SQL; } catch { }
                result.Add(new QueryInfo(name, sql));
            }
            return result;
        }
        finally
        {
            db.Close();
        }
    }

    public IReadOnlyList<string[]> QuerySql(string dbPath, string sql)
    {
        dbPath = ResolveAbsPath(dbPath);
        dynamic db = OpenDaoDatabase(dbPath);
        try
        {
            // ACE SQL パーサーは Unicode 識別子を正しく扱えないため、
            // 単純な "SELECT * FROM [table]" の場合は TableDef をインデックスで検索して Recordset を開く
            dynamic rs;
            string? tableOnly = TryExtractTableName(sql);
            if (tableOnly is not null)
            {
                dynamic? foundTd = FindTableDefByName(db, tableOnly);
                if (foundTd is null)
                    throw new InvalidOperationException($"テーブルが見つかりません: {tableOnly}");
                rs = foundTd.OpenRecordset(4); // 4 = dbOpenSnapshot
            }
            else
            {
                rs = db.OpenRecordset(sql, 4);
            }

            var results = new List<string[]>();
            int fieldCount = (int)rs.Fields.Count;

            var header = new string[fieldCount];
            for (int i = 0; i < fieldCount; i++)
                header[i] = (string)rs.Fields(i).Name;
            results.Add(header);

            while (!(bool)rs.EOF)
            {
                var row = new string[fieldCount];
                for (int i = 0; i < fieldCount; i++)
                {
                    object val = rs.Fields(i).Value;
                    row[i] = val is null || val is DBNull ? "" : val.ToString()!;
                }
                results.Add(row);
                rs.MoveNext();
            }
            rs.Close();
            return results;
        }
        finally
        {
            db.Close();
        }
    }

    /// <summary>
    /// TableDefs コレクションをインデックスでイテレートして .NET 側で名前比較する
    /// （COM 引数に日本語文字列を渡すと失敗するため）
    /// </summary>
    private static dynamic? FindTableDefByName(dynamic db, string name)
    {
        int count = (int)db.TableDefs.Count;
        for (int i = 0; i < count; i++)
        {
            dynamic td = db.TableDefs(i);
            if ((string)td.Name == name)
                return td;
        }
        return null;
    }

    /// <summary>
    /// "SELECT * FROM [table]" または "SELECT * FROM table" の形式のとき
    /// テーブル名だけを返す（WHERE句などがある場合は null）
    /// </summary>
    private static string? TryExtractTableName(string sql)
    {
        // 正規化: 前後の空白と改行を除去
        var s = sql.Trim();
        // "SELECT" で始まり "FROM" を含む場合のみ
        var m = System.Text.RegularExpressions.Regex.Match(
            s,
            @"^\s*SELECT\s+\*\s+FROM\s+\[?([^\]\s]+)\]?\s*$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return m.Success ? m.Groups[1].Value : null;
    }

    public void ExecSql(string dbPath, string sql)
    {
        dbPath = ResolveAbsPath(dbPath);

        // INSERT の場合は DAO AddNew/Update で処理（ACE SQL パーサーの Unicode 問題を回避）
        var insert = TryParseInsert(sql);
        if (insert is not null)
        {
            ExecInsert(dbPath, insert.Value.table, insert.Value.columns, insert.Value.values);
            return;
        }

        // その他（UPDATE/DELETE、英語テーブル名）は直接実行
        dynamic db = OpenDaoDatabase(dbPath);
        try
        {
            db.Execute(sql, 128); // 128 = dbFailOnError
        }
        finally
        {
            db.Close();
        }
    }

    private void ExecInsert(string dbPath, string tableName, string[] columns, string[] values)
    {
        dynamic db = OpenDaoDatabase(dbPath);
        try
        {
            dynamic? td = FindTableDefByName(db, tableName);
            if (td is null)
                throw new InvalidOperationException($"テーブルが見つかりません: {tableName}");

            dynamic rs = td.OpenRecordset(2); // 2 = dbOpenDynaset

            // フィールド名→インデックスのマップを .NET 側で構築
            int fcount = (int)rs.Fields.Count;
            var fieldMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < fcount; i++)
                fieldMap[(string)rs.Fields(i).Name] = i;

            rs.AddNew();
            for (int i = 0; i < columns.Length; i++)
            {
                if (fieldMap.TryGetValue(columns[i], out int fi))
                    rs.Fields(fi).Value = ParseSqlLiteral(values[i]);
            }
            rs.Update();
            rs.Close();
        }
        finally
        {
            db.Close();
        }
    }

    private static (string table, string[] columns, string[] values)? TryParseInsert(string sql)
    {
        var m = System.Text.RegularExpressions.Regex.Match(
            sql.Trim(),
            @"^\s*INSERT\s+INTO\s+\[?([^\]\s,]+)\]?\s*\(([^)]+)\)\s*VALUES\s*\(([^)]+)\)\s*;?\s*$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);
        if (!m.Success) return null;

        string table = m.Groups[1].Value.Trim();
        string[] cols = SplitSqlList(m.Groups[2].Value);
        string[] vals = SplitSqlList(m.Groups[3].Value);
        // [ ] を除去
        cols = cols.Select(c => c.Trim('[', ']', ' ')).ToArray();
        return (table, cols, vals);
    }

    private static string[] SplitSqlList(string s)
    {
        var items = new List<string>();
        int depth = 0;
        bool inQuote = false;
        int start = 0;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c == '\'' && !inQuote) inQuote = true;
            else if (c == '\'' && inQuote) inQuote = false;
            else if (c == ',' && !inQuote && depth == 0)
            {
                items.Add(s[start..i].Trim());
                start = i + 1;
            }
        }
        items.Add(s[start..].Trim());
        return items.ToArray();
    }

    private static object ParseSqlLiteral(string s)
    {
        s = s.Trim();
        if (s.StartsWith('\'') && s.EndsWith('\''))
            return s[1..^1].Replace("''", "'");
        if (long.TryParse(s, out long l)) return l;
        if (double.TryParse(s, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out double d)) return d;
        if (s.Equals("NULL", StringComparison.OrdinalIgnoreCase)) return DBNull.Value;
        return s;
    }

    public string GetQuerySql(string dbPath, string queryName)
    {
        dbPath = ResolveAbsPath(dbPath);
        dynamic db = OpenDaoDatabase(dbPath);
        try
        {
            int count = (int)db.QueryDefs.Count;
            for (int i = 0; i < count; i++)
            {
                dynamic qd = db.QueryDefs(i);
                string name = (string)qd.Name;
                if (name == queryName || name.Contains(queryName))
                    return (string)qd.SQL;
            }
            throw new InvalidOperationException($"クエリが見つかりません: {queryName}");
        }
        finally
        {
            db.Close();
        }
    }

    public void ExportAllQueries(string dbPath, string outputDir)
    {
        dbPath = ResolveAbsPath(dbPath);
        Directory.CreateDirectory(outputDir);
        dynamic db = OpenDaoDatabase(dbPath);
        try
        {
            int count = (int)db.QueryDefs.Count;
            for (int i = 0; i < count; i++)
            {
                dynamic qd = db.QueryDefs(i);
                string name = (string)qd.Name;
                if (name.StartsWith("~")) continue;
                string sql = "";
                try { sql = (string)qd.SQL; } catch { }
                string safeName = string.Join("_", name.Split(Path.GetInvalidFileNameChars()));
                File.WriteAllText(Path.Combine(outputDir, safeName + ".sql"), sql, System.Text.Encoding.UTF8);
                Console.WriteLine(name);
            }
        }
        finally
        {
            db.Close();
        }
    }

    public void SetQuerySql(string dbPath, string queryName, string sql)
    {
        dbPath = ResolveAbsPath(dbPath);
        dynamic db = OpenDaoDatabase(dbPath);
        try
        {
            int count = (int)db.QueryDefs.Count;
            for (int i = 0; i < count; i++)
            {
                dynamic qd = db.QueryDefs(i);
                if ((string)qd.Name == queryName)
                {
                    qd.SQL = sql;
                    return;
                }
            }
            throw new InvalidOperationException($"クエリが見つかりません: {queryName}");
        }
        finally
        {
            db.Close();
        }
    }

    // ─── Helpers ────────────────────────────────────────────────────

    private static string ResolveAbsPath(string path)
    {
        path = Path.GetFullPath(path);
        if (!File.Exists(path))
            throw new FileNotFoundException($"ファイルが見つかりません: {path}");
        return path;
    }

    private static dynamic OpenDaoDatabase(string dbPath)
    {
        Type? t = Type.GetTypeFromProgID("DAO.DBEngine.120")
            ?? throw new InvalidOperationException("DAO.DBEngine.120 が見つかりません。Access Database Engine をインストールしてください");
        dynamic engine = Activator.CreateInstance(t)
            ?? throw new InvalidOperationException("DAO.DBEngine.120 を起動できませんでした");
        return engine.Workspaces(0).OpenDatabase(dbPath, false, false, "");
    }

    private static dynamic CreateAccessApp()
    {
        Type? t = Type.GetTypeFromProgID("Access.Application")
            ?? throw new InvalidOperationException("Microsoft Accessがインストールされていません");
        return Activator.CreateInstance(t)
            ?? throw new InvalidOperationException("Access.Applicationを起動できませんでした");
    }

    private static void SafeQuit(dynamic access)
    {
        try { access.CloseCurrentDatabase(); } catch { }
        try { access.Quit(); } catch { }
        try { Marshal.ReleaseComObject(access); } catch { }
    }

    private static byte[] BuildPattern(byte[] strBytes)
    {
        byte[] len = new byte[4];
        BinaryPrimitives.WriteUInt32LittleEndian(len, (uint)strBytes.Length);
        return [.. len, .. strBytes];
    }

    private static int IndexOf(byte[] haystack, byte[] needle, int start = 0)
    {
        int limit = haystack.Length - needle.Length;
        for (int i = start; i <= limit; i++)
        {
            bool match = true;
            for (int j = 0; j < needle.Length; j++)
            {
                if (haystack[i + j] != needle[j]) { match = false; break; }
            }
            if (match) return i;
        }
        return -1;
    }

    private static byte[] ReplaceAll(byte[] data, byte[] oldPat, byte[] newPat)
    {
        // Count occurrences first for pre-sizing
        int count = 0;
        int idx = 0;
        while ((idx = IndexOf(data, oldPat, idx)) >= 0) { count++; idx++; }

        int newSize = data.Length + count * (newPat.Length - oldPat.Length);
        byte[] result = new byte[newSize];
        int src = 0, dst = 0;

        while (src < data.Length)
        {
            int found = IndexOf(data, oldPat, src);
            if (found < 0)
            {
                // Copy remainder
                int rem = data.Length - src;
                Array.Copy(data, src, result, dst, rem);
                dst += rem;
                break;
            }
            // Copy up to match
            int before = found - src;
            Array.Copy(data, src, result, dst, before);
            dst += before;
            // Copy replacement
            Array.Copy(newPat, 0, result, dst, newPat.Length);
            dst += newPat.Length;
            src = found + oldPat.Length;
        }
        return result;
    }
}
