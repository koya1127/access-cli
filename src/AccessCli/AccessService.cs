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

    // ─── Helpers ────────────────────────────────────────────────────

    private static string ResolveAbsPath(string path)
    {
        path = Path.GetFullPath(path);
        if (!File.Exists(path))
            throw new FileNotFoundException($"ファイルが見つかりません: {path}");
        return path;
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
