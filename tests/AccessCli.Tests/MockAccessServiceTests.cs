using AccessCli;

namespace AccessCli.Tests;

/// <summary>
/// IAccessService インターフェースのモック実装を使ったテスト例
/// </summary>
public class MockAccessServiceTests
{
    [Fact]
    public void ListModules_ReturnsCorrectModuleInfo()
    {
        var svc = new MockAccessService();
        svc.AddModule("Module1", "Module", 42);
        svc.AddModule("Form_Main", "Form/Report", 10);

        var modules = svc.ListModules("dummy.accdb");

        Assert.Equal(2, modules.Count);
        Assert.Equal("Module1", modules[0].Name);
        Assert.Equal("Module", modules[0].Type);
        Assert.Equal(42, modules[0].Lines);
    }

    [Fact]
    public void ReadVba_ReturnsStoredCode()
    {
        var svc = new MockAccessService();
        svc.SetVbaCode("Module1", "Sub Hello()\nEnd Sub");

        string code = svc.ReadVba("dummy.accdb", "Module1");

        Assert.Equal("Sub Hello()\nEnd Sub", code);
    }

    [Fact]
    public void WriteVba_OverwritesCode()
    {
        var svc = new MockAccessService();
        svc.SetVbaCode("Module1", "Old Code");

        svc.WriteVba("dummy.accdb", "Module1", "New Code");

        Assert.Equal("New Code", svc.ReadVba("dummy.accdb", "Module1"));
    }

    [Fact]
    public void ListForms_ReturnsFormNames()
    {
        var svc = new MockAccessService();
        svc.AddForm("メイン画面");
        svc.AddForm("設定画面");

        var forms = svc.ListForms("dummy.accdb");

        Assert.Equal(2, forms.Count);
        Assert.Contains("メイン画面", forms);
    }
}

/// <summary>テスト用のインメモリモック実装</summary>
public class MockAccessService : IAccessService
{
    private readonly List<ModuleInfo> _modules = [];
    private readonly Dictionary<string, string> _vbaCode = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<string> _forms = [];
    private readonly Dictionary<string, List<ControlInfo>> _controls = new(StringComparer.OrdinalIgnoreCase);

    public void AddModule(string name, string type, int lines) =>
        _modules.Add(new ModuleInfo(name, type, lines));

    public void SetVbaCode(string module, string code) =>
        _vbaCode[module] = code;

    public void AddForm(string name) => _forms.Add(name);

    public void AddControl(string form, ControlInfo ctl) =>
        (_controls.TryGetValue(form, out var list) ? list : _controls[form] = []).Add(ctl);

    public IReadOnlyList<ModuleInfo> ListModules(string dbPath) => _modules;

    public string ReadVba(string dbPath, string moduleName) =>
        _vbaCode.TryGetValue(moduleName, out var c) ? c : string.Empty;

    public void WriteVba(string dbPath, string moduleName, string code) =>
        _vbaCode[moduleName] = code;

    public IReadOnlyList<string> ListForms(string dbPath) => _forms;

    public IReadOnlyList<ControlInfo> ListControls(string dbPath, string formName) =>
        _controls.TryGetValue(formName, out var list) ? list : [];

    public void ExportForm(string dbPath, string formName, string outputPath) =>
        File.WriteAllText(outputPath, $"[Form: {formName}]");

    public void ImportForm(string dbPath, string formName, string inputPath) { }

    public int SetCaption(string dbPath, string oldCaption, string newCaption) => 0;
}
