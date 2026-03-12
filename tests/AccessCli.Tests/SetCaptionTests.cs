using System.Text;
using AccessCli;

namespace AccessCli.Tests;

/// <summary>
/// SetCaption のバイナリ置換ロジックをファイルなしでテスト
/// </summary>
public class SetCaptionTests
{
    private static byte[] BuildTestFile(params string[] captions)
    {
        // accdbに現れるパターンを模倣: [4-byte LE length][UTF-16-LE string]
        var buf = new List<byte>();
        buf.AddRange([0xAA, 0xBB]); // 前置ごみデータ
        foreach (var cap in captions)
        {
            byte[] strBytes = Encoding.Unicode.GetBytes(cap);
            buf.Add((byte)(strBytes.Length & 0xFF));
            buf.Add((byte)((strBytes.Length >> 8) & 0xFF));
            buf.Add(0x00);
            buf.Add(0x00);
            buf.AddRange(strBytes);
        }
        buf.AddRange([0xCC, 0xDD]); // 後置ごみデータ
        return [.. buf];
    }

    [Fact]
    public void SetCaption_SameLength_ReplacesCorrectly()
    {
        var svc = new AccessService();
        string tmpFile = Path.GetTempFileName();
        try
        {
            File.WriteAllBytes(tmpFile, BuildTestFile("旧ラベル"));
            int count = svc.SetCaption(tmpFile, "旧ラベル", "新ラベル");
            Assert.Equal(1, count);

            byte[] result = File.ReadAllBytes(tmpFile);
            byte[] expected = BuildTestFile("新ラベル");
            Assert.Equal(expected, result);
        }
        finally { File.Delete(tmpFile); }
    }

    [Fact]
    public void SetCaption_MultipleOccurrences_ReplacesAll()
    {
        var svc = new AccessService();
        string tmpFile = Path.GetTempFileName();
        try
        {
            File.WriteAllBytes(tmpFile, BuildTestFile("ABC", "ABC", "XYZ"));
            int count = svc.SetCaption(tmpFile, "ABC", "DEF");
            Assert.Equal(2, count);

            byte[] result = File.ReadAllBytes(tmpFile);
            byte[] expected = BuildTestFile("DEF", "DEF", "XYZ");
            Assert.Equal(expected, result);
        }
        finally { File.Delete(tmpFile); }
    }

    [Fact]
    public void SetCaption_NotFound_ReturnsZero()
    {
        var svc = new AccessService();
        string tmpFile = Path.GetTempFileName();
        try
        {
            byte[] original = BuildTestFile("存在しない");
            File.WriteAllBytes(tmpFile, original);
            int count = svc.SetCaption(tmpFile, "探す文字列", "置換後");
            Assert.Equal(0, count);

            // ファイルは変更されていない
            Assert.Equal(original, File.ReadAllBytes(tmpFile));
        }
        finally { File.Delete(tmpFile); }
    }

    [Fact]
    public void SetCaption_DifferentByteLength_ReplacesWithNewPattern()
    {
        // バイト長が異なる場合でも置換は実行される（ただしaccdbとして壊れる可能性あり）
        var svc = new AccessService();
        string tmpFile = Path.GetTempFileName();
        try
        {
            File.WriteAllBytes(tmpFile, BuildTestFile("短い"));
            int count = svc.SetCaption(tmpFile, "短い", "より長い文字列");
            Assert.Equal(1, count);

            byte[] result = File.ReadAllBytes(tmpFile);
            byte[] expected = BuildTestFile("より長い文字列");
            Assert.Equal(expected, result);
        }
        finally { File.Delete(tmpFile); }
    }
}
