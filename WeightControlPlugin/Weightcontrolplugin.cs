using System.IO;
using System.Runtime.CompilerServices;
using YukkuriMovieMaker.Plugin;

namespace WeightControlPlugin;

/// <summary>
/// プラグインエントリポイント。
/// [ModuleInitializer] により DLL ロード時に自動実行されます。
/// </summary>
public class WeightControlPlugin : IPlugin
{
    public string Name => "テキスト制御コード挿入プラグイン";

    public void Initialize() => Startup.Run();
    public void Dispose() { }

    /// <summary>
    /// DLL ロード直後に自動実行。
    /// YMM4 がどんな方式でプラグインを読み込んでも確実に初期化されます。
    /// CA2255 は「ライブラリでの ModuleInitializer 使用」の警告ですが、
    /// ここではプラグイン初期化目的なので意図的に使用しています。
    /// </summary>
#pragma warning disable CA2255
    [ModuleInitializer]
    internal static void OnModuleLoad() => Startup.Run();
#pragma warning restore CA2255
}

internal static class Startup
{
    private static int _ran;

    public static void Run()
    {
        if (System.Threading.Interlocked.Exchange(ref _ran, 1) != 0) return;

        var log = Path.Combine(Path.GetTempPath(), "WeightControlPlugin_log.txt");
        File.AppendAllText(log, $"[{DateTime.Now:HH:mm:ss.fff}] Startup.Run() 開始\n");
        System.Diagnostics.Debug.WriteLine("[WeightControlPlugin] Startup.Run()");

        // ① RichTextEditor コンテキストメニュー注入
        try
        {
            RichTextEditorPatcher.Apply();
            File.AppendAllText(log,
                $"[{DateTime.Now:HH:mm:ss.fff}] RichTextEditorPatcher 完了\n");
        }
        catch (Exception ex)
        {
            File.AppendAllText(log,
                $"[{DateTime.Now:HH:mm:ss.fff}] RichTextEditorPatcher 例外: {ex}\n");
        }

        // ② TextSource パッチ (<w><c><p> タグ処理)
        try
        {
            WeightTextSourcePatcher.Apply();
            File.AppendAllText(log,
                $"[{DateTime.Now:HH:mm:ss.fff}] WeightTextSourcePatcher 完了\n");
        }
        catch (Exception ex)
        {
            File.AppendAllText(log,
                $"[{DateTime.Now:HH:mm:ss.fff}] WeightTextSourcePatcher 例外: {ex}\n");
        }
    }
}