using System.Runtime.CompilerServices;
using YukkuriMovieMaker.Plugin;

namespace WeightControlPlugin;

public class WeightControlPlugin : IPlugin
{
    public string Name => "テキスト制御コード挿入プラグイン";

    public void Initialize() => Startup.Run();
    public void Dispose() { }

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

        // 各パッチャーを個別に try-catch: 1つ失敗しても他は続行
        try { RichTextEditorPatcher.Apply(); } catch { }  // TextItem テキスト入力UI
        try { SerifTextBoxPatcher.Apply(); } catch { }  // VoiceItem セリフ入力UI
        try { VoiceItemPatcher.Apply(); } catch { }  // Hatsuon から制御タグを除去
        try { WeightTextSourcePatcher.Apply(); } catch { }  // 映像レンダラーパッチ
    }
}