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
        RichTextEditorPatcher.Apply();
        WeightTextSourcePatcher.Apply();
    }
}