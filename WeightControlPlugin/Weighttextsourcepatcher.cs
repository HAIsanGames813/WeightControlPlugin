using HarmonyLib;
using System.Reflection;
using System.Text.RegularExpressions;

namespace WeightControlPlugin;

/// <summary>
/// TextSource.Update() をパッチし、TextItem.Text に含まれる制御タグを
/// 一時的に除去してから描画させます。
/// </summary>
internal static class WeightTextSourcePatcher
{
    private const string HarmonyId = "com.weightcontrolplugin.textsource";
    private static Harmony? _harmony;

    private static FieldInfo? _itemField;
    private static PropertyInfo? _textProp;

    public static void Apply()
    {
        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            TryPatch(asm);
        AppDomain.CurrentDomain.AssemblyLoad += (_, args) => TryPatch(args.LoadedAssembly);
    }

    private static void TryPatch(Assembly asm)
    {
        if (_harmony != null) return;
        if (asm.GetName().Name != "YukkuriMovieMaker") return;

        Type? textSourceType = null;
        try
        {
            textSourceType =
                asm.GetType("YukkuriMovieMaker.Player.Video.Items.TextSource")
                ?? asm.GetTypes().FirstOrDefault(t => t.Name == "TextSource");
        }
        catch { return; }

        if (textSourceType == null) return;

        _itemField = textSourceType.GetField("item",
            BindingFlags.Instance | BindingFlags.NonPublic);
        if (_itemField != null)
            _textProp = _itemField.FieldType.GetProperty("Text",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

        var updateMethod = textSourceType.GetMethod("Update",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
        if (updateMethod == null) return;

        _harmony = new Harmony(HarmonyId);
        _harmony.Patch(updateMethod,
            prefix: new HarmonyMethod(typeof(WeightTextSourcePatcher), nameof(Prefix_Update)),
            postfix: new HarmonyMethod(typeof(WeightTextSourcePatcher), nameof(Postfix_Update)));
    }

    [HarmonyPrefix]
    private static void Prefix_Update(object __instance, out string? __state)
    {
        __state = null;
        if (_itemField == null || _textProp == null) return;

        var item = _itemField.GetValue(__instance);
        if (item == null) return;

        var original = _textProp.GetValue(item) as string ?? string.Empty;
        if (!WeightTagParser.HasControlTags(original)) return;

        __state = original;
        _textProp.SetValue(item, WeightTagParser.RemoveTags(original));
    }

    [HarmonyPostfix]
    private static void Postfix_Update(object __instance, string? __state)
    {
        if (__state == null || _itemField == null || _textProp == null) return;
        var item = _itemField.GetValue(__instance);
        if (item != null) _textProp.SetValue(item, __state);
    }
}

/// <summary>
/// 制御タグのパーサー
///   &lt;w秒数&gt;  : ウェイト
///   &lt;w*係数&gt; : 文字数×係数秒
///   &lt;c&gt;      : クリア
///   &lt;p,x,y&gt;  : 座標指定
/// </summary>
internal static class WeightTagParser
{
    private static readonly Regex TagRegex = new(
        @"<w\*[\d.]+>|<w[\d.]+>|<c>|<p,[\d.-]+,[\d.-]+>",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static bool HasControlTags(string text) =>
        !string.IsNullOrEmpty(text) && TagRegex.IsMatch(text);

    public static string RemoveTags(string text) =>
        TagRegex.Replace(text, string.Empty);
}