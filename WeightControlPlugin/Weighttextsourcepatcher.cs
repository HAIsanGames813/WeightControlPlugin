using HarmonyLib;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace WeightControlPlugin;

/// <summary>
/// TextSource.Update() を Harmony でパッチし、
/// TextItem.Text に含まれる制御タグを一時的に除去してから描画させます。
///
/// ISource / ISourceOutput は internal なので直接参照できません。
/// TextSource インスタンスが保持する TextItem (item フィールド) を
/// リフレクションで取得し、Text プロパティを一時差し替えする方式を使います。
/// </summary>
internal static class WeightTextSourcePatcher
{
    private const string HarmonyId = "com.weightcontrolplugin.textsource";
    private static Harmony? _harmony;
    private static readonly string LogPath =
        Path.Combine(Path.GetTempPath(), "WeightControlPlugin_log.txt");

    // TextSource.item フィールド (初回パッチ時にキャッシュ)
    private static FieldInfo? _itemField;
    // TextItem.Text プロパティ (初回パッチ時にキャッシュ)
    private static PropertyInfo? _textProp;

    public static void Apply()
    {
        Log("WeightTextSourcePatcher.Apply() 開始");

        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            TryPatch(asm);

        AppDomain.CurrentDomain.AssemblyLoad += (_, args) => TryPatch(args.LoadedAssembly);
    }

    private static void TryPatch(Assembly asm)
    {
        if (_harmony != null) return;

        // YukkuriMovieMaker.dll (本体) のみ対象
        var name = asm.GetName().Name ?? "";
        if (name != "YukkuriMovieMaker") return;

        // TextSource 型を探す (internal)
        Type? textSourceType = null;
        try
        {
            textSourceType = asm.GetType("YukkuriMovieMaker.Player.Video.Items.TextSource");
            if (textSourceType == null)
                textSourceType = asm.GetTypes().FirstOrDefault(t => t.Name == "TextSource");
        }
        catch (Exception ex)
        {
            Log($"TextSource 検索例外: {ex.Message}");
            return;
        }

        if (textSourceType == null)
        {
            Log("TextSource が見つかりません");
            return;
        }

        Log($"TextSource 発見: {textSourceType.FullName}");

        // TextItem 型も探す
        Type? textItemType = null;
        try
        {
            textItemType = asm.GetType("YukkuriMovieMaker.Project.Items.TextItem")
                ?? asm.GetTypes().FirstOrDefault(t => t.Name == "TextItem");
        }
        catch { }

        // フィールド・プロパティをキャッシュ
        _itemField = textSourceType.GetField("item",
            BindingFlags.Instance | BindingFlags.NonPublic);

        if (_itemField == null)
        {
            Log("item フィールドが見つかりません。全フィールドをダンプ:");
            foreach (var f in textSourceType.GetFields(
                BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public))
                Log($"  [{f.FieldType.Name}] {f.Name}");
        }
        else
        {
            Log($"item フィールド: {_itemField.FieldType.Name}");
            _textProp = _itemField.FieldType.GetProperty("Text",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            Log($"Text プロパティ: {_textProp?.Name ?? "見つからない"}");
        }

        // Update(TimelineItemSourceDescription) をパッチ
        var updateMethod = textSourceType.GetMethod("Update",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

        if (updateMethod == null)
        {
            Log("Update が見つかりません");
            return;
        }

        Log($"パッチ対象: {textSourceType.Name}.{updateMethod.Name}");
        _harmony = new Harmony(HarmonyId);
        _harmony.Patch(updateMethod,
            prefix: new HarmonyMethod(typeof(WeightTextSourcePatcher), nameof(Prefix_Update)),
            postfix: new HarmonyMethod(typeof(WeightTextSourcePatcher), nameof(Postfix_Update)));
        Log("TextSource.Update パッチ完了");
    }

    // ── Prefix: Update() 前にテキストをクリーンに差し替え ────────

    [HarmonyPrefix]
    private static void Prefix_Update(object __instance, out string? __state)
    {
        __state = null;
        if (_itemField == null || _textProp == null) return;

        try
        {
            var item = _itemField.GetValue(__instance);
            if (item == null) return;

            var original = _textProp.GetValue(item) as string ?? string.Empty;

            if (!WeightTagParser.HasControlTags(original)) return;

            // 元テキストを __state に保存してからクリーンテキストを設定
            __state = original;
            var (clean, _) = WeightTagParser.Parse(original);
            _textProp.SetValue(item, clean);
        }
        catch (Exception ex)
        {
            Log($"Prefix_Update 例外: {ex.Message}");
        }
    }

    // ── Postfix: Update() 後に元テキストを復元 ───────────────────

    [HarmonyPostfix]
    private static void Postfix_Update(object __instance, string? __state)
    {
        if (__state == null || _itemField == null || _textProp == null) return;

        try
        {
            var item = _itemField.GetValue(__instance);
            if (item != null)
                _textProp.SetValue(item, __state);
        }
        catch (Exception ex)
        {
            Log($"Postfix_Update 例外: {ex.Message}");
        }
    }

    private static void Log(string msg)
    {
        System.Diagnostics.Debug.WriteLine($"[WeightControlPlugin] {msg}");
        try { File.AppendAllText(LogPath, $"[{DateTime.Now:HH:mm:ss.fff}] {msg}\n"); }
        catch { }
    }
}

// ─────────────────────────────────────────────────────────────────
// 制御タグパーサー
// ─────────────────────────────────────────────────────────────────

/// <summary>
/// AviUtl 互換制御タグのパーサー
///   &lt;w秒数&gt;   … ウェイト
///   &lt;w*係数&gt;  … 文字数×係数秒
///   &lt;c&gt;       … クリア
///   &lt;p,x,y&gt;   … 座標指定
/// </summary>
internal static class WeightTagParser
{
    private static readonly Regex TagRegex = new Regex(
        @"<w\*(?<coeff>[\d.]+)>|<w(?<sec>[\d.]+)>|<c>|<p,(?<x>[\d.-]+),(?<y>[\d.-]+)>",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static bool HasControlTags(string text) =>
        !string.IsNullOrEmpty(text) && TagRegex.IsMatch(text);

    public static (string cleanText, List<WeightTag> tags) Parse(string text)
    {
        var tags = new List<WeightTag>();
        int offset = 0; // タグ除去による文字位置のズレ

        var clean = TagRegex.Replace(text, match =>
        {
            // クリーンテキスト上での位置 = マッチ前の文字数 - これまでのタグ除去分
            int cleanPos = match.Index - offset;
            offset += match.Length;

            if (match.Groups["coeff"].Success)
            {
                double coeff = double.Parse(match.Groups["coeff"].Value,
                    System.Globalization.CultureInfo.InvariantCulture);
                tags.Add(new WeightTag(cleanPos, WeightTagType.WaitCoeff, coeff, 0, 0));
            }
            else if (match.Groups["sec"].Success)
            {
                double sec = double.Parse(match.Groups["sec"].Value,
                    System.Globalization.CultureInfo.InvariantCulture);
                tags.Add(new WeightTag(cleanPos, WeightTagType.Wait, sec, 0, 0));
            }
            else if (match.Value.Equals("<c>", StringComparison.OrdinalIgnoreCase))
            {
                tags.Add(new WeightTag(cleanPos, WeightTagType.Clear, 0, 0, 0));
            }
            else if (match.Groups["x"].Success)
            {
                double x = double.Parse(match.Groups["x"].Value,
                    System.Globalization.CultureInfo.InvariantCulture);
                double y = double.Parse(match.Groups["y"].Value,
                    System.Globalization.CultureInfo.InvariantCulture);
                tags.Add(new WeightTag(cleanPos, WeightTagType.Position, 0, x, y));
            }
            return string.Empty;
        });

        return (clean, tags);
    }
}

internal enum WeightTagType { Wait, WaitCoeff, Clear, Position }

internal record WeightTag(
    int Position,
    WeightTagType Type,
    double Value,
    double X,
    double Y
);