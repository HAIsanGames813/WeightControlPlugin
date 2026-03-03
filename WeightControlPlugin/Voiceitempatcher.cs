using HarmonyLib;
using System.Reflection;

namespace WeightControlPlugin;

/// <summary>
/// VoiceItem に関する Harmony パッチ。
///
/// 目的:
///   セリフ（Serif）に含まれる制御タグを音声合成に渡さないようにする。
///
/// アプローチ:
///   1. SerifToHatsuonAsync prefix: serif フィールドからタグを一時除去して
///      KanjiToYomi 変換入力を clean にする。postfix で元の値を復元。
///   2. Hatsuon setter postfix: pass-through 型エンジン向けの保険的除去。
/// </summary>
internal static class VoiceItemPatcher
{
    private const string HarmonyId = "com.weightcontrolplugin.voiceitem";
    private static Harmony? _harmony;

    private static FieldInfo? _serifField;
    private static FieldInfo? _hatsuonField;
    private static MethodInfo? _serifToHatsuonAsync;
    private static MethodInfo? _hatsuonSetter;

    // SerifToHatsuonAsync は async メソッドなので
    // prefix で退避した値を postfix で復元するため ThreadStatic を使う
    [ThreadStatic]
    private static string? _savedSerif;

    public static void Apply()
    {
        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            TryPatch(asm);
        AppDomain.CurrentDomain.AssemblyLoad += (_, args) => TryPatch(args.LoadedAssembly);
    }

    private static void TryPatch(Assembly asm)
    {
        if (_harmony != null) return;

        Type? voiceItemType = null;
        try
        {
            voiceItemType = asm.GetTypes()
                .FirstOrDefault(t => t.Name == "VoiceItem" && !t.IsInterface && !t.IsAbstract);
        }
        catch (ReflectionTypeLoadException ex)
        {
            voiceItemType = ex.Types.Where(t => t != null).Select(t => t!)
                .FirstOrDefault(t => t.Name == "VoiceItem" && !t.IsInterface && !t.IsAbstract);
        }
        catch { return; }

        if (voiceItemType == null) return;

        _serifField = voiceItemType.GetField("serif", BindingFlags.Instance | BindingFlags.NonPublic);
        _hatsuonField = voiceItemType.GetField("hatsuon", BindingFlags.Instance | BindingFlags.NonPublic);
        if (_serifField == null || _hatsuonField == null) return;

        _serifToHatsuonAsync = voiceItemType.GetMethod("SerifToHatsuonAsync",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
        _hatsuonSetter = voiceItemType
            .GetProperty("Hatsuon", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
            ?.GetSetMethod(nonPublic: true);

        _harmony = new Harmony(HarmonyId);

        if (_serifToHatsuonAsync != null)
        {
            _harmony.Patch(_serifToHatsuonAsync,
                prefix: new HarmonyMethod(typeof(VoiceItemPatcher), nameof(Prefix_SerifToHatsuon)),
                postfix: new HarmonyMethod(typeof(VoiceItemPatcher), nameof(Postfix_SerifToHatsuon)));
        }
        if (_hatsuonSetter != null)
        {
            _harmony.Patch(_hatsuonSetter,
                postfix: new HarmonyMethod(typeof(VoiceItemPatcher), nameof(Postfix_HatsuonSetter)));
        }
    }

    /// <summary>
    /// SerifToHatsuonAsync の prefix。
    /// async メソッドの同期プロローグ（Serif?.Normalize(...)）実行前に呼ばれる。
    /// serif フィールドをタグ除去済みに差し替えて KanjiToYomi への入力を clean にする。
    /// </summary>
    [HarmonyPrefix]
    private static void Prefix_SerifToHatsuon(object __instance)
    {
        if (_serifField == null) return;
        try
        {
            var current = _serifField.GetValue(__instance) as string;
            if (current == null || !WeightTagParser.HasControlTags(current)) return;
            _savedSerif = current;
            _serifField.SetValue(__instance, WeightTagParser.RemoveTags(current));
        }
        catch { }
    }

    /// <summary>
    /// SerifToHatsuonAsync の postfix。
    /// async メソッドが Task を返した直後（ローカル変数 normalizedSerif 確定後）に実行される。
    /// prefix で除去したタグ付き serif を復元する。
    /// </summary>
    [HarmonyPostfix]
    private static void Postfix_SerifToHatsuon(object __instance)
    {
        if (_serifField == null || _savedSerif == null) return;
        try { _serifField.SetValue(__instance, _savedSerif); }
        catch { }
        finally { _savedSerif = null; }
    }

    /// <summary>
    /// Hatsuon setter の postfix。
    /// pass-through 型エンジン（AquesTalk Custom 等）向けの保険。
    /// 変換結果にタグが残っていた場合に除去する。
    /// </summary>
    [HarmonyPostfix]
    private static void Postfix_HatsuonSetter(object __instance)
    {
        if (_hatsuonField == null) return;
        try
        {
            var current = _hatsuonField.GetValue(__instance) as string;
            if (current == null || !WeightTagParser.HasControlTags(current)) return;
            _hatsuonField.SetValue(__instance, WeightTagParser.RemoveTags(current));
        }
        catch { }
    }
}