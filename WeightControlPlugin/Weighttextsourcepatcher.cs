using HarmonyLib;
using System.Collections;
using System.Drawing;
using System.Reflection;
using System.Text.RegularExpressions;

namespace WeightControlPlugin;

internal static class WeightTextSourcePatcher
{
    private const string HarmonyId = "com.weightcontrolplugin.textsource";
    private static Harmony? _harmony;

    private static Type? _textItemType;
    private static PropertyInfo? _tiTextProp;

    private static FieldInfo? _tsItemField;
    private static FieldInfo? _tsOutputsListField;
    private static FieldInfo? _tsIsDividedField;
    private static FieldInfo? _tsDisplayIntField;
    private static FieldInfo? _tsHideIntField;
    private static PropertyInfo? _tsOutputsProp;

    private static FieldInfo? _soTimeOffsetField;
    private static FieldInfo? _soDurationOffsetField;
    private static FieldInfo? _soDrawingOffsetField;

    private const double ForcedInterval = 0.001;

    // TextItem ごとの「全文字の初期自然位置キャッシュ」
    // TextSource が毎フレーム出力数が変わっても位置が不変になるよう保持する
    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<object, NaturalPosCache>
        _posCache = new();

    private sealed class NaturalPosCache
    {
        public PointF[] Positions = Array.Empty<PointF>();
        public int Count;
        public string Text = string.Empty;
    }

    // -----------------------------------------------------------------

    public static void Apply()
    {
        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            TryPatch(asm);
        AppDomain.CurrentDomain.AssemblyLoad += (_, args) => TryPatch(args.LoadedAssembly);
    }

    private static void TryPatch(Assembly asm)
    {
        if (_harmony != null) return;
        if (!(asm.GetName().Name?.StartsWith("YukkuriMovieMaker") ?? false)) return;

        Type[]? types = null;
        try { types = asm.GetTypes(); }
        catch (ReflectionTypeLoadException ex) { types = ex.Types.Where(t => t != null).ToArray()!; }
        catch { return; }
        if (types == null) return;

        _textItemType = types.FirstOrDefault(t => t.FullName == "YukkuriMovieMaker.Project.Items.TextItem");
        var textSourceType = types.FirstOrDefault(t => t.FullName == "YukkuriMovieMaker.Player.Video.Items.TextSource");
        var layoutDescType = types.FirstOrDefault(t => t.Name == "TextLayoutDescription");

        _tiTextProp = _textItemType?.GetProperty("Text", BindingFlags.Instance | BindingFlags.Public);

        if (textSourceType != null)
        {
            _tsItemField = NF(textSourceType, "item");
            _tsOutputsListField = NF(textSourceType, "outputs");
            _tsIsDividedField = NF(textSourceType, "isDevided");
            _tsDisplayIntField = NF(textSourceType, "displayInterval");
            _tsHideIntField = NF(textSourceType, "hideInterval");
            _tsOutputsProp = textSourceType.GetProperty("Outputs",
                BindingFlags.Instance | BindingFlags.Public);
        }

        _harmony = new Harmony(HarmonyId);

        if (layoutDescType != null)
        {
            var ctor = layoutDescType
                .GetConstructors(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .FirstOrDefault(c => c.GetParameters().Any(p => p.ParameterType == typeof(string)));
            if (ctor != null)
                _harmony.Patch(ctor,
                    prefix: new HarmonyMethod(typeof(WeightTextSourcePatcher), nameof(PatchLayoutDescCtor)));
        }

        if (_textItemType != null)
        {
            PatchGetter(_textItemType, "IsDevidedPerCharacter", nameof(Patch_IsDevided));
            PatchGetter(_textItemType, "DisplayInterval", nameof(Patch_DisplayInt));
            PatchGetter(_textItemType, "HideInterval", nameof(Patch_HideInt));
        }

        if (textSourceType != null)
        {
            var update = textSourceType.GetMethod("Update",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (update != null)
                _harmony.Patch(update,
                    postfix: new HarmonyMethod(typeof(WeightTextSourcePatcher), nameof(PatchUpdate)));
        }
    }

    private static void PatchGetter(Type type, string prop, string method)
    {
        var g = type.GetProperty(prop,
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
            ?.GetGetMethod(nonPublic: true);
        if (g != null)
            _harmony!.Patch(g, postfix: new HarmonyMethod(typeof(WeightTextSourcePatcher), method));
    }

    [HarmonyPrefix]
    private static void PatchLayoutDescCtor(object[] __args)
    {
        if (__args == null) return;
        for (int i = 0; i < __args.Length; i++)
            if (__args[i] is string s && WeightTagParser.HasControlTags(s))
                __args[i] = WeightTagParser.RemoveTags(s);
    }

    [HarmonyPostfix]
    private static void Patch_IsDevided(object __instance, ref bool __result)
    { if (HasTag(__instance)) __result = true; }
    [HarmonyPostfix]
    private static void Patch_DisplayInt(object __instance, ref double __result)
    { if (HasTag(__instance)) __result = Math.Max(__result, ForcedInterval); }
    [HarmonyPostfix]
    private static void Patch_HideInt(object __instance, ref double __result)
    { if (HasTag(__instance)) __result = Math.Max(__result, ForcedInterval); }

    private static bool HasTag(object ti) =>
        _tiTextProp?.GetValue(ti) is string t && WeightTagParser.HasControlTags(t);

    [HarmonyPostfix]
    private static void PatchUpdate(object __instance, object timelineItemSourceDescription)
    {
        try { ApplyTagEffects(__instance, timelineItemSourceDescription); }
        catch { }
    }

    // =================================================================
    // メイン処理
    // =================================================================

    private static void ApplyTagEffects(object ts, object tsDesc)
    {
        var item = _tsItemField?.GetValue(ts);
        if (item == null) return;
        var original = _tiTextProp?.GetValue(item) as string ?? "";
        if (!WeightTagParser.HasControlTags(original)) return;

        var list = _tsOutputsListField?.GetValue(ts) as IList;
        if (list == null || list.Count == 0) return;

        EnsureCache(list[0]!);
        if (_soTimeOffsetField == null) return;

        var duration = GetTime(tsDesc, "ItemDuration");
        var now = GetTime(tsDesc, "ItemPosition");

        var (cleanText, tokens) = WeightTagParser.Parse(original);

        int visCount = cleanText.Count(c => c != '\r' && c != '\n');
        int count = list.Count;
        bool listHasNl = (count > visCount);

        int TagPosToListIdx(int rawPos)
        {
            rawPos = Math.Min(rawPos, cleanText.Length);
            if (listHasNl) return Math.Clamp(rawPos, 0, count);
            int nl = 0;
            for (int i = 0; i < rawPos; i++)
                if (cleanText[i] == '\r' || cleanText[i] == '\n') nl++;
            return Math.Clamp(rawPos - nl, 0, count);
        }

        // ── 自然位置キャッシュ ──────────────────────────────────
        // TextSource が毎フレームlistを再生成するかどうかに関わらず、
        // list.Count が全文字数と一致したフレームの位置を固定して使う。
        // こうすることで RebuildOutputs で一部を除外した後のフレームでも
        // 位置計算の基準が変わらない。
        var cache = _posCache.GetOrCreateValue(item);
        if (list.Count >= cache.Count || cache.Text != original)
        {
            // 全文字が揃っているフレーム → キャッシュを更新
            cache.Count = list.Count;
            cache.Text = original;
            cache.Positions = new PointF[list.Count];
            for (int i = 0; i < list.Count; i++)
                cache.Positions[i] = ReadOffset(list[i]!);
        }
        var naturalPos = (PointF[])cache.Positions.Clone();

        // テキストアイテムの配置設定
        var (hAlign, vAlign, isVertical) = ReadAlignment(item);

        // ── 効果テーブル ──────────────────────────────────────────
        var appearSec = new double[count];
        var disappearSec = Enumerable.Repeat(double.MaxValue, count).ToArray();
        var segmentId = new int[count];

        var segStarts = new List<int> { 0 };
        var clearSegs = new List<(int from, int to, double at)>();
        double cumSec = 0.0;
        int prevClearTo = 0;
        int currentSeg = 0;

        var ordered = tokens.OrderBy(t => t.Pos).ToList();
        for (int ti = 0; ti < ordered.Count; ti++)
        {
            var tok = ordered[ti];
            int pos = TagPosToListIdx(tok.Pos);

            switch (tok.Type)
            {
                case TagType.Wait:
                    cumSec += tok.Value;
                    for (int i = pos; i < count; i++) appearSec[i] = cumSec;
                    break;

                case TagType.WaitCoeff:
                    {
                        int nextPos = count;
                        if (ti + 1 < ordered.Count)
                            nextPos = TagPosToListIdx(ordered[ti + 1].Pos);
                        cumSec += (nextPos - pos) * tok.Value;
                        for (int i = pos; i < count; i++) appearSec[i] = cumSec;
                        break;
                    }

                case TagType.Clear:
                    {
                        double clearAt = cumSec + tok.Value;
                        clearSegs.Add((prevClearTo, pos, clearAt));
                        prevClearTo = pos;
                        currentSeg++;
                        segStarts.Add(pos);
                        for (int i = pos; i < count; i++)
                        {
                            segmentId[i] = currentSeg;
                            appearSec[i] = clearAt;
                        }
                        cumSec = clearAt;
                        break;
                    }


                case TagType.Position:
                    {
                        if (pos >= count) break;
                        // pos以降の可視文字のアンカー点が (X,Y) になるようにシフトする
                        var rangeVis = AllVisIdx(pos, count);
                        var rangeAnchor = CalcAnchor(rangeVis);
                        float sx = (float)tok.X - rangeAnchor.X;
                        float sy = (float)tok.Y - rangeAnchor.Y;
                        for (int i = pos; i < count; i++)
                            naturalPos[i] = new PointF(naturalPos[i].X + sx, naturalPos[i].Y + sy);
                        break;
                    }
            }
        }

        foreach (var (from, to, at) in clearSegs)
            for (int i = from; i < to; i++)
                disappearSec[i] = at;

        // ── クリアセグメントの位置補正 ─────────────────────────
        // 可視文字のインデックス一覧 (改行キャラを除く)
        // 改行に対応する output の DrawingOffset は通常の文字と大きく異なるので除外する
        int[] AllVisIdx(int from, int to)
        {
            if (!listHasNl)
                return Enumerable.Range(from, Math.Max(0, to - from)).ToArray();
            // listHasNl の場合: 極端な位置の output は改行とみなして除外
            // 全文字の平均位置を基準に、大きく外れるものを除外
            var all = Enumerable.Range(from, Math.Max(0, to - from))
                .Where(i => i < naturalPos.Length)
                .ToArray();
            if (all.Length <= 1) return all;
            float avgX = all.Average(i => naturalPos[i].X);
            float avgY = all.Average(i => naturalPos[i].Y);
            float range = Math.Max(
                all.Max(i => Math.Abs(naturalPos[i].X - avgX)),
                all.Max(i => Math.Abs(naturalPos[i].Y - avgY)));
            float thresh = range * 2f + 1f;
            var vis = all.Where(i =>
                Math.Abs(naturalPos[i].X - avgX) < thresh &&
                Math.Abs(naturalPos[i].Y - avgY) < thresh).ToArray();
            return vis.Length > 0 ? vis : all;
        }

        // フル文字列の可視インデックス
        var fullVis = AllVisIdx(0, count);

        // アンカー計算: 縦書きと横書きで主軸・副軸を入れ替える
        // 横書き: X方向 → hAlign, Y方向 → vAlign
        // 縦書き: X方向 → vAlign相当(右中左), Y方向 → hAlign相当(上中下)
        PointF CalcAnchor(int[] vis)
        {
            if (vis.Length == 0) return PointF.Empty;

            float minX = vis.Min(i => naturalPos[i].X);
            float maxX = vis.Max(i => naturalPos[i].X);
            float minY = vis.Min(i => naturalPos[i].Y);
            float maxY = vis.Max(i => naturalPos[i].Y);
            float midX = (minX + maxX) / 2f;
            float midY = (minY + maxY) / 2f;

            // 横書き・縦書き共通:
            //   hAlign (Left/Center/Right) → X 軸の基準点
            //   vAlign (Top/Middle/Bottom) → Y 軸の基準点
            //
            // 縦書きの場合、YMM4 の DrawingOffset は
            //   X: 列の位置 (右端の列が大きい X)
            //   Y: 列内の文字位置 (上が小さい Y)
            // hAlign=Right → 先頭列(最大 X), Left → 末尾列(最小 X)
            // vAlign=Top   → 列の先頭(最小 Y), Bottom → 列の末尾(最大 Y)
            // これは横書きと同じ min/max の対応になるため共通化できる。

            float ax = hAlign switch
            {
                HAlign.Left => minX,
                HAlign.Center => midX,
                HAlign.Right => maxX,
                _ => minX,
            };
            float ay = vAlign switch
            {
                VAlign.Top => minY,
                VAlign.Middle => midY,
                VAlign.Bottom => maxY,
                _ => minY,
            };
            return new PointF(ax, ay);
        }

        var fullAnchor = CalcAnchor(fullVis);

        var segShiftX = new float[segStarts.Count];
        var segShiftY = new float[segStarts.Count];

        for (int s = 0; s < segStarts.Count; s++)
        {
            int from = segStarts[s];
            int to = (s + 1 < segStarts.Count) ? segStarts[s + 1] : count;
            if (from >= count) continue;

            var segVis = AllVisIdx(from, to);
            var segAnchor = CalcAnchor(segVis);

            segShiftX[s] = fullAnchor.X - segAnchor.X;
            segShiftY[s] = fullAnchor.Y - segAnchor.Y;
        }

        // ── Output に書き込む ──────────────────────────────────
        for (int i = 0; i < count; i++)
        {
            var output = list[i]!;
            int seg = segmentId[i];

            _soTimeOffsetField.SetValue(output, TimeSpan.FromSeconds(-appearSec[i]));

            if (_soDurationOffsetField != null)
            {
                TimeSpan dOff = disappearSec[i] == double.MaxValue
                    ? TimeSpan.Zero
                    : TimeSpan.FromSeconds(disappearSec[i]) - duration;
                _soDurationOffsetField.SetValue(output, dOff);
            }

            WriteOffset(output, new PointF(
                naturalPos[i].X + segShiftX[seg],
                naturalPos[i].Y + segShiftY[seg]));
        }

        RebuildOutputs(ts, list, now, duration);
    }

    // =================================================================
    // 配置設定の読み取り
    // =================================================================

    private enum HAlign { Left, Center, Right }
    private enum VAlign { Top, Middle, Bottom }

    private static (HAlign h, VAlign v, bool isVertical) ReadAlignment(object item)
    {
        if (_textItemType == null) return (HAlign.Left, VAlign.Top, false);

        // ── 縦書き判定 ────────────────────────────────────────────
        bool isV = false;
        foreach (var wProp in new[] { "WritingMode", "TextDirection", "Direction", "VerticalText" })
        {
            var wval = _textItemType
                .GetProperty(wProp, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                ?.GetValue(item)?.ToString() ?? "";
            if (wval.Length == 0) continue;
            isV = wval.Contains("Vertical", StringComparison.OrdinalIgnoreCase)
               || wval.Contains("Tate", StringComparison.OrdinalIgnoreCase)
               || wval == "1" || wval == "2" || wval == "True";
            break;
        }

        // ── 配置設定の読み取り ────────────────────────────────────
        // パターン1: 水平・垂直を別プロパティで保持
        // パターン2: BasePoint 単一プロパティ (例: "TopCenter", "MiddleLeft")
        // パターン3: 数値 enum (0=TL,1=TC,2=TR,3=ML,4=MC,5=MR,6=BL,7=BC,8=BR)

        HAlign h = HAlign.Left;
        VAlign v = VAlign.Top;
        bool parsed = false;

        var hStr = ReadPropStr(item, "HorizontalAlignment", "TextHorizontalAlignment", "AlignX");
        var vStr = ReadPropStr(item, "VerticalAlignment", "TextVerticalAlignment", "AlignY");
        if (hStr != null) { h = ParseHStr(hStr); parsed = true; }
        if (vStr != null) { v = ParseVStr(vStr); parsed = true; }

        if (!parsed)
        {
            var bp = ReadPropStr(item, "BasePoint", "Alignment", "TextAlignment", "AnchorPoint") ?? "";

            // 数値: 0=TL,1=TC,2=TR,3=ML,4=MC,5=MR,6=BL,7=BC,8=BR
            if (int.TryParse(bp, out int num))
            {
                h = (HAlign)(num % 3);
                v = (VAlign)(num / 3);
            }
            else
            {
                var words = Regex.Matches(bp, "[A-Z][a-z]+|[A-Z]+(?=[A-Z]|$)")
                    .Select(m => m.Value)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                h = ParseHWords(words);
                v = ParseVWords(words);
            }
        }

        return (h, v, isV);
    }

    private static string? ReadPropStr(object item, params string[] names)
    {
        if (_textItemType == null) return null;
        foreach (var name in names)
        {
            var val = _textItemType
                .GetProperty(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                ?.GetValue(item)?.ToString();
            if (!string.IsNullOrEmpty(val)) return val;
        }
        return null;
    }

    private static HAlign ParseHStr(string s)
    {
        if (int.TryParse(s, out int n)) return (HAlign)Math.Clamp(n, 0, 2);
        if (s.Contains("Center", StringComparison.OrdinalIgnoreCase)
         || s.Contains("Centre", StringComparison.OrdinalIgnoreCase)) return HAlign.Center;
        if (s.Contains("Right", StringComparison.OrdinalIgnoreCase)) return HAlign.Right;
        return HAlign.Left;
    }

    private static VAlign ParseVStr(string s)
    {
        if (int.TryParse(s, out int n)) return (VAlign)Math.Clamp(n, 0, 2);
        if (s.Contains("Middle", StringComparison.OrdinalIgnoreCase)
         || s.Contains("Center", StringComparison.OrdinalIgnoreCase)
         || s.Contains("Centre", StringComparison.OrdinalIgnoreCase)) return VAlign.Middle;
        if (s.Contains("Bottom", StringComparison.OrdinalIgnoreCase)
         || s.Contains("Lower", StringComparison.OrdinalIgnoreCase)) return VAlign.Bottom;
        return VAlign.Top;
    }

    private static HAlign ParseHWords(HashSet<string> words)
    {
        if (words.Contains("Center") || words.Contains("Centre")) return HAlign.Center;
        if (words.Contains("Right")) return HAlign.Right;
        return HAlign.Left;
    }

    private static VAlign ParseVWords(HashSet<string> words)
    {
        if (words.Contains("Upper") || words.Contains("Top")) return VAlign.Top;
        else if (words.Contains("Middle")) return VAlign.Middle;
        else if (words.Contains("Lower") || words.Contains("Bottom")) return VAlign.Bottom;
        return VAlign.Top;
    }

    // =================================================================
    // Outputs 再構築
    // =================================================================

    private static void RebuildOutputs(object ts, IList all, TimeSpan now, TimeSpan dur)
    {
        if (_tsOutputsProp == null) return;

        var elemType = all.GetType().GetGenericArguments()[0];
        var result = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(elemType))!;

        foreach (var output in all)
        {
            if (output == null) continue;
            var tOff = _soTimeOffsetField != null ? (TimeSpan)_soTimeOffsetField.GetValue(output)! : TimeSpan.Zero;
            var dOff = _soDurationOffsetField != null ? (TimeSpan)_soDurationOffsetField.GetValue(output)! : TimeSpan.Zero;

            if (now + tOff < TimeSpan.Zero) continue;
            if (dOff != TimeSpan.Zero && now >= dur + dOff) continue;

            result.Add(output);
        }

        _tsOutputsProp.SetValue(ts, result);
    }

    // =================================================================
    // DrawingOffset 読み書き
    // =================================================================

    private static PointF ReadOffset(object output)
    {
        if (_soDrawingOffsetField == null) return PointF.Empty;
        return ToPointF(_soDrawingOffsetField.GetValue(output));
    }

    private static void WriteOffset(object output, PointF v)
    {
        if (_soDrawingOffsetField == null) return;
        var ft = _soDrawingOffsetField.FieldType;
        if (ft == typeof(PointF)) { _soDrawingOffsetField.SetValue(output, v); return; }
        try
        {
            var inst = Activator.CreateInstance(ft)!;
            (ft.GetField("X") ?? ft.GetField("x"))?.SetValue(inst, Convert.ChangeType(v.X, typeof(float)));
            (ft.GetField("Y") ?? ft.GetField("y"))?.SetValue(inst, Convert.ChangeType(v.Y, typeof(float)));
            _soDrawingOffsetField.SetValue(output, inst);
        }
        catch { }
    }

    private static PointF ToPointF(object? v)
    {
        if (v == null) return PointF.Empty;
        if (v is PointF pf) return pf;
        var t = v.GetType();
        return new PointF(
            ToF((t.GetField("X") ?? t.GetField("x"))?.GetValue(v)),
            ToF((t.GetField("Y") ?? t.GetField("y"))?.GetValue(v)));
    }
    private static float ToF(object? v) => v == null ? 0f : Convert.ToSingle(v);

    // =================================================================
    // ヘルパー
    // =================================================================

    private static void EnsureCache(object first)
    {
        if (_soTimeOffsetField != null) return;
        var t = first.GetType();
        _soTimeOffsetField = BF(t, "TimeOffset");
        _soDurationOffsetField = BF(t, "DurationOffset");
        _soDrawingOffsetField = BF(t, "DrawingOffset");
    }

    private static FieldInfo? BF(Type t, string name)
        => t.GetField($"<{name}>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic)
        ?? t.GetField(char.ToLower(name[0]) + name[1..], BindingFlags.Instance | BindingFlags.NonPublic)
        ?? t.GetField("_" + char.ToLower(name[0]) + name[1..], BindingFlags.Instance | BindingFlags.NonPublic)
        ?? t.GetField(name, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);

    private static FieldInfo? NF(Type t, string name)
        => t.GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);

    private static TimeSpan GetTime(object desc, string propName)
    {
        try
        {
            var inner = desc.GetType().GetProperty(propName)?.GetValue(desc);
            var time = inner?.GetType().GetProperty("Time")?.GetValue(inner);
            return time is TimeSpan ts ? ts : TimeSpan.Zero;
        }
        catch { return TimeSpan.Zero; }
    }
}