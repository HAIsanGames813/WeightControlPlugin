using HarmonyLib;
using System.Collections;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace WeightControlPlugin;

internal static class WeightTextSourcePatcher
{
    private const string HarmonyId = "com.weightcontrolplugin.textsource";
    private static Harmony? _harmony;

    private sealed class SourceTypeInfo
    {
        public required PropertyInfo OutputsProp;
        public required FieldInfo ItemField;
        public required Func<object, string?> GetText;
    }

    // DrawingOffset フィールド (SourceOutput に存在、遅延初期化)
    private static FieldInfo? _drawingOffsetField;
    private static readonly object _drawingOffsetLock = new();

    // source インスタンスごとの「全文字の自然位置キャッシュ」
    private sealed class NaturalPosCache
    {
        public PointF[] Positions = Array.Empty<PointF>();
        public string Text = string.Empty;
        public int Count;
        // 前回書き込んだ output[0] の DrawingOffset を保存する。
        // YMM4 が outputs を再生成したとき (BasePoint 変更など) は
        // output[0] の値がここと異なるのでキャッシュを破棄できる。
        public PointF LastWrittenPos0;
    }
    private static readonly ConditionalWeakTable<object, NaturalPosCache> _posCache = new();

    private static readonly Dictionary<Type, SourceTypeInfo?> _typeCache = new();
    private static readonly object _cacheLock = new();
    private static readonly ConditionalWeakTable<object, object> _taggedChars = new();
    private static readonly object _marker = new();

    // -----------------------------------------------------------------
    public static void Apply()
    {
        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            try { TryPatch(asm); } catch { }
        AppDomain.CurrentDomain.AssemblyLoad += (_, args) =>
        { try { TryPatch(args.LoadedAssembly); } catch { } };
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

        _harmony = new Harmony(HarmonyId);

        // TextLayoutDescription ctor: タグ除去
        var ldType = types.FirstOrDefault(t => t.Name == "TextLayoutDescription");
        if (ldType != null)
        {
            try
            {
                var ctor = ldType
                    .GetConstructors(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                    .FirstOrDefault(c => c.GetParameters().Any(p => p.ParameterType == typeof(string)));
                if (ctor != null)
                    _harmony.Patch(ctor,
                        prefix: new HarmonyMethod(typeof(WeightTextSourcePatcher), nameof(PatchLayoutDescCtor)));
            }
            catch { }
        }

        // IsDevidedPerCharacter=true, DisplayInterval=0, HideInterval=0 に強制
        // → GetDisplayOutputs が全文字を Outputs に返す
        foreach (var t in types)
        {
            try
            {
                if (t.GetProperty("IsDevidedPerCharacter",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic) == null) continue;

                bool isItem = t.FullName?.StartsWith("YukkuriMovieMaker.Project.Items.") ?? false;
                if (isItem)
                {
                    PatchGetter(t, "IsDevidedPerCharacter", nameof(Getter_IsDevided_Item));
                    PatchGetter(t, "DisplayInterval", nameof(Getter_DisplayInt_Item));
                    PatchGetter(t, "HideInterval", nameof(Getter_HideInt_Item));
                }
                else
                {
                    PatchGetter(t, "IsDevidedPerCharacter", nameof(Getter_IsDevided_Char));
                    PatchGetter(t, "DisplayInterval", nameof(Getter_DisplayInt_Char));
                    PatchGetter(t, "HideInterval", nameof(Getter_HideInt_Char));
                }
            }
            catch { }
        }

        // TextSource / JimakuSource の Update をパッチ
        foreach (var srcType in types)
        {
            try
            {
                if (FindField(srcType, "outputs") == null) continue;
                if (FindField(srcType, "isDevided", "isDivided") == null) continue;
                if (srcType.GetField("item", BindingFlags.Instance | BindingFlags.NonPublic) == null) continue;
                if (srcType.GetProperty("Outputs", BindingFlags.Instance | BindingFlags.Public) == null) continue;

                var update = srcType.GetMethod("Update",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (update == null) continue;

                _harmony.Patch(update,
                    prefix: new HarmonyMethod(typeof(WeightTextSourcePatcher), nameof(PrefixUpdate)),
                    postfix: new HarmonyMethod(typeof(WeightTextSourcePatcher), nameof(PostfixUpdate)));
            }
            catch { }
        }
    }

    private static void PatchGetter(Type t, string prop, string method)
    {
        var g = t.GetProperty(prop, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                 ?.GetGetMethod(nonPublic: true);
        if (g != null)
            _harmony!.Patch(g, postfix: new HarmonyMethod(typeof(WeightTextSourcePatcher), method));
    }

    // ----------------------------------------------------------------- ゲッターパッチ

    [HarmonyPostfix]
    private static void Getter_IsDevided_Item(object __instance, ref bool __result)
    { if (ItemHasTag(__instance)) __result = true; }

    [HarmonyPostfix]
    private static void Getter_DisplayInt_Item(object __instance, ref double __result)
    { if (ItemHasTag(__instance)) __result = 0.0; }

    [HarmonyPostfix]
    private static void Getter_HideInt_Item(object __instance, ref double __result)
    { if (ItemHasTag(__instance)) __result = 0.0; }

    [HarmonyPostfix]
    private static void Getter_IsDevided_Char(object __instance, ref bool __result)
    { if (_taggedChars.TryGetValue(__instance, out _)) __result = true; }

    [HarmonyPostfix]
    private static void Getter_DisplayInt_Char(object __instance, ref double __result)
    { if (_taggedChars.TryGetValue(__instance, out _)) __result = 0.0; }

    [HarmonyPostfix]
    private static void Getter_HideInt_Char(object __instance, ref double __result)
    { if (_taggedChars.TryGetValue(__instance, out _)) __result = 0.0; }

    private static bool ItemHasTag(object item)
    {
        foreach (var name in new[] { "Serif", "Text", "Body", "Content" })
        {
            var v = item.GetType()
                .GetProperty(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                ?.GetValue(item) as string;
            if (v != null) return WeightTagParser.HasControlTags(v);
        }
        return false;
    }

    // ----------------------------------------------------------------- TextLayoutDescription ctor

    [HarmonyPrefix]
    private static void PatchLayoutDescCtor(object[] __args)
    {
        if (__args == null) return;
        for (int i = 0; i < __args.Length; i++)
            if (__args[i] is string s && WeightTagParser.HasControlTags(s))
                __args[i] = WeightTagParser.RemoveTags(s);
    }

    // ----------------------------------------------------------------- Update prefix

    [HarmonyPrefix]
    private static void PrefixUpdate(object __instance)
    {
        try
        {
            var info = GetInfo(__instance);
            if (info == null) return;
            var item = info.ItemField.GetValue(__instance);
            if (item == null) return;

            bool hasTag = WeightTagParser.HasControlTags(info.GetText(__instance));

            var character = item.GetType()
                .GetProperty("Character", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                ?.GetValue(item);
            if (character == null) return;

            if (hasTag)
                _taggedChars.AddOrUpdate(character, _marker);
            else
                _taggedChars.Remove(character);
        }
        catch { }
    }

    // ----------------------------------------------------------------- Update postfix
    // __args を使う: TextSource と JimakuSource でパラメータ名が異なるため

    [HarmonyPostfix]
    private static void PostfixUpdate(object __instance, object[] __args)
    {
        try
        {
            var tsDesc = (__args != null && __args.Length > 0) ? __args[0] : null;
            FilterOutputsByTags(__instance, tsDesc);
        }
        catch { }
    }

    // ================================================================= メイン処理

    private static void FilterOutputsByTags(object ts, object? tsDesc)
    {
        var info = GetInfo(ts);
        if (info == null) return;

        var original = info.GetText(ts) ?? "";
        if (!WeightTagParser.HasControlTags(original)) return;

        // displayInterval=0 により GetDisplayOutputs が全文字を Outputs に入れている
        var allOutputs = info.OutputsProp.GetValue(ts) as System.Collections.Generic.IReadOnlyList<object>;
        if (allOutputs == null || allOutputs.Count == 0) return;

        int count = allOutputs.Count;
        double now = GetTimeSec(tsDesc, "ItemPosition");

        // DrawingOffset フィールドの遅延初期化
        EnsureDrawingOffsetField(allOutputs[0]);

        var (cleanText, tokens) = WeightTagParser.Parse(original);

        // 改行を除いた文字数
        int visChars = 0;
        foreach (char c in cleanText)
            if (c != '\r' && c != '\n') visChars++;

        if (count < 2 && visChars > 1) return;

        // ── 自然位置キャッシュ ────────────────────────────────────────────
        // テキスト・文字数変化、または YMM4 が outputs を再生成したときにキャッシュを破棄する。
        // 再生成検出: 前回書き込んだ output[0] の値と現在値が異なれば再生成済み。
        var cache = _posCache.GetOrCreateValue(ts);
        bool cacheInvalid = cache.Count != count || cache.Text != original;
        if (!cacheInvalid && count > 0)
        {
            var cur0 = ReadDrawingOffset(allOutputs[0]);
            if (cur0.X != cache.LastWrittenPos0.X || cur0.Y != cache.LastWrittenPos0.Y)
                cacheInvalid = true;
        }
        if (cacheInvalid)
        {
            cache.Count = count;
            cache.Text = original;
            cache.Positions = new PointF[count];
            for (int i = 0; i < count; i++)
                cache.Positions[i] = ReadDrawingOffset(allOutputs[i]);
        }
        var naturalPos = (PointF[])cache.Positions.Clone();

        // ── cleanText 内の文字位置 → outputs インデックス ─────────
        int ToIdx(int rawPos)
        {
            rawPos = Math.Min(rawPos, cleanText.Length);
            int nl = 0;
            for (int i = 0; i < rawPos; i++)
            {
                char c = cleanText[i];
                if (c == '\r' || c == '\n') nl++;
            }
            return Math.Clamp(rawPos - nl, 0, count);
        }

        // ── アンカー検出 (BasePoint を自動推定) ──────────────────
        int hKind = 1, vKind = 1;   // 0=min, 1=center, 2=max
        static int NearestZeroKind(float a, float b, float c)
        {
            float da = Math.Abs(a), db = Math.Abs(b), dc = Math.Abs(c);
            if (da <= db && da <= dc) return 0;
            if (dc <= da && dc <= db) return 2;
            return 1;
        }
        static float ByKind(float a, float b, float c, int k) => k switch { 0 => a, 2 => c, _ => b };

        if (count > 1)
        {
            float minX = naturalPos.Min(p => p.X), maxX = naturalPos.Max(p => p.X);
            float minY = naturalPos.Min(p => p.Y), maxY = naturalPos.Max(p => p.Y);
            hKind = NearestZeroKind(minX, (minX + maxX) / 2f, maxX);
            vKind = NearestZeroKind(minY, (minY + maxY) / 2f, maxY);
        }

        PointF CalcAnchor(int from, int to)
        {
            to = Math.Min(to, count);
            if (from >= to) return PointF.Empty;
            float minX = float.MaxValue, maxX = float.MinValue;
            float minY = float.MaxValue, maxY = float.MinValue;
            for (int i = from; i < to; i++)
            {
                minX = Math.Min(minX, naturalPos[i].X); maxX = Math.Max(maxX, naturalPos[i].X);
                minY = Math.Min(minY, naturalPos[i].Y); maxY = Math.Max(maxY, naturalPos[i].Y);
            }
            return new PointF(ByKind(minX, (minX + maxX) / 2f, maxX, hKind),
                              ByKind(minY, (minY + maxY) / 2f, maxY, vKind));
        }

        // ── タグ解析: appear/disappear + セグメント境界 ──────────
        var appearSec = new double[count];
        var disappearSec = new double[count];
        for (int i = 0; i < count; i++) disappearSec[i] = double.MaxValue;

        var segStarts = new System.Collections.Generic.List<int> { 0 };
        double cumSec = 0.0;
        int prevClearTo = 0;
        var clearPairs = new System.Collections.Generic.List<(int from, int to, double at)>();

        var ordered = tokens.OrderBy(t => t.Pos).ToList();
        for (int ti = 0; ti < ordered.Count; ti++)
        {
            var tok = ordered[ti];
            int pos = ToIdx(tok.Pos);

            switch (tok.Type)
            {
                case TagType.Wait:
                    cumSec += tok.Value;
                    for (int i = pos; i < count; i++) appearSec[i] = cumSec;
                    break;

                case TagType.WaitCoeff:
                    {
                        int npos = (ti + 1 < ordered.Count) ? ToIdx(ordered[ti + 1].Pos) : count;
                        cumSec += (npos - pos) * tok.Value;
                        for (int i = pos; i < count; i++) appearSec[i] = cumSec;
                        break;
                    }

                case TagType.Clear:
                    {
                        double at = cumSec + tok.Value;
                        clearPairs.Add((prevClearTo, pos, at));
                        prevClearTo = pos;
                        segStarts.Add(pos);
                        for (int i = pos; i < count; i++) appearSec[i] = at;
                        cumSec = at;
                        break;
                    }

                case TagType.ClearCoeff:
                    {
                        int npos = (ti + 1 < ordered.Count) ? ToIdx(ordered[ti + 1].Pos) : count;
                        double at = cumSec + (npos - pos) * tok.Value;
                        clearPairs.Add((prevClearTo, pos, at));
                        prevClearTo = pos;
                        segStarts.Add(pos);
                        for (int i = pos; i < count; i++) appearSec[i] = at;
                        cumSec = at;
                        break;
                    }

                case TagType.Position:
                    {
                        // <p> タグ: pos 以降を (tok.X, tok.Y) からのオフセットに移動
                        if (pos >= count) break;
                        var anchor = CalcAnchor(pos, count);
                        float sx = (float)tok.X - anchor.X;
                        float sy = (float)tok.Y - anchor.Y;
                        for (int i = pos; i < count; i++)
                            naturalPos[i] = new PointF(naturalPos[i].X + sx, naturalPos[i].Y + sy);
                        break;
                    }
            }
        }

        foreach (var (from, to, at) in clearPairs)
            for (int i = from; i < to; i++)
                disappearSec[i] = at;

        // ── セグメント位置揃え ─────────────────────────────────
        // 全セグメントのアンカーを全文字のアンカーに合わせる
        var fullAnchor = CalcAnchor(0, count);
        for (int s = 0; s < segStarts.Count; s++)
        {
            int from = segStarts[s];
            int to = (s + 1 < segStarts.Count) ? segStarts[s + 1] : count;
            var segAnchor = CalcAnchor(from, to);
            float sx = fullAnchor.X - segAnchor.X;
            float sy = fullAnchor.Y - segAnchor.Y;
            for (int i = from; i < to; i++)
                naturalPos[i] = new PointF(naturalPos[i].X + sx, naturalPos[i].Y + sy);
        }

        // ── DrawingOffset 書き込み & フィルタリング ──────────────
        var elemType = allOutputs.GetType().GetGenericArguments().FirstOrDefault() ?? typeof(object);
        var result = (IList)Activator.CreateInstance(typeof(System.Collections.Generic.List<>).MakeGenericType(elemType))!;

        for (int i = 0; i < count; i++)
        {
            // 位置を書き込む (表示されるかどうかに関わらず自然位置を更新)
            WriteDrawingOffset(allOutputs[i], naturalPos[i]);

            // タイミングフィルタ
            if (now < appearSec[i]) continue;
            if (disappearSec[i] != double.MaxValue && now >= disappearSec[i]) continue;
            result.Add(allOutputs[i]);
        }

        // output[0] の書き込み後の位置を記録 (次フレームのキャッシュ検証用)
        if (count > 0) cache.LastWrittenPos0 = naturalPos[0];

        SetOutputs(ts, info, result);
    }

    // ================================================================= DrawingOffset 読み書き

    private static void EnsureDrawingOffsetField(object output)
    {
        if (_drawingOffsetField != null) return;
        lock (_drawingOffsetLock)
        {
            if (_drawingOffsetField != null) return;
            var t = output.GetType();
            // バッキングフィールドまたは camelCase フィールドを探す
            _drawingOffsetField =
                t.GetField("<DrawingOffset>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic)
                ?? t.GetField("drawingOffset", BindingFlags.Instance | BindingFlags.NonPublic)
                ?? t.GetField("_drawingOffset", BindingFlags.Instance | BindingFlags.NonPublic)
                ?? t.GetProperty("DrawingOffset", BindingFlags.Instance | BindingFlags.Public)?
                    .GetType().GetField("value");  // fallback (通常不要)
        }
    }

    private static PointF ReadDrawingOffset(object output)
    {
        if (_drawingOffsetField == null) return PointF.Empty;
        return ToPointF(_drawingOffsetField.GetValue(output));
    }

    private static void WriteDrawingOffset(object output, PointF v)
    {
        if (_drawingOffsetField == null) return;
        var ft = _drawingOffsetField.FieldType;
        if (ft == typeof(PointF))
        {
            _drawingOffsetField.SetValue(output, v);
            return;
        }
        // 独自構造体の場合 (X/Y フィールドに書き込む)
        try
        {
            var inst = Activator.CreateInstance(ft)!;
            (ft.GetField("X") ?? ft.GetField("x"))?.SetValue(inst, Convert.ChangeType(v.X, typeof(float)));
            (ft.GetField("Y") ?? ft.GetField("y"))?.SetValue(inst, Convert.ChangeType(v.Y, typeof(float)));
            _drawingOffsetField.SetValue(output, inst);
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

    // ================================================================= キャッシュ

    private static SourceTypeInfo? GetInfo(object src)
    {
        var t = src.GetType();
        lock (_cacheLock)
        {
            if (_typeCache.TryGetValue(t, out var cached)) return cached;
        }

        if (FindField(t, "outputs") == null || FindField(t, "isDevided", "isDivided") == null)
        {
            lock (_cacheLock) { _typeCache[t] = null; }
            return null;
        }

        var itemFld = t.GetField("item", BindingFlags.Instance | BindingFlags.NonPublic);
        var outsProp = t.GetProperty("Outputs", BindingFlags.Instance | BindingFlags.Public);

        if (itemFld == null || outsProp == null)
        {
            lock (_cacheLock) { _typeCache[t] = null; }
            return null;
        }

        var captured = itemFld;
        static string? ReadText(FieldInfo fld, object s)
        {
            var item = fld.GetValue(s);
            if (item == null) return null;
            foreach (var name in new[] { "Serif", "Text", "Body", "Content" })
            {
                var v = item.GetType()
                    .GetProperty(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                    ?.GetValue(item) as string;
                if (v != null) return v;
            }
            return null;
        }

        var info = new SourceTypeInfo
        {
            ItemField = itemFld,
            OutputsProp = outsProp,
            GetText = s => ReadText(captured, s),
        };
        lock (_cacheLock) { _typeCache[t] = info; }
        return info;
    }

    // ================================================================= ヘルパー

    private static void SetOutputs(object ts, SourceTypeInfo info, IList value)
    {
        var setter = info.OutputsProp.GetSetMethod(nonPublic: true);
        if (setter != null) { setter.Invoke(ts, new object[] { value }); return; }
        var bf = ts.GetType().GetField("<Outputs>k__BackingField",
            BindingFlags.Instance | BindingFlags.NonPublic);
        bf?.SetValue(ts, value);
    }

    private static double GetTimeSec(object? desc, string prop)
    {
        if (desc == null) return 0.0;
        try
        {
            var inner = desc.GetType().GetProperty(prop)?.GetValue(desc);
            var time = inner?.GetType().GetProperty("Time")?.GetValue(inner);
            return time is TimeSpan ts ? ts.TotalSeconds : 0.0;
        }
        catch { return 0.0; }
    }

    private static FieldInfo? FindField(Type t, params string[] names)
    {
        for (var cur = t; cur != null && cur != typeof(object); cur = cur.BaseType)
            foreach (var name in names)
            {
                var f = cur.GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
                if (f != null) return f;
            }
        return null;
    }
}