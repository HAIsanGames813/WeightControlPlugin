using HarmonyLib;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;

namespace WeightControlPlugin;

/// <summary>
/// Harmony で YukkuriMovieMaker.Controls.RichTextEditor (YukkuriMovieMaker.Plugin.dll) の
/// OnApplyTemplate() にパッチを当て、PARTS_TextBox / PARTS_RichTextBox の
/// ContextMenu に制御コード挿入項目を追加します。
///
/// ■ 有識者情報に基づく実装
///   - 対象クラス: YukkuriMovieMaker.Controls.RichTextEditor
///   - 対象アセンブリ: YukkuriMovieMaker.Plugin.dll
///   - 内部コントロール名: PARTS_TextBox / PARTS_RichTextBox (GetTemplateChild で取得)
/// </summary>
internal static class RichTextEditorPatcher
{
    private const string HarmonyId = "com.weightcontrolplugin";
    private static Harmony? _harmony;
    private const string MenuTag = "WeightPlugin";
    private static readonly string LogPath =
        Path.Combine(Path.GetTempPath(), "WeightControlPlugin_log.txt");

    // 購読済みコントロールの追跡 (ダブル登録防止)
    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<object, object>
        _hooked = new();
    private static readonly object _marker = new();

    public static void Apply()
    {
        Log("Apply() 開始");

        // YukkuriMovieMaker.Plugin.dll はプラグイン読み込み前にロード済みのはず
        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            TryPatch(asm);

        // 万が一まだロードされていない場合の保険
        AppDomain.CurrentDomain.AssemblyLoad += (_, args) => TryPatch(args.LoadedAssembly);

        Log($"Apply() 完了 (パッチ済み={_harmony != null})");
    }

    public static void Unpatch()
    {
        AppDomain.CurrentDomain.AssemblyLoad -= (_, args) => TryPatch(args.LoadedAssembly);
        _harmony?.UnpatchAll(HarmonyId);
        _harmony = null;
    }

    private static void TryPatch(Assembly asm)
    {
        if (_harmony != null) return;

        // 対象: YukkuriMovieMaker.Plugin.dll
        if (!asm.GetName().Name!.Contains("YukkuriMovieMaker"))
            return;

        Type? editorType = null;
        try
        {
            // YukkuriMovieMaker.Controls.RichTextEditor を探す
            editorType = asm.GetTypes().FirstOrDefault(t =>
                t.FullName == "YukkuriMovieMaker.Controls.RichTextEditor" ||
                (t.Name == "RichTextEditor" &&
                 typeof(System.Windows.Controls.Control).IsAssignableFrom(t)));
        }
        catch (ReflectionTypeLoadException ex)
        {
            // 一部の型がロードできなくても続行
            editorType = ex.Types.FirstOrDefault(t =>
                t?.Name == "RichTextEditor" &&
                t.IsSubclassOf(typeof(System.Windows.Controls.Control)));
            Log($"ReflectionTypeLoadException (継続): {ex.LoaderExceptions.FirstOrDefault()?.Message}");
        }
        catch (Exception ex)
        {
            Log($"GetTypes 例外 [{asm.GetName().Name}]: {ex.Message}");
            return;
        }

        if (editorType == null) return;

        Log($"RichTextEditor 発見: {editorType.FullName} ({asm.GetName().Name})");

        // OnApplyTemplate は public override → Harmony から安全にパッチできる
        var method = editorType.GetMethod("OnApplyTemplate",
            BindingFlags.Instance | BindingFlags.Public);

        if (method == null)
        {
            Log("OnApplyTemplate が見つかりません");
            return;
        }

        _harmony = new Harmony(HarmonyId);
        _harmony.Patch(method,
            postfix: new HarmonyMethod(typeof(RichTextEditorPatcher),
                nameof(Postfix_OnApplyTemplate)));

        Log("Harmony パッチ完了");
    }

    // ── Postfix: OnApplyTemplate() 完了後 ────────────────────

    [HarmonyPostfix]
    private static void Postfix_OnApplyTemplate(object __instance)
    {
        if (_hooked.TryGetValue(__instance, out _)) return;
        _hooked.Add(__instance, _marker);

        Log($"Postfix_OnApplyTemplate: {__instance.GetType().Name}");

        // RichTextEditor は Control なので GetTemplateChild を呼べる
        // ただし internal なので型を経由して呼ぶ
        var getTemplateChild = __instance.GetType().GetMethod(
            "GetTemplateChild",
            BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);

        if (getTemplateChild == null)
        {
            // FrameworkElement の GetTemplateChild (protected) を取得
            getTemplateChild = typeof(FrameworkElement).GetMethod(
                "GetTemplateChild",
                BindingFlags.Instance | BindingFlags.NonPublic);
        }

        if (getTemplateChild == null)
        {
            Log("GetTemplateChild が見つかりません");
            return;
        }

        // PARTS_TextBox / PARTS_RichTextBox を取得
        var textBox = getTemplateChild.Invoke(__instance, ["PARTS_TextBox"]) as TextBox;
        var richTextBox = getTemplateChild.Invoke(__instance, ["PARTS_RichTextBox"]) as RichTextBox;

        Log($"textBox={textBox?.GetType().Name ?? "null"}, richTextBox={richTextBox?.GetType().Name ?? "null"}");

        // OnApplyTemplate 直後はまだ null になることがある → Loaded で再試行
        if (textBox == null || richTextBox == null)
        {
            Log("null → Loaded イベントで再試行");
            if (__instance is FrameworkElement fe)
            {
                fe.Loaded += delegate
                {
                    var tb = getTemplateChild.Invoke(__instance, ["PARTS_TextBox"]) as TextBox;
                    var rtb = getTemplateChild.Invoke(__instance, ["PARTS_RichTextBox"]) as RichTextBox;
                    if (tb != null && rtb != null)
                        SubscribeContextMenu(tb, rtb);
                    else
                        Log("Loaded 後も null");
                };
            }
            return;
        }

        SubscribeContextMenu(textBox, richTextBox);
    }

    // ── ContextMenuOpening 購読 ───────────────────────────────

    private static void SubscribeContextMenu(TextBox textBox, RichTextBox richTextBox)
    {
        Log("SubscribeContextMenu");

        textBox.ContextMenuOpening += (s, e) =>
            InjectMenuItems(s as FrameworkElement, textBox, richTextBox, useRich: false);

        richTextBox.ContextMenuOpening += (s, e) =>
            InjectMenuItems(s as FrameworkElement, textBox, richTextBox, useRich: true);

        Log("ContextMenuOpening 購読完了");
    }

    private static void InjectMenuItems(
        FrameworkElement? source,
        TextBox textBox, RichTextBox richTextBox,
        bool useRich)
    {
        var menu = source?.ContextMenu;
        if (menu == null)
        {
            Log("ContextMenu が null");
            return;
        }

        Log($"メニュー注入 (items={menu.Items.Count}, useRich={useRich})");

        // 前回分を削除
        menu.Items.OfType<FrameworkElement>()
            .Where(x => x.Tag as string == MenuTag)
            .ToList()
            .ForEach(x => menu.Items.Remove(x));

        // ── 区切り ────────────────────────────────────────
        menu.Items.Add(new Separator { Tag = MenuTag });

        // ── <w> サブメニュー ──────────────────────────────
        var wMenu = new MenuItem { Header = "ウェイト <w> を挿入", Tag = MenuTag };

        foreach (var (label, tag) in new[]
        {
            ("0.1 秒  <w0.1>", "<w0.1>"),
            ("0.5 秒  <w0.5>", "<w0.5>"),
            ("1 秒    <w1>",   "<w1>"),
            ("2 秒    <w2>",   "<w2>"),
            ("3 秒    <w3>",   "<w3>"),
        })
        {
            var t = tag;
            var mi = new MenuItem { Header = label };
            mi.Click += (_, __) => Insert(textBox, richTextBox, useRich, t);
            wMenu.Items.Add(mi);
        }

        wMenu.Items.Add(new Separator());
        var wMult = new MenuItem { Header = "文字数 × 係数  <w*係数>" };
        foreach (var (label, tag) in new[]
        {
            ("× 0.1  <w*0.1>", "<w*0.1>"),
            ("× 0.2  <w*0.2>", "<w*0.2>"),
            ("× 0.5  <w*0.5>", "<w*0.5>"),
        })
        {
            var t = tag;
            var mi = new MenuItem { Header = label };
            mi.Click += (_, __) => Insert(textBox, richTextBox, useRich, t);
            wMult.Items.Add(mi);
        }
        wMenu.Items.Add(wMult);
        wMenu.Items.Add(new Separator());

        var wCustom = new MenuItem { Header = "カスタム入力..." };
        wCustom.Click += (_, __) =>
        {
            var win = new TagInputWindow("ウェイト挿入", "待機秒数 (例: 1.5)", "1",
                s => double.TryParse(s,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out double _),
                s => $"<w{s}>")
            { Owner = Window.GetWindow(source) };
            if (win.ShowDialog() == true)
                Insert(textBox, richTextBox, useRich, win.ResultTag);
        };
        wMenu.Items.Add(wCustom);
        menu.Items.Add(wMenu);

        // ── <c> ──────────────────────────────────────────
        var cItem = new MenuItem
        {
            Header = "表示クリア <c> を挿入",
            Tag = MenuTag,
            ToolTip = "それまでに表示したテキストを消去"
        };
        cItem.Click += (_, __) => Insert(textBox, richTextBox, useRich, "<c>");
        menu.Items.Add(cItem);

        // ── <p,x,y> ──────────────────────────────────────
        var pItem = new MenuItem
        {
            Header = "座標指定 <p,x,y> を挿入...",
            Tag = MenuTag,
            ToolTip = "テキスト表示座標を変更"
        };
        pItem.Click += (_, __) =>
        {
            var win = new TagInputWindow("座標指定挿入", "X,Y (例: 100,200)", "0,0",
                s =>
                {
                    var p = s.Split(',');
                    return p.Length == 2
                        && double.TryParse(p[0].Trim(),
                            System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out double _)
                        && double.TryParse(p[1].Trim(),
                            System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out double _);
                },
                s => $"<p,{s}>")
            { Owner = Window.GetWindow(source) };
            if (win.ShowDialog() == true)
                Insert(textBox, richTextBox, useRich, win.ResultTag);
        };
        menu.Items.Add(pItem);
    }

    // ── テキスト挿入 ──────────────────────────────────────────

    private static void Insert(
        TextBox textBox, RichTextBox richTextBox,
        bool useRich, string tag)
    {
        if (useRich && richTextBox.Visibility == Visibility.Visible)
            InsertRich(richTextBox, tag);
        else
            InsertText(textBox, tag);
    }

    private static void InsertText(TextBox tb, string tag)
    {
        try
        {
            int pos = tb.SelectionStart;
            tb.Text = tb.Text.Insert(pos, tag);
            tb.CaretIndex = pos + tag.Length;
            tb.Focus();
        }
        catch (Exception ex) { Log($"InsertText: {ex.Message}"); }
    }

    private static void InsertRich(RichTextBox rtb, string tag)
    {
        try
        {
            rtb.BeginChange();
            try
            {
                var pos = rtb.CaretPosition;
                if (pos.Parent is System.Windows.Documents.Run)
                {
                    pos.InsertTextInRun(tag);
                    rtb.CaretPosition = pos.GetPositionAtOffset(tag.Length) ?? pos;
                }
                else
                {
                    var run = new System.Windows.Documents.Run(tag);
                    pos.Paragraph?.Inlines.Add(run);
                    rtb.CaretPosition = run.ElementEnd;
                }
            }
            finally { rtb.EndChange(); }
            rtb.Focus();
        }
        catch (Exception ex) { Log($"InsertRich: {ex.Message}"); }
    }

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss.fff}] {msg}";
        System.Diagnostics.Debug.WriteLine($"[WeightControlPlugin] {msg}");
        try { File.AppendAllText(LogPath, line + "\n"); }
        catch { }
    }
}