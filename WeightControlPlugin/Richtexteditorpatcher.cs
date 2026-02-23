using HarmonyLib;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;

namespace WeightControlPlugin;

/// <summary>
/// RichTextEditor.OnApplyTemplate() にパッチを当て、
/// PARTS_TextBox / PARTS_RichTextBox の ContextMenu に制御コード挿入メニューを追加します。
/// </summary>
internal static class RichTextEditorPatcher
{
    private const string HarmonyId = "com.weightcontrolplugin.richtexteditor";
    private static Harmony? _harmony;
    private const string MenuTag = "WeightPlugin";

    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<object, object>
        _hooked = new();
    private static readonly object _marker = new();

    public static void Apply()
    {
        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            TryPatch(asm);
        AppDomain.CurrentDomain.AssemblyLoad += (_, args) => TryPatch(args.LoadedAssembly);
    }

    private static void TryPatch(Assembly asm)
    {
        if (_harmony != null) return;

        Type? editorType = null;
        try
        {
            editorType = asm.GetTypes().FirstOrDefault(t =>
                t.Name == "RichTextEditor" &&
                typeof(Control).IsAssignableFrom(t));
        }
        catch { return; }

        if (editorType == null) return;

        var method = editorType.GetMethod("OnApplyTemplate",
            BindingFlags.Instance | BindingFlags.Public);
        if (method == null) return;

        _harmony = new Harmony(HarmonyId);
        _harmony.Patch(method,
            postfix: new HarmonyMethod(typeof(RichTextEditorPatcher),
                nameof(Postfix_OnApplyTemplate)));
    }

    [HarmonyPostfix]
    private static void Postfix_OnApplyTemplate(object __instance)
    {
        if (_hooked.TryGetValue(__instance, out _)) return;
        _hooked.Add(__instance, _marker);

        var getTemplateChild = typeof(FrameworkElement).GetMethod(
            "GetTemplateChild", BindingFlags.Instance | BindingFlags.NonPublic);
        if (getTemplateChild == null) return;

        FrameworkElement? fe = __instance as FrameworkElement;

        void Subscribe()
        {
            var tb  = getTemplateChild.Invoke(__instance, ["PARTS_TextBox"])     as TextBox;
            var rtb = getTemplateChild.Invoke(__instance, ["PARTS_RichTextBox"]) as RichTextBox;
            if (tb == null || rtb == null) return;

            tb.ContextMenuOpening  += (s, _) =>
                InjectMenuItems(s as FrameworkElement, tb, rtb, useRich: false);
            rtb.ContextMenuOpening += (s, _) =>
                InjectMenuItems(s as FrameworkElement, tb, rtb, useRich: true);
        }

        // OnApplyTemplate 直後はまだ null の場合があるので Loaded でも試みる
        Subscribe();
        if (fe != null)
            fe.Loaded += delegate { Subscribe(); };
    }

    private static void InjectMenuItems(
        FrameworkElement? source,
        TextBox tb, RichTextBox rtb,
        bool useRich)
    {
        var menu = source?.ContextMenu;
        if (menu == null) return;

        // 前回分を削除
        menu.Items.OfType<FrameworkElement>()
            .Where(x => x.Tag as string == MenuTag)
            .ToList()
            .ForEach(x => menu.Items.Remove(x));

        menu.Items.Add(new Separator { Tag = MenuTag });

        // ── <w> ──────────────────────────────────────────────
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
            mi.Click += (_, _) => Insert(tb, rtb, useRich, t);
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
            mi.Click += (_, _) => Insert(tb, rtb, useRich, t);
            wMult.Items.Add(mi);
        }
        wMenu.Items.Add(wMult);
        wMenu.Items.Add(new Separator());

        var wCustom = new MenuItem { Header = "カスタム入力..." };
        wCustom.Click += (_, _) =>
        {
            var win = new TagInputWindow("ウェイト挿入", "待機秒数 (例: 1.5)", "1",
                s => double.TryParse(s,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out double _),
                s => $"<w{s}>")
            { Owner = Window.GetWindow(source) };
            if (win.ShowDialog() == true)
                Insert(tb, rtb, useRich, win.ResultTag);
        };
        wMenu.Items.Add(wCustom);
        menu.Items.Add(wMenu);

        // ── <c> ──────────────────────────────────────────────
        var cItem = new MenuItem
        {
            Header = "表示クリア <c> を挿入", Tag = MenuTag,
            ToolTip = "それまでに表示したテキストを消去"
        };
        cItem.Click += (_, _) => Insert(tb, rtb, useRich, "<c>");
        menu.Items.Add(cItem);

        // ── <p,x,y> ──────────────────────────────────────────
        var pItem = new MenuItem
        {
            Header = "座標指定 <p,x,y> を挿入...", Tag = MenuTag,
            ToolTip = "テキスト表示座標を変更"
        };
        pItem.Click += (_, _) =>
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
                Insert(tb, rtb, useRich, win.ResultTag);
        };
        menu.Items.Add(pItem);
    }

    private static void Insert(TextBox tb, RichTextBox rtb, bool useRich, string tag)
    {
        if (useRich && rtb.Visibility == Visibility.Visible)
            InsertRich(rtb, tag);
        else
            InsertText(tb, tag);
    }

    private static void InsertText(TextBox tb, string tag)
    {
        int pos = tb.SelectionStart;
        tb.Text = tb.Text.Insert(pos, tag);
        tb.CaretIndex = pos + tag.Length;
        tb.Focus();
    }

    private static void InsertRich(RichTextBox rtb, string tag)
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
}