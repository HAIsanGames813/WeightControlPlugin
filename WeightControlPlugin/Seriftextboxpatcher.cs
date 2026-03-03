using HarmonyLib;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;

namespace WeightControlPlugin;

/// <summary>
/// VoiceItem のセリフ入力欄（SerifTextBox）に制御コード挿入メニューを追加する。
/// SerifTextBox は TextBox を継承した YukkuriMovieMaker.Controls 内の public class。
/// </summary>
internal static class SerifTextBoxPatcher
{
    private const string HarmonyId = "com.weightcontrolplugin.seriftextbox";
    private const string MenuTag = "WeightPluginSerif";

    private static Harmony? _harmony;

    // パッチ済みインスタンスを追跡（二重フックを防ぐ）
    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<object, object>
        _hooked = new();
    private static readonly object _marker = new();

    public static void Apply()
    {
        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            TryPatch(asm);
        AppDomain.CurrentDomain.AssemblyLoad += (_, args) =>
        {
            try { TryPatch(args.LoadedAssembly); } catch { }
        };
    }

    private static void TryPatch(Assembly asm)
    {
        if (_harmony != null) return;

        // SerifTextBox を検索（YukkuriMovieMaker.Controls.dll にある）
        Type? stbType = null;
        try
        {
            stbType = asm.GetTypes()
                .FirstOrDefault(t => t.Name == "SerifTextBox"
                                  && typeof(Control).IsAssignableFrom(t));
        }
        catch { return; }

        if (stbType == null) return;

        // OnApplyTemplate をパッチ
        var method = stbType.GetMethod("OnApplyTemplate",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
        if (method == null) return;

        try
        {
            _harmony = new Harmony(HarmonyId);
            _harmony.Patch(method,
                postfix: new HarmonyMethod(
                    typeof(SerifTextBoxPatcher),
                    nameof(Postfix_OnApplyTemplate)));
        }
        catch
        {
            _harmony = null;
        }
    }

    [HarmonyPostfix]
    private static void Postfix_OnApplyTemplate(object __instance)
    {
        try
        {
            if (_hooked.TryGetValue(__instance, out _)) return;
            _hooked.Add(__instance, _marker);

            // SerifTextBox は TextBox 直接継承の可能性が高い
            if (__instance is TextBox tb)
            {
                tb.ContextMenuOpening += (_, _) => InjectMenu(tb);
                return;
            }

            // Control 継承の場合: テンプレート内の TextBox を探す
            var fe = __instance as FrameworkElement;
            if (fe == null) return;

            var getTemplateChild = typeof(FrameworkElement).GetMethod(
                "GetTemplateChild", BindingFlags.Instance | BindingFlags.NonPublic);

            void TryHook()
            {
                foreach (var name in new[] {
                    "PARTS_TextBox", "PART_TextBox", "TextBox",
                    "PARTS_RichTextBox", "PART_RichTextBox" })
                {
                    try
                    {
                        var child = getTemplateChild?.Invoke(__instance, new object[] { name });
                        if (child is TextBox t2)
                        {
                            t2.ContextMenuOpening += (_, _) => InjectMenu(t2);
                            return;
                        }
                    }
                    catch { }
                }
            }

            TryHook();
            fe.Loaded += delegate
            {
                try { TryHook(); } catch { }
            };
        }
        catch { }
    }

    private static void InjectMenu(TextBox tb)
    {
        try
        {
            var menu = tb.ContextMenu;
            if (menu == null) return;

            // 前回分を削除
            var toRemove = menu.Items.OfType<FrameworkElement>()
                .Where(x => x.Tag as string == MenuTag)
                .ToList();
            foreach (var item in toRemove)
                menu.Items.Remove(item);

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
                mi.Click += (_, _) => Insert(tb, t);
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
                mi.Click += (_, _) => Insert(tb, t);
                wMult.Items.Add(mi);
            }
            wMenu.Items.Add(wMult);
            wMenu.Items.Add(new Separator());

            var wCustom = new MenuItem { Header = "カスタム入力..." };
            wCustom.Click += (_, _) =>
            {
                try
                {
                    var win = new TagInputWindow(
                        "ウェイト挿入", "待機秒数 (例: 1.5)", "1",
                        s => double.TryParse(s,
                            System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out _),
                        s => $"<w{s}>")
                    { Owner = Window.GetWindow(tb) };
                    if (win.ShowDialog() == true) Insert(tb, win.ResultTag);
                }
                catch { }
            };
            wMenu.Items.Add(wCustom);
            menu.Items.Add(wMenu);

            // ── <c> ──────────────────────────────────────────────
            var cItem = new MenuItem
            {
                Header = "表示クリア <c> を挿入",
                Tag = MenuTag,
                ToolTip = "それまでに表示したテキストを消去"
            };
            cItem.Click += (_, _) => Insert(tb, "<c>");
            menu.Items.Add(cItem);

            // ── <p,x,y> ──────────────────────────────────────────
            var pItem = new MenuItem
            {
                Header = "座標指定 <p,x,y> を挿入...",
                Tag = MenuTag,
                ToolTip = "テキスト表示座標を変更"
            };
            pItem.Click += (_, _) =>
            {
                try
                {
                    var win = new TagInputWindow(
                        "座標指定挿入", "X,Y (例: 100,200)", "0,0",
                        s =>
                        {
                            var p = s.Split(',');
                            return p.Length == 2
                                && double.TryParse(p[0].Trim(),
                                    System.Globalization.NumberStyles.Float,
                                    System.Globalization.CultureInfo.InvariantCulture, out _)
                                && double.TryParse(p[1].Trim(),
                                    System.Globalization.NumberStyles.Float,
                                    System.Globalization.CultureInfo.InvariantCulture, out _);
                        },
                        s => $"<p,{s}>")
                    { Owner = Window.GetWindow(tb) };
                    if (win.ShowDialog() == true) Insert(tb, win.ResultTag);
                }
                catch { }
            };
            menu.Items.Add(pItem);
        }
        catch { }
    }

    private static void Insert(TextBox tb, string tag)
    {
        try
        {
            int pos = tb.SelectionStart;
            tb.Text = tb.Text.Insert(pos, tag);
            tb.CaretIndex = pos + tag.Length;
            tb.Focus();
        }
        catch { }
    }
}