using System.Windows;
using System.Windows.Controls;

namespace WeightControlPlugin;

/// <summary>
/// 制御タグのパラメータ入力ダイアログ（XAML なし・コードのみ）
/// </summary>
internal sealed class TagInputWindow : Window
{
    private readonly Func<string, bool>    _validate;
    private readonly Func<string, string>  _buildTag;
    private readonly TextBox               _input;
    private readonly Label                 _preview;
    private readonly Label                 _error;
    private readonly Button                _ok;

    public string ResultTag { get; private set; } = string.Empty;

    public TagInputWindow(
        string title,
        string prompt,
        string defaultValue,
        Func<string, bool>   validate,
        Func<string, string> buildTag)
    {
        _validate = validate;
        _buildTag = buildTag;

        Title         = title;
        Width         = 360;
        Height        = 200;
        ResizeMode    = ResizeMode.NoResize;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;

        var panel = new StackPanel { Margin = new Thickness(12) };

        panel.Children.Add(new Label { Content = prompt });

        _input = new TextBox { Text = defaultValue, Margin = new Thickness(0, 4, 0, 4) };
        _input.TextChanged += OnTextChanged;
        _input.KeyDown += (_, e) =>
        {
            if (e.Key == System.Windows.Input.Key.Return && _ok.IsEnabled)
                Accept();
            else if (e.Key == System.Windows.Input.Key.Escape)
                DialogResult = false;
        };
        panel.Children.Add(_input);

        _preview = new Label { Content = "", Foreground = System.Windows.Media.Brushes.Gray };
        panel.Children.Add(_preview);

        _error = new Label { Content = "", Foreground = System.Windows.Media.Brushes.Red };
        panel.Children.Add(_error);

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 8, 0, 0)
        };
        _ok = new Button
        {
            Content = "OK", Width = 75, IsDefault = true,
            Margin = new Thickness(0, 0, 8, 0)
        };
        _ok.Click += (_, _) => Accept();
        var cancel = new Button { Content = "キャンセル", Width = 75, IsCancel = true };
        buttons.Children.Add(_ok);
        buttons.Children.Add(cancel);
        panel.Children.Add(buttons);

        Content = panel;
        Loaded += (_, _) =>
        {
            _input.Focus();
            _input.SelectAll();
            Validate(_input.Text);
        };
    }

    private void OnTextChanged(object sender, TextChangedEventArgs e) =>
        Validate(_input.Text);

    private void Validate(string text)
    {
        var valid = _validate(text);
        _ok.IsEnabled     = valid;
        _preview.Content  = valid ? $"挿入: {_buildTag(text)}" : "";
        _error.Content    = valid ? "" : "入力値が無効です";
    }

    private void Accept()
    {
        ResultTag    = _buildTag(_input.Text);
        DialogResult = true;
    }
}