using System.Windows;
using System.Windows.Controls;

namespace WeightControlPlugin;

/// <summary>
/// 制御コードのパラメータを入力するための汎用シンプルダイアログ。
/// XAML なしで完全にコードのみで構築します。
/// </summary>
internal class TagInputWindow : Window
{
    private readonly System.Func<string, bool> _validate;
    private readonly System.Func<string, string> _buildTag;
    private readonly TextBox _inputBox;
    private readonly Button _okButton;
    private readonly TextBlock _errorLabel;

    /// <summary>確定されたタグ文字列。ShowDialog() が true を返した後に取得できます。</summary>
    public string ResultTag { get; private set; } = string.Empty;

    /// <param name="title">ウィンドウタイトル</param>
    /// <param name="message">入力欄上に表示するメッセージ</param>
    /// <param name="defaultValue">入力欄の初期値</param>
    /// <param name="validate">入力値が正しいか検証する関数</param>
    /// <param name="buildTag">入力値からタグ文字列を生成する関数</param>
    public TagInputWindow(
        string title,
        string message,
        string defaultValue,
        System.Func<string, bool> validate,
        System.Func<string, string> buildTag)
    {
        _validate = validate;
        _buildTag = buildTag;

        // ── ウィンドウ設定 ────────────────────────────────────
        Title = title;
        Width = 400;
        SizeToContent = SizeToContent.Height;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        ShowInTaskbar = false;

        // ── レイアウト ────────────────────────────────────────
        var panel = new StackPanel { Margin = new Thickness(16) };

        // 説明ラベル
        panel.Children.Add(new TextBlock
        {
            Text = message,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8)
        });

        // 入力欄
        _inputBox = new TextBox
        {
            Text = defaultValue,
            Margin = new Thickness(0, 0, 0, 4)
        };
        _inputBox.TextChanged += OnTextChanged;
        _inputBox.KeyDown += OnKeyDown;
        panel.Children.Add(_inputBox);

        // エラーラベル
        _errorLabel = new TextBlock
        {
            Foreground = System.Windows.Media.Brushes.Red,
            FontSize = 11,
            Margin = new Thickness(0, 0, 0, 8),
            Visibility = Visibility.Collapsed
        };
        panel.Children.Add(_errorLabel);

        // プレビューラベル
        var previewHeader = new TextBlock
        {
            Text = "プレビュー:",
            FontSize = 11,
            Foreground = System.Windows.Media.Brushes.Gray,
            Margin = new Thickness(0, 0, 0, 2)
        };
        panel.Children.Add(previewHeader);

        var previewBox = new TextBlock
        {
            FontFamily = new System.Windows.Media.FontFamily("Consolas, Courier New, monospace"),
            FontSize = 13,
            Margin = new Thickness(0, 0, 0, 12),
            Name = "PreviewBox"
        };
        panel.Children.Add(previewBox);
        // プレビュー更新用に参照を保持
        _inputBox.TextChanged += (_, __) =>
        {
            previewBox.Text = _validate(_inputBox.Text)
                ? _buildTag(_inputBox.Text)
                : "(無効な値)";
        };

        // ボタン行
        var btnRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };

        _okButton = new Button
        {
            Content = "挿入",
            Width = 80,
            Margin = new Thickness(0, 0, 8, 0),
            IsDefault = true
        };
        _okButton.Click += OnOkClick;
        btnRow.Children.Add(_okButton);

        var cancelButton = new Button
        {
            Content = "キャンセル",
            Width = 80,
            IsCancel = true
        };
        cancelButton.Click += (_, __) => DialogResult = false;
        btnRow.Children.Add(cancelButton);

        panel.Children.Add(btnRow);

        Content = panel;

        // 初期値でバリデーション実行
        Loaded += (_, __) =>
        {
            _inputBox.SelectAll();
            _inputBox.Focus();
            UpdateOkButton();
            previewBox.Text = _validate(defaultValue) ? _buildTag(defaultValue) : "(無効な値)";
        };
    }

    private void OnTextChanged(object sender, TextChangedEventArgs e) => UpdateOkButton();

    private void UpdateOkButton()
    {
        bool valid = _validate(_inputBox.Text);
        _okButton.IsEnabled = valid;

        if (!valid && !string.IsNullOrEmpty(_inputBox.Text))
        {
            _errorLabel.Text = "入力値が正しくありません";
            _errorLabel.Visibility = Visibility.Visible;
        }
        else
        {
            _errorLabel.Visibility = Visibility.Collapsed;
        }
    }

    private void OnKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Enter && _okButton.IsEnabled)
            OnOkClick(sender, e);
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        if (!_validate(_inputBox.Text)) return;
        ResultTag = _buildTag(_inputBox.Text);
        DialogResult = true;
    }
}