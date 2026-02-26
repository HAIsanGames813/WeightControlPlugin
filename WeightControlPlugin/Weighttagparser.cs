using System.Text.RegularExpressions;

namespace WeightControlPlugin;

internal enum TagType
{
    Wait,      // <w秒>
    WaitCoeff, // <w*係数>  ← 後ろの文字数×係数
    Clear,     // <c秒>
    Position,  // <p x,y>  ← 常に相対座標
}

internal record TagToken(int Pos, TagType Type, double Value, double X, double Y);

internal static class WeightTagParser
{
    private static readonly Regex TagRegex = new(
        @"<w\*(?<wcoeff>[\d.]+)>"
        + @"|<w(?<wsec>[\d.]+)>"
        + @"|<c(?<csec>[\d.]+)>"
        + @"|<p(?<px>[+-]?[\d.]+),(?<py>[+-]?[\d.]+)>",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // 色・フォント・速度タグは除去のみ
    private static readonly Regex StripRegex = new(
        @"<#[0-9A-Fa-f]{6}(?:,[0-9A-Fa-f]{6})?>"
        + @"|<#>"
        + @"|<s[\d.,A-Za-z]*>"
        + @"|<s>"
        + @"|<r[\d.]*>"
        + @"|<r>",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static bool HasControlTags(string text) =>
        !string.IsNullOrEmpty(text) && (TagRegex.IsMatch(text) || StripRegex.IsMatch(text));

    public static string RemoveTags(string text) =>
        StripRegex.Replace(TagRegex.Replace(text, ""), "");

    /// <summary>
    /// タグを解析してトークンリストを返す。
    /// Pos は cleanText (改行除く可視文字) 上のインデックス。
    /// </summary>
    public static (string cleanText, List<TagToken> tokens) Parse(string original)
    {
        // まず色・フォント・速度タグを除去
        var stripped = StripRegex.Replace(original, "");

        var tokens = new List<TagToken>();
        int tagOfs = 0;

        var clean = TagRegex.Replace(stripped, m =>
        {
            // raw pos (改行込み) → 可視文字 pos
            int rawPos = m.Index - tagOfs;
            tagOfs += m.Length;
            int visPos = ToVisibleIndex(stripped.Remove(m.Index, m.Length), rawPos);

            var g = m.Groups;
            if (g["wcoeff"].Success)
                tokens.Add(new TagToken(visPos, TagType.WaitCoeff, D(g["wcoeff"].Value), 0, 0));
            else if (g["wsec"].Success)
                tokens.Add(new TagToken(visPos, TagType.Wait, D(g["wsec"].Value), 0, 0));
            else if (g["csec"].Success)
                tokens.Add(new TagToken(visPos, TagType.Clear, D(g["csec"].Value), 0, 0));
            else if (g["px"].Success)
                tokens.Add(new TagToken(visPos, TagType.Position, 0,
                    D(g["px"].Value), D(g["py"].Value)));

            return string.Empty;
        });

        return (clean, tokens);
    }

    /// <summary>改行を除いた可視文字インデックスを返す</summary>
    public static int ToVisibleIndex(string text, int rawPos)
    {
        rawPos = Math.Min(rawPos, text.Length);
        int nl = 0;
        for (int i = 0; i < rawPos; i++)
            if (text[i] == '\r' || text[i] == '\n') nl++;
        return rawPos - nl;
    }

    private static double D(string s) =>
        double.TryParse(s, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;
}