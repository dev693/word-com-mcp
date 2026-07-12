using System.Text.Json;
using WordComMcp.Com;
using WordComMcp.Infrastructure;
using WdBuiltinStyle = NetOffice.WordApi.Enums.WdBuiltinStyle;

namespace WordComMcp.Markdown;

/// <summary>The adjusted-Markdown constructs that map to a Word paragraph/character style.</summary>
public enum MarkdownConstruct
{
    Heading1,
    Heading2,
    Heading3,
    Paragraph,
    BulletList,
    NumberedList,
    Quote,
    CodeBlock,
    CodeSpan,
}

/// <summary>
/// Bidirectional map between adjusted-Markdown constructs and Word Formatvorlagen
/// (issue 0.12 / Conventions "interface"). Built-in styles are referenced by their
/// language-independent <see cref="WdBuiltinStyle"/> enum ID (so the map is correct on
/// a German install), while the localized <c>NameLocal</c> defaults let the serializer
/// decide when a paragraph/span needs an explicit <c>{style=…}</c> annotation.
///
/// The default localized names below are the German built-ins plus the custom
/// <c>Code</c>/<c>Zitat</c> Formatvorlagen used by the reference fixture; override any of
/// them from a JSON file via <c>MCP_STYLE_MAP</c> (issue 0.8).
/// </summary>
public sealed class StyleMap
{
    private static readonly IReadOnlyDictionary<MarkdownConstruct, WdBuiltinStyle?> s_builtIns =
        new Dictionary<MarkdownConstruct, WdBuiltinStyle?>
        {
            [MarkdownConstruct.Heading1] = WdBuiltinStyle.wdStyleHeading1,
            [MarkdownConstruct.Heading2] = WdBuiltinStyle.wdStyleHeading2,
            [MarkdownConstruct.Heading3] = WdBuiltinStyle.wdStyleHeading3,
            [MarkdownConstruct.Paragraph] = WdBuiltinStyle.wdStyleNormal,
            [MarkdownConstruct.BulletList] = WdBuiltinStyle.wdStyleListBullet,
            [MarkdownConstruct.NumberedList] = WdBuiltinStyle.wdStyleListNumber,
            [MarkdownConstruct.Quote] = WdBuiltinStyle.wdStyleQuote,
            [MarkdownConstruct.CodeBlock] = null, // custom Formatvorlage
            [MarkdownConstruct.CodeSpan] = null,  // custom Formatvorlage
        };

    private static readonly IReadOnlyDictionary<MarkdownConstruct, string> s_defaultLocalNames =
        new Dictionary<MarkdownConstruct, string>
        {
            [MarkdownConstruct.Heading1] = "Überschrift 1",
            [MarkdownConstruct.Heading2] = "Überschrift 2",
            [MarkdownConstruct.Heading3] = "Überschrift 3",
            [MarkdownConstruct.Paragraph] = "Standard",
            [MarkdownConstruct.BulletList] = "Listenabsatz",
            [MarkdownConstruct.NumberedList] = "Listenabsatz",
            [MarkdownConstruct.Quote] = "Zitat",
            [MarkdownConstruct.CodeBlock] = "Code",
            [MarkdownConstruct.CodeSpan] = "Code",
        };

    private readonly Dictionary<MarkdownConstruct, string> m_localNames;

    private StyleMap(Dictionary<MarkdownConstruct, string> localNames)
    {
        this.m_localNames = localNames;
    }

    /// <summary>The default style map (German built-ins + <c>Code</c>/<c>Zitat</c>).</summary>
    public static StyleMap Default => new(new Dictionary<MarkdownConstruct, string>(s_defaultLocalNames));

    /// <summary>
    /// Load a style map, applying overrides from <paramref name="styleMapPath"/> when set
    /// (issue 0.8). The override JSON maps construct names to localized style names, e.g.
    /// <c>{ "Quote": "Zitat", "CodeBlock": "Quellcode" }</c>.
    /// </summary>
    public static StyleMap Load(string? styleMapPath)
    {
        var names = new Dictionary<MarkdownConstruct, string>(s_defaultLocalNames);
        if (string.IsNullOrWhiteSpace(styleMapPath))
        {
            return new StyleMap(names);
        }

        var json = File.ReadAllText(styleMapPath);
        var overrides = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
        if (overrides is not null)
        {
            foreach (var (key, value) in overrides)
            {
                if (Enum.TryParse<MarkdownConstruct>(key, ignoreCase: true, out var construct) &&
                    !string.IsNullOrWhiteSpace(value))
                {
                    names[construct] = value;
                }
            }
        }

        return new StyleMap(names);
    }

    /// <summary>The language-independent built-in style ID for a construct, or null if it maps to a custom style.</summary>
    public WdBuiltinStyle? BuiltInStyle(MarkdownConstruct construct) => s_builtIns[construct];

    /// <summary>The default localized Formatvorlage name for a construct.</summary>
    public string DefaultLocalName(MarkdownConstruct construct) => this.m_localNames[construct];

    /// <summary>
    /// True when <paramref name="styleName"/> is the default for <paramref name="construct"/>,
    /// i.e. the serializer need <b>not</b> emit an explicit <c>{style=…}</c> annotation.
    /// A null/empty style always counts as the default.
    /// </summary>
    public bool IsDefaultFor(MarkdownConstruct construct, string? styleName) =>
        string.IsNullOrEmpty(styleName) ||
        string.Equals(styleName, this.DefaultLocalName(construct), StringComparison.Ordinal);

    /// <summary>
    /// Validate that <paramref name="styleName"/> exists among <paramref name="availableStyles"/>
    /// (a document's <c>NameLocal</c> values). Throws a structured
    /// <see cref="WordConnectionException"/> listing the available styles when it does not
    /// (issue 0.12 / Conventions Q3 "unknown style name"). Never auto-creates styles.
    /// </summary>
    public static void ValidateStyle(string styleName, IReadOnlyCollection<string> availableStyles)
    {
        ArgumentNullException.ThrowIfNull(styleName);
        ArgumentNullException.ThrowIfNull(availableStyles);

        foreach (var candidate in availableStyles)
        {
            if (string.Equals(candidate, styleName, StringComparison.Ordinal))
            {
                return;
            }
        }

        throw new WordConnectionException(
            McpResult.Errors.UnknownStyle,
            $"Unknown style '{styleName}'. Available styles: {string.Join(", ", availableStyles)}");
    }
}
