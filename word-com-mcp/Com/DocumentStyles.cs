using Word = NetOffice.WordApi;
using WdStyleType = NetOffice.WordApi.Enums.WdStyleType;

namespace WordComMcp.Com;

/// <summary>
/// One entry of a document's style catalogue (issue 1.1 <c>list_styles</c>). Word's object
/// model exposes only the localized name (<c>NameLocal</c>) on a <c>Style</c>, so
/// <see cref="Name"/> mirrors <see cref="NameLocal"/> for schema compatibility.
/// </summary>
public sealed record StyleInfo(string Name, string NameLocal, string Type, bool BuiltIn, bool InUse);

/// <summary>
/// Enumerates a live document's <c>Styles</c> collection (issue 1.1). There is no Phase-0
/// live-style reader; this backs <c>word_live_list_styles</c> and supplies the
/// <c>NameLocal</c> list that <c>StyleMap.ValidateStyle</c> checks against before a tool
/// applies a custom Formatvorlage. Must run on the STA thread (<see cref="StaDispatcher"/>).
/// </summary>
public static class DocumentStyles
{
    /// <summary>Read every style as <see cref="StyleInfo"/> (name, localized name, type, built-in, in-use).</summary>
    public static IReadOnlyList<StyleInfo> Enumerate(Word.Document doc)
    {
        ArgumentNullException.ThrowIfNull(doc);

        var result = new List<StyleInfo>();
        var styles = doc.Styles;
        var count = styles.Count;
        for (var i = 1; i <= count; i++)
        {
            var style = styles[i];
            try
            {
                var nameLocal = style.NameLocal;
                result.Add(new StyleInfo(
                    Name: nameLocal,
                    NameLocal: nameLocal,
                    Type: MapType(style.Type),
                    BuiltIn: style.BuiltIn,
                    InUse: style.InUse));
            }
            finally
            {
                style.Dispose();
            }
        }

        styles.Dispose();
        return result;
    }

    /// <summary>The document's localized style names — the <c>availableStyles</c> for style validation.</summary>
    public static IReadOnlyList<string> LocalNames(Word.Document doc)
    {
        ArgumentNullException.ThrowIfNull(doc);

        var names = new List<string>();
        var styles = doc.Styles;
        var count = styles.Count;
        for (var i = 1; i <= count; i++)
        {
            var style = styles[i];
            try
            {
                names.Add(style.NameLocal);
            }
            finally
            {
                style.Dispose();
            }
        }

        styles.Dispose();
        return names;
    }

    private static string MapType(WdStyleType type) => type switch
    {
        WdStyleType.wdStyleTypeParagraph => "paragraph",
        WdStyleType.wdStyleTypeParagraphOnly => "paragraph",
        WdStyleType.wdStyleTypeCharacter => "character",
        WdStyleType.wdStyleTypeLinked => "linked",
        WdStyleType.wdStyleTypeTable => "table",
        WdStyleType.wdStyleTypeList => "list",
        _ => "paragraph",
    };
}
