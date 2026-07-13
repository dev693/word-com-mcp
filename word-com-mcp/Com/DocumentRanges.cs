using Word = NetOffice.WordApi;

namespace WordComMcp.Com;

/// <summary>
/// Helpers to resolve a live Word <c>Range</c> for a "scope" — currently a heading section
/// (used by <c>get_markdown</c>, <c>replace_text</c>, <c>delete_text</c>). Kept independent of
/// the Markdown layer by taking the localized heading style names as input. STA-thread only.
/// </summary>
public static class DocumentRanges
{
    /// <summary>
    /// Return the range spanning the section that starts at the first heading whose final text
    /// matches <paramref name="anchor"/> (case-insensitive) and runs to the next heading of the
    /// same or higher level (or the document end). Returns <c>null</c> when no such heading exists.
    /// </summary>
    public static Word.Range? ResolveHeadingSection(Word.Document doc, string anchor, IReadOnlyList<string> headingLocalNames)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(anchor);
        ArgumentNullException.ThrowIfNull(headingLocalNames);

        var wanted = anchor.Trim();
        var paragraphs = doc.Paragraphs;
        var count = paragraphs.Count;
        var startPos = -1;
        var startLevel = 0;
        var endPos = doc.Content.End;

        for (var i = 1; i <= count; i++)
        {
            var range = paragraphs[i].Range;
            var level = HeadingLevel(StyleNameOf(range), headingLocalNames);

            if (startPos < 0)
            {
                if (level > 0 &&
                    WordComMcp.Markdown.OoxmlFinalText.GetFinalText(range).Trim()
                        .Equals(wanted, StringComparison.OrdinalIgnoreCase))
                {
                    startPos = range.Start;
                    startLevel = level;
                }
            }
            else if (level > 0 && level <= startLevel)
            {
                endPos = range.Start;
                break;
            }
        }

        return startPos < 0 ? null : doc.Range(startPos, endPos);
    }

    /// <summary>The localized name of a range's paragraph style, or empty when unavailable.</summary>
    public static string StyleNameOf(Word.Range range)
    {
        try
        {
            return range.Style is Word.Style style ? style.NameLocal : string.Empty;
        }
        catch (Exception)
        {
            return string.Empty;
        }
    }

    /// <summary>1-based heading level (1..N) when <paramref name="styleName"/> is a heading style, else 0.</summary>
    public static int HeadingLevel(string styleName, IReadOnlyList<string> headingLocalNames)
    {
        for (var i = 0; i < headingLocalNames.Count; i++)
        {
            if (string.Equals(styleName, headingLocalNames[i], StringComparison.Ordinal))
            {
                return i + 1;
            }
        }

        return 0;
    }
}
