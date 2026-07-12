using System.Text;

namespace WordComMcp.Infrastructure;

/// <summary>
/// Control-char safety + newline normalization for all Markdown-derived text
/// (Conventions: "Control-char safety" and the reference's newline-&gt;\r rule).
///
/// Stray control characters (e.g. the <c>\x07</c> cell mark) corrupt Word's
/// Find/Replace and have historically lost documents, so they are stripped before
/// any insert/find. Word's paragraph mark is a bare CR (<c>\r</c>), so incoming
/// <c>\r\n</c>/<c>\n</c> are normalized to <c>\r</c>.
/// </summary>
public static class TextSanitizer
{
    /// <summary>
    /// Remove control characters that must never reach Word. Tab (<c>\t</c>) and
    /// the CR paragraph mark (<c>\r</c>) are preserved; everything else below
    /// U+0020 (plus the DEL range) is dropped.
    /// </summary>
    public static string StripControlChars(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text ?? string.Empty;
        }

        var builder = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            if (IsAllowed(ch))
            {
                builder.Append(ch);
            }
        }

        return builder.ToString();
    }

    /// <summary>Normalize <c>\r\n</c> and lone <c>\n</c> to Word's paragraph mark <c>\r</c>.</summary>
    public static string NormalizeNewlines(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text ?? string.Empty;
        }

        return text.Replace("\r\n", "\r").Replace('\n', '\r');
    }

    /// <summary>Strip control chars then normalize newlines — the standard pre-insert pipeline.</summary>
    public static string Sanitize(string? text) => NormalizeNewlines(StripControlChars(text));

    private static bool IsAllowed(char ch)
    {
        if (ch == '\t' || ch == '\r' || ch == '\n')
        {
            return true;
        }

        // Drop every control char: C0 (U+0000–U+001F), DEL (U+007F), and C1 (U+0080–U+009F).
        // char.IsControl covers all three ranges; \t \r \n were already allowed above.
        return !char.IsControl(ch);
    }
}
