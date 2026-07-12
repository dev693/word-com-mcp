using System.Text;
using WordComMcp.Infrastructure;

namespace WordComMcp.Markdown;

/// <summary>
/// Parser for the adjusted-Markdown subset (issue 0.14): headings h1–h3, paragraphs,
/// bullet/numbered lists, bold/italic/underline, links, simple tables, blockquote, code —
/// plus Pandoc style annotations (<c>::: {style="Zitat"} … :::</c> for paragraphs and
/// <c>[text]{style="Code"}</c> / <c>`x`{style="Code"}</c> for spans). All text is run
/// through <see cref="TextSanitizer"/> so control characters never reach Word.
/// </summary>
public static class MarkdownParser
{
    public static MarkdownDocument Parse(string markdown)
    {
        ArgumentNullException.ThrowIfNull(markdown);

        var lines = TextSanitizer.StripControlChars(markdown).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var blocks = new List<MarkdownBlock>();
        var index = 0;

        while (index < lines.Length)
        {
            var line = lines[index];

            if (string.IsNullOrWhiteSpace(line))
            {
                index++;
                continue;
            }

            if (TryParseStyledDiv(lines, ref index, blocks) ||
                TryParseFencedCode(lines, ref index, blocks) ||
                TryParseHeading(line, ref index, blocks) ||
                TryParseTable(lines, ref index, blocks) ||
                TryParseBlockQuote(lines, ref index, blocks) ||
                TryParseList(lines, ref index, blocks))
            {
                continue;
            }

            ParseParagraph(lines, ref index, blocks);
        }

        return new MarkdownDocument(blocks);
    }

    private static bool TryParseStyledDiv(string[] lines, ref int index, List<MarkdownBlock> blocks)
    {
        var open = lines[index].TrimEnd();
        if (!open.StartsWith(":::", StringComparison.Ordinal))
        {
            return false;
        }

        var style = ExtractDivStyle(open);
        var content = new List<string>();
        var i = index + 1;
        while (i < lines.Length && lines[i].TrimEnd() != ":::")
        {
            content.Add(lines[i]);
            i++;
        }

        // Skip the closing ':::' if present.
        index = i < lines.Length ? i + 1 : i;

        var text = string.Join("\n", content).Trim();
        blocks.Add(new ParagraphBlock(ParseInlines(text), style));
        return true;
    }

    private static string? ExtractDivStyle(string openLine)
    {
        var braceStart = openLine.IndexOf('{');
        if (braceStart < 0)
        {
            return null;
        }

        return ParseStyleAttribute(openLine, braceStart, out _);
    }

    private static bool TryParseFencedCode(string[] lines, ref int index, List<MarkdownBlock> blocks)
    {
        var open = lines[index];
        if (!open.StartsWith("```", StringComparison.Ordinal))
        {
            return false;
        }

        var info = open[3..].Trim();
        var code = new List<string>();
        var i = index + 1;
        while (i < lines.Length && !lines[i].StartsWith("```", StringComparison.Ordinal))
        {
            code.Add(lines[i]);
            i++;
        }

        index = i < lines.Length ? i + 1 : i;
        blocks.Add(new CodeBlock(string.Join("\n", code), string.IsNullOrEmpty(info) ? null : info));
        return true;
    }

    private static bool TryParseHeading(string line, ref int index, List<MarkdownBlock> blocks)
    {
        var level = 0;
        while (level < line.Length && line[level] == '#')
        {
            level++;
        }

        if (level is < 1 or > 3 || level >= line.Length || line[level] != ' ')
        {
            return false;
        }

        var text = line[(level + 1)..].Trim();
        blocks.Add(new HeadingBlock(level, ParseInlines(text)));
        index++;
        return true;
    }

    private static bool TryParseBlockQuote(string[] lines, ref int index, List<MarkdownBlock> blocks)
    {
        if (!lines[index].StartsWith(">", StringComparison.Ordinal))
        {
            return false;
        }

        var content = new List<string>();
        while (index < lines.Length && lines[index].StartsWith(">", StringComparison.Ordinal))
        {
            var stripped = lines[index][1..];
            if (stripped.StartsWith(' '))
            {
                stripped = stripped[1..];
            }

            content.Add(stripped);
            index++;
        }

        blocks.Add(new BlockQuoteBlock(ParseInlines(string.Join("\n", content))));
        return true;
    }

    private static bool TryParseList(string[] lines, ref int index, List<MarkdownBlock> blocks)
    {
        var ordered = IsOrderedItem(lines[index], out _);
        var bullet = IsBulletItem(lines[index]);
        if (!ordered && !bullet)
        {
            return false;
        }

        var items = new List<ListItem>();
        while (index < lines.Length)
        {
            var line = lines[index];
            string? itemText = null;

            if (ordered && IsOrderedItem(line, out itemText))
            {
                // continue
            }
            else if (bullet && IsBulletItem(line))
            {
                itemText = line[2..];
            }
            else
            {
                break;
            }

            items.Add(new ListItem(ParseInlines(itemText!.Trim())));
            index++;
        }

        blocks.Add(new ListBlock(ordered, items));
        return true;
    }

    private static bool IsBulletItem(string line) =>
        line.StartsWith("- ", StringComparison.Ordinal) || line.StartsWith("* ", StringComparison.Ordinal);

    private static bool IsOrderedItem(string line, out string? text)
    {
        text = null;
        var i = 0;
        while (i < line.Length && char.IsDigit(line[i]))
        {
            i++;
        }

        if (i == 0 || i + 1 >= line.Length || line[i] != '.' || line[i + 1] != ' ')
        {
            return false;
        }

        text = line[(i + 2)..];
        return true;
    }

    private static bool TryParseTable(string[] lines, ref int index, List<MarkdownBlock> blocks)
    {
        if (!IsTableRow(lines[index]))
        {
            return false;
        }

        // Require a delimiter row (|---|---|) as the second line to be a table.
        if (index + 1 >= lines.Length || !IsTableDelimiter(lines[index + 1]))
        {
            return false;
        }

        var rows = new List<TableRow>();
        rows.Add(new TableRow(SplitTableCells(lines[index])));
        index += 2; // header + delimiter

        while (index < lines.Length && IsTableRow(lines[index]))
        {
            rows.Add(new TableRow(SplitTableCells(lines[index])));
            index++;
        }

        blocks.Add(new TableBlock(rows));
        return true;
    }

    private static bool IsTableRow(string line) => line.TrimStart().StartsWith('|');

    private static bool IsTableDelimiter(string line)
    {
        var trimmed = line.Trim();
        if (!trimmed.StartsWith('|'))
        {
            return false;
        }

        foreach (var ch in trimmed)
        {
            if (ch is not ('|' or '-' or ':' or ' '))
            {
                return false;
            }
        }

        return trimmed.Contains('-');
    }

    private static List<IReadOnlyList<MarkdownInline>> SplitTableCells(string line)
    {
        var trimmed = line.Trim();
        trimmed = trimmed.Trim('|');
        var cells = new List<IReadOnlyList<MarkdownInline>>();
        foreach (var cell in trimmed.Split('|'))
        {
            cells.Add(ParseInlines(cell.Trim()));
        }

        return cells;
    }

    private static void ParseParagraph(string[] lines, ref int index, List<MarkdownBlock> blocks)
    {
        var content = new List<string>();
        while (index < lines.Length && !string.IsNullOrWhiteSpace(lines[index]) && IsPlainLine(lines[index]))
        {
            content.Add(lines[index]);
            index++;
        }

        if (content.Count == 0)
        {
            // Defensive: consume the line so we never spin.
            content.Add(lines[index]);
            index++;
        }

        blocks.Add(new ParagraphBlock(ParseInlines(string.Join("\n", content))));
    }

    private static bool IsPlainLine(string line)
    {
        var t = line.TrimStart();
        if (t.StartsWith("```", StringComparison.Ordinal) ||
            t.StartsWith(":::", StringComparison.Ordinal) ||
            t.StartsWith('>') ||
            t.StartsWith('|') ||
            IsBulletItem(line) ||
            IsOrderedItem(line, out _))
        {
            return false;
        }

        var hashes = 0;
        while (hashes < line.Length && line[hashes] == '#')
        {
            hashes++;
        }

        return !(hashes is >= 1 and <= 3 && hashes < line.Length && line[hashes] == ' ');
    }

    // ---- inline parsing ------------------------------------------------------

    internal static IReadOnlyList<MarkdownInline> ParseInlines(string text)
    {
        var inlines = new List<MarkdownInline>();
        var buffer = new StringBuilder();
        var i = 0;

        void FlushText()
        {
            if (buffer.Length > 0)
            {
                inlines.Add(new TextInline(buffer.ToString()));
                buffer.Clear();
            }
        }

        while (i < text.Length)
        {
            var ch = text[i];

            if (ch == '`')
            {
                var close = text.IndexOf('`', i + 1);
                if (close > i)
                {
                    FlushText();
                    var code = text[(i + 1)..close];
                    var next = close + 1;
                    var style = TryReadTrailingStyle(text, ref next);
                    inlines.Add(new CodeSpanInline(code, style));
                    i = next;
                    continue;
                }
            }
            else if (ch == '[')
            {
                if (TryParseBracket(text, i, inlines, FlushText, out var consumed))
                {
                    i = consumed;
                    continue;
                }
            }
            else if (ch == '*' && Peek(text, i + 1) == '*')
            {
                if (TryParseDelimited(text, i, "**", out var inner, out var end))
                {
                    FlushText();
                    inlines.Add(new EmphasisInline(EmphasisKind.Bold, inner));
                    i = end;
                    continue;
                }
            }
            else if (ch == '*')
            {
                if (TryParseDelimited(text, i, "*", out var inner, out var end))
                {
                    FlushText();
                    inlines.Add(new EmphasisInline(EmphasisKind.Italic, inner));
                    i = end;
                    continue;
                }
            }
            else if (ch == '_')
            {
                if (TryParseDelimited(text, i, "_", out var inner, out var end))
                {
                    FlushText();
                    inlines.Add(new EmphasisInline(EmphasisKind.Underline, inner));
                    i = end;
                    continue;
                }
            }

            buffer.Append(ch);
            i++;
        }

        FlushText();
        return inlines;
    }

    private static bool TryParseBracket(
        string text, int start, List<MarkdownInline> inlines, Action flushText, out int consumed)
    {
        consumed = start;
        var close = text.IndexOf(']', start + 1);
        if (close < 0)
        {
            return false;
        }

        var label = text[(start + 1)..close];
        var after = close + 1;

        // Link: [text](url)
        if (after < text.Length && text[after] == '(')
        {
            var urlEnd = text.IndexOf(')', after + 1);
            if (urlEnd > after)
            {
                flushText();
                inlines.Add(new LinkInline(label, text[(after + 1)..urlEnd]));
                consumed = urlEnd + 1;
                return true;
            }
        }

        // Styled span: [text]{style="X"}
        if (after < text.Length && text[after] == '{')
        {
            var style = ParseStyleAttribute(text, after, out var attrEnd);
            if (style is not null)
            {
                flushText();
                inlines.Add(new StyledSpanInline(label, style));
                consumed = attrEnd;
                return true;
            }
        }

        return false;
    }

    private static bool TryParseDelimited(string text, int start, string delimiter, out string inner, out int end)
    {
        inner = string.Empty;
        end = start;

        var contentStart = start + delimiter.Length;
        var close = text.IndexOf(delimiter, contentStart, StringComparison.Ordinal);
        if (close < contentStart)
        {
            return false;
        }

        inner = text[contentStart..close];
        if (inner.Length == 0)
        {
            return false;
        }

        end = close + delimiter.Length;
        return true;
    }

    private static string? TryReadTrailingStyle(string text, ref int pos)
    {
        if (pos < text.Length && text[pos] == '{')
        {
            var style = ParseStyleAttribute(text, pos, out var attrEnd);
            if (style is not null)
            {
                pos = attrEnd;
                return style;
            }
        }

        return null;
    }

    /// <summary>
    /// Parse a Pandoc attribute of the form <c>{style="Name"}</c> starting at the '{' in
    /// <paramref name="pos"/>. Returns the style name and sets <paramref name="end"/> to
    /// just past the closing '}', or null when the text there is not such an attribute.
    /// </summary>
    private static string? ParseStyleAttribute(string text, int pos, out int end)
    {
        end = pos;
        const string prefix = "{style=\"";
        if (pos + prefix.Length > text.Length || !text.AsSpan(pos, prefix.Length).SequenceEqual(prefix))
        {
            return null;
        }

        var valueStart = pos + prefix.Length;
        var quoteEnd = text.IndexOf('"', valueStart);
        if (quoteEnd < 0 || quoteEnd + 1 >= text.Length || text[quoteEnd + 1] != '}')
        {
            return null;
        }

        end = quoteEnd + 2;
        return text[valueStart..quoteEnd];
    }

    private static char Peek(string text, int index) => index < text.Length ? text[index] : '\0';
}
