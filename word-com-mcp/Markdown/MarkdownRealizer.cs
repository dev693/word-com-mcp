using WordComMcp.Com;
using WordComMcp.Infrastructure;
using Word = NetOffice.WordApi;
using WdBuiltinStyle = NetOffice.WordApi.Enums.WdBuiltinStyle;
using WdUnderline = NetOffice.WordApi.Enums.WdUnderline;

namespace WordComMcp.Markdown;

/// <summary>
/// Realizes an adjusted-Markdown AST into Word paragraphs at a collapsed insertion point
/// (issue 1.4 <c>insert_markdown</c>). Each block becomes one or more new paragraphs whose
/// Formatvorlage is applied (built-ins by language-independent <see cref="WdBuiltinStyle"/> ID,
/// custom by validated <c>NameLocal</c>); inline bold/italic/underline, character styles, and
/// links are applied over sub-ranges. Lists are realized via built-in list styles as a clean
/// insertion. All text is control-char sanitized; every applied style is read back and any
/// non-stick is reported as a warning. Runs under the caller's UndoRecord/FastEdit/tracking scopes.
/// </summary>
public static class MarkdownRealizer
{
    /// <summary>The outcome of a realize: how many paragraphs were inserted and any soft warnings.</summary>
    public sealed record RealizeResult(int Paragraphs, IReadOnlyList<string> Warnings);

    /// <summary>Insert <paramref name="document"/> at <paramref name="insertionPoint"/> (a collapsed range).</summary>
    public static RealizeResult Realize(
        Word.Document doc,
        Word.Range insertionPoint,
        MarkdownDocument document,
        StyleMap styleMap)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(insertionPoint);
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(styleMap);

        var context = new RealizeContext(doc, styleMap);
        var units = context.BuildUnits(document);

        var cursor = insertionPoint;
        var count = 0;
        foreach (var unit in units)
        {
            context.InsertUnit(ref cursor, unit);
            count++;
        }

        return new RealizeResult(count, context.Warnings);
    }

    /// <summary>
    /// Replace <paramref name="target"/>'s content with <paramref name="replacementMarkdown"/> as a
    /// minimal redline (issue 1.5). When tracking is on, setting the range text yields a tracked
    /// delete+insert over just this span. Inline emphasis / <c>`code`</c> / <c>{style=…}</c> spans
    /// in the replacement are applied over the inserted text. Returns any soft warnings.
    /// </summary>
    public static IReadOnlyList<string> ReplaceRange(
        Word.Document doc,
        Word.Range target,
        string replacementMarkdown,
        StyleMap styleMap)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(styleMap);

        var context = new RealizeContext(doc, styleMap);
        var inlines = MarkdownParser.ParseInlines(replacementMarkdown ?? string.Empty);
        var (text, spans) = context.BuildInlineText(inlines);

        // Replace the matched span only — a minimal ins/del redline under tracking.
        target.Text = text;

        if (spans.Count > 0)
        {
            context.ApplySpans(target.Start, spans);
        }

        return context.Warnings;
    }

    /// <summary>
    /// Apply the paragraph and inline formatting represented by a single Markdown block without
    /// replacing its text. Used by Tier-2 style-only diffs so Word records a formatting revision.
    /// </summary>
    public static IReadOnlyList<string> ApplyBlockFormatting(
        Word.Document doc,
        Word.Range target,
        MarkdownBlock block,
        StyleMap styleMap)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(block);
        ArgumentNullException.ThrowIfNull(styleMap);

        var context = new RealizeContext(doc, styleMap);
        var units = context.BuildUnits(new MarkdownDocument([block]));
        if (units.Count != 1)
        {
            return ["block formatting could not be mapped to one paragraph"];
        }

        var unit = units[0];
        var content = doc.Range(target.Start, Math.Max(target.Start, target.End - 1));
        context.ApplyParagraphStyle(content, unit);

        // Remove prior inline formatting/hyperlinks, then realize the requested spans in place.
        var hyperlinks = content.Hyperlinks;
        for (var i = hyperlinks.Count; i >= 1; i--)
        {
            hyperlinks[i].Delete();
        }

        content.Style = doc.Styles[WdBuiltinStyle.wdStyleDefaultParagraphFont];
        content.Font.Bold = 0;
        content.Font.Italic = 0;
        content.Font.Underline = WdUnderline.wdUnderlineNone;
        context.ApplySpans(content.Start, unit.Spans);
        return context.Warnings;
    }

    private enum SpanKind
    {
        Bold,
        Italic,
        Underline,
        CharacterStyle,
        Link,
    }

    private readonly record struct InlineSpan(int Offset, int Length, SpanKind Kind, string? Style, string? Url);

    private sealed record ParaUnit(string Text, WdBuiltinStyle? BuiltIn, string? CustomStyle, IReadOnlyList<InlineSpan> Spans);

    private sealed class RealizeContext(Word.Document doc, StyleMap styleMap)
    {
        private readonly List<string> m_warnings = new();
        private IReadOnlyList<string>? m_availableStyles;

        public IReadOnlyList<string> Warnings => this.m_warnings;

        public IReadOnlyList<ParaUnit> BuildUnits(MarkdownDocument document)
        {
            var units = new List<ParaUnit>();
            foreach (var block in document.Blocks)
            {
                this.AppendBlock(units, block);
            }

            return units;
        }

        private void AppendBlock(List<ParaUnit> units, MarkdownBlock block)
        {
            switch (block)
            {
                case HeadingBlock heading:
                    units.Add(this.InlineUnit(heading.Inlines, HeadingBuiltIn(heading.Level), custom: null));
                    break;

                case ParagraphBlock paragraph:
                    if (styleMap.IsDefaultFor(MarkdownConstruct.Paragraph, paragraph.Style))
                    {
                        units.Add(this.InlineUnit(paragraph.Inlines, WdBuiltinStyle.wdStyleNormal, custom: null));
                    }
                    else
                    {
                        units.Add(this.InlineUnit(paragraph.Inlines, builtIn: null, custom: paragraph.Style));
                    }

                    break;

                case BlockQuoteBlock quote:
                    units.Add(this.InlineUnit(quote.Inlines, WdBuiltinStyle.wdStyleQuote, custom: null));
                    break;

                case CodeBlock code:
                    var codeStyle = styleMap.DefaultLocalName(MarkdownConstruct.CodeBlock);
                    foreach (var line in SplitLines(code.Code))
                    {
                        units.Add(new ParaUnit(TextSanitizer.StripControlChars(line) ?? string.Empty, null, codeStyle, Array.Empty<InlineSpan>()));
                    }

                    break;

                case ListBlock list:
                    var listStyle = list.Ordered ? WdBuiltinStyle.wdStyleListNumber : WdBuiltinStyle.wdStyleListBullet;
                    foreach (var item in list.Items)
                    {
                        units.Add(this.InlineUnit(item.Inlines, listStyle, custom: null));
                    }

                    break;

                case TableBlock:
                    this.m_warnings.Add("table insertion is not supported in Tier 1; the table was skipped");
                    break;
            }
        }

        private ParaUnit InlineUnit(IReadOnlyList<MarkdownInline> inlines, WdBuiltinStyle? builtIn, string? custom)
        {
            var (text, spans) = this.BuildInlineText(inlines);
            return new ParaUnit(text, builtIn, custom, spans);
        }

        public (string Text, IReadOnlyList<InlineSpan> Spans) BuildInlineText(IReadOnlyList<MarkdownInline> inlines)
        {
            var builder = new System.Text.StringBuilder();
            var spans = new List<InlineSpan>();

            foreach (var inline in inlines)
            {
                switch (inline)
                {
                    case TextInline text:
                        builder.Append(Clean(text.Text));
                        break;

                    case EmphasisInline emphasis:
                        AppendSpan(builder, spans, Clean(emphasis.Text), EmphasisSpan(emphasis.Kind), style: null, url: null);
                        break;

                    case CodeSpanInline code:
                        var codeStyle = code.Style ?? styleMap.DefaultLocalName(MarkdownConstruct.CodeSpan);
                        AppendSpan(builder, spans, Clean(code.Code), SpanKind.CharacterStyle, codeStyle, url: null);
                        break;

                    case StyledSpanInline styled:
                        AppendSpan(builder, spans, Clean(styled.Text), SpanKind.CharacterStyle, styled.Style, url: null);
                        break;

                    case LinkInline link:
                        AppendSpan(builder, spans, Clean(link.Text), SpanKind.Link, style: null, link.Url);
                        break;
                }
            }

            return (builder.ToString(), spans);
        }

        private static void AppendSpan(
            System.Text.StringBuilder builder,
            List<InlineSpan> spans,
            string text,
            SpanKind kind,
            string? style,
            string? url)
        {
            if (text.Length == 0)
            {
                return;
            }

            spans.Add(new InlineSpan(builder.Length, text.Length, kind, style, url));
            builder.Append(text);
        }

        public void InsertUnit(ref Word.Range cursor, ParaUnit unit)
        {
            var startPos = cursor.End;

            // Insert the paragraph text followed by a paragraph mark (a tracked insertion when tracking is on).
            cursor.InsertAfter(unit.Text + "\r");

            var paraRange = doc.Range(startPos, startPos + unit.Text.Length);
            this.ApplyParagraphStyle(paraRange, unit);
            this.ApplySpans(startPos, unit.Spans);

            // Advance past the inserted paragraph mark.
            cursor.Start = cursor.End;
        }

        public void ApplyParagraphStyle(Word.Range range, ParaUnit unit)
        {
            if (unit.BuiltIn is WdBuiltinStyle builtIn)
            {
                var style = doc.Styles[builtIn];
                range.Style = style;
                this.VerifyStyle(range, style.NameLocal);
            }
            else if (!string.IsNullOrEmpty(unit.CustomStyle))
            {
                StyleMap.ValidateStyle(unit.CustomStyle, this.AvailableStyles());
                range.Style = unit.CustomStyle;
                this.VerifyStyle(range, unit.CustomStyle);
            }
        }

        public void ApplySpans(int paragraphStart, IReadOnlyList<InlineSpan> spans)
        {
            foreach (var span in spans)
            {
                var range = doc.Range(paragraphStart + span.Offset, paragraphStart + span.Offset + span.Length);
                switch (span.Kind)
                {
                    case SpanKind.Bold:
                        range.Font.Bold = 1;
                        break;
                    case SpanKind.Italic:
                        range.Font.Italic = 1;
                        break;
                    case SpanKind.Underline:
                        range.Font.Underline = WdUnderline.wdUnderlineSingle;
                        break;
                    case SpanKind.CharacterStyle when !string.IsNullOrEmpty(span.Style):
                        StyleMap.ValidateStyle(span.Style, this.AvailableStyles());
                        range.Style = span.Style;
                        this.VerifyStyle(range, span.Style);
                        break;
                    case SpanKind.Link when !string.IsNullOrEmpty(span.Url):
                        this.AddHyperlink(range, span.Url);
                        break;
                }
            }
        }

        private void AddHyperlink(Word.Range range, string url)
        {
            try
            {
                doc.Hyperlinks.Add(range, url);
            }
            catch (Exception)
            {
                this.m_warnings.Add($"could not create hyperlink to '{url}'");
            }
        }

        private void VerifyStyle(Word.Range range, string expected)
        {
            try
            {
                var actual = range.Style is Word.Style style ? style.NameLocal : string.Empty;
                if (!string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    this.m_warnings.Add($"style '{expected}' did not apply (got '{actual}')");
                }
            }
            catch (Exception)
            {
                // Reading back a style must never fail the insert.
            }
        }

        private IReadOnlyList<string> AvailableStyles() =>
            this.m_availableStyles ??= DocumentStyles.LocalNames(doc);

        private static string Clean(string? text) => TextSanitizer.StripControlChars(text) ?? string.Empty;

        private static SpanKind EmphasisSpan(EmphasisKind kind) => kind switch
        {
            EmphasisKind.Bold => SpanKind.Bold,
            EmphasisKind.Italic => SpanKind.Italic,
            EmphasisKind.Underline => SpanKind.Underline,
            _ => SpanKind.Bold,
        };

        private static WdBuiltinStyle HeadingBuiltIn(int level) => level switch
        {
            1 => WdBuiltinStyle.wdStyleHeading1,
            2 => WdBuiltinStyle.wdStyleHeading2,
            3 => WdBuiltinStyle.wdStyleHeading3,
            _ => WdBuiltinStyle.wdStyleHeading3,
        };

        private static IEnumerable<string> SplitLines(string text) =>
            text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
    }
}
