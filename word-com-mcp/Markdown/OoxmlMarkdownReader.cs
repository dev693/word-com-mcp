using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OoxmlTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace WordComMcp.Markdown;

/// <summary>
/// Reads a Word range's OOXML (the shape <c>Range.WordOpenXML</c> returns) into the
/// adjusted-Markdown AST (issue 1.3 <c>get_markdown</c>). Revisions are resolved with the
/// shared <see cref="OoxmlRevisions"/> walk (default <see cref="RevisionView.Final"/>), paragraph
/// styles are classified against the <see cref="StyleMap"/> (custom Formatvorlagen surface as
/// <c>{style=…}</c> after serialization), and inline runs are grouped by direct formatting
/// (bold/italic/underline) and character style. Word-free and unit-testable from Flat-OPC.
/// </summary>
public static class OoxmlMarkdownReader
{
    /// <summary>Read from a Flat-OPC XML string (from <c>Range.WordOpenXML</c>).</summary>
    public static MarkdownDocument Read(string wordOpenXml, StyleMap styleMap, RevisionView view = RevisionView.Final)
    {
        ArgumentException.ThrowIfNullOrEmpty(wordOpenXml);
        ArgumentNullException.ThrowIfNull(styleMap);

        using var doc = WordprocessingDocument.FromFlatOpcString(wordOpenXml);
        return ReadDocument(doc, styleMap, view);
    }

    /// <summary>Read from a full <c>.docx</c> package (byte array). Test-friendly.</summary>
    public static MarkdownDocument ReadFromPackage(byte[] docxBytes, StyleMap styleMap, RevisionView view = RevisionView.Final)
    {
        ArgumentNullException.ThrowIfNull(docxBytes);
        ArgumentNullException.ThrowIfNull(styleMap);

        using var stream = new MemoryStream();
        stream.Write(docxBytes, 0, docxBytes.Length);
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, isEditable: true);
        return ReadDocument(doc, styleMap, view);
    }

    private static MarkdownDocument ReadDocument(WordprocessingDocument doc, StyleMap styleMap, RevisionView view)
    {
        var mainPart = doc.MainDocumentPart;
        var body = mainPart?.Document?.Body;
        if (mainPart is null || body is null)
        {
            return new MarkdownDocument(Array.Empty<MarkdownBlock>());
        }

        var context = new ReadContext(styleMap, BuildStyleNameMap(mainPart), mainPart);
        OoxmlRevisions.Apply(body, view);

        var blocks = new List<MarkdownBlock>();
        List<ListItem>? pendingList = null;
        var pendingOrdered = false;

        foreach (var element in body.Elements())
        {
            if (element is Paragraph paragraph)
            {
                var classified = context.Classify(paragraph);

                if (classified.Kind == ParagraphKind.ListItem)
                {
                    pendingList ??= new List<ListItem>();
                    if (pendingList.Count == 0)
                    {
                        pendingOrdered = classified.Ordered;
                    }

                    pendingList.Add(new ListItem(context.ReadInlines(paragraph)));
                    continue;
                }

                FlushList(blocks, ref pendingList, pendingOrdered);

                var block = context.BuildBlock(paragraph, classified);
                if (block is not null)
                {
                    blocks.Add(block);
                }
            }
            else if (element is Table table)
            {
                FlushList(blocks, ref pendingList, pendingOrdered);
                blocks.Add(context.ReadTable(table));
            }
        }

        FlushList(blocks, ref pendingList, pendingOrdered);
        return new MarkdownDocument(blocks);
    }

    private static void FlushList(List<MarkdownBlock> blocks, ref List<ListItem>? pending, bool ordered)
    {
        if (pending is { Count: > 0 })
        {
            blocks.Add(new ListBlock(ordered, pending));
        }

        pending = null;
    }

    /// <summary>Map each style's <c>styleId</c> to its localized <c>w:name</c> for classification.</summary>
    private static IReadOnlyDictionary<string, string> BuildStyleNameMap(MainDocumentPart mainPart)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        var styles = mainPart.StyleDefinitionsPart?.Styles;
        if (styles is null)
        {
            return map;
        }

        foreach (var style in styles.Elements<Style>())
        {
            var id = style.StyleId?.Value;
            var name = style.StyleName?.Val?.Value;
            if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(name))
            {
                map[id] = name;
            }
        }

        return map;
    }

    private enum ParagraphKind
    {
        Heading,
        Quote,
        CodeBlock,
        ListItem,
        Paragraph,
    }

    private readonly record struct Classification(ParagraphKind Kind, int HeadingLevel, bool Ordered, string? StyleName);

    private sealed class ReadContext(StyleMap styleMap, IReadOnlyDictionary<string, string> styleNames, MainDocumentPart mainPart)
    {
        public Classification Classify(Paragraph paragraph)
        {
            var pPr = paragraph.ParagraphProperties;
            var styleId = pPr?.ParagraphStyleId?.Val?.Value;
            var local = ResolveStyleName(styleId);
            var hasNumbering = pPr?.NumberingProperties is not null;

            var bulletDefault = styleMap.DefaultLocalName(MarkdownConstruct.BulletList);
            if (hasNumbering || NameEquals(local, bulletDefault))
            {
                return new Classification(ParagraphKind.ListItem, 0, ResolveOrdered(pPr), null);
            }

            if (NameEquals(local, styleMap.DefaultLocalName(MarkdownConstruct.Heading1)))
            {
                return new Classification(ParagraphKind.Heading, 1, false, null);
            }

            if (NameEquals(local, styleMap.DefaultLocalName(MarkdownConstruct.Heading2)))
            {
                return new Classification(ParagraphKind.Heading, 2, false, null);
            }

            if (NameEquals(local, styleMap.DefaultLocalName(MarkdownConstruct.Heading3)))
            {
                return new Classification(ParagraphKind.Heading, 3, false, null);
            }

            if (NameEquals(local, styleMap.DefaultLocalName(MarkdownConstruct.Quote)))
            {
                return new Classification(ParagraphKind.Quote, 0, false, null);
            }

            if (NameEquals(local, styleMap.DefaultLocalName(MarkdownConstruct.CodeBlock)))
            {
                return new Classification(ParagraphKind.CodeBlock, 0, false, null);
            }

            // Default paragraph or a custom paragraph Formatvorlage → emit {style=…} when non-default.
            var styleName = styleMap.IsDefaultFor(MarkdownConstruct.Paragraph, local) ? null : local;
            return new Classification(ParagraphKind.Paragraph, 0, false, styleName);
        }

        public MarkdownBlock? BuildBlock(Paragraph paragraph, Classification classified)
        {
            if (classified.Kind == ParagraphKind.CodeBlock)
            {
                return new CodeBlock(ReadPlainText(paragraph));
            }

            var inlines = this.ReadInlines(paragraph);
            return classified.Kind switch
            {
                ParagraphKind.Heading => new HeadingBlock(classified.HeadingLevel, inlines),
                ParagraphKind.Quote => new BlockQuoteBlock(inlines),
                _ => IsEmpty(inlines) ? null : new ParagraphBlock(inlines, classified.StyleName),
            };
        }

        public IReadOnlyList<MarkdownInline> ReadInlines(Paragraph paragraph)
        {
            var inlines = new List<MarkdownInline>();
            foreach (var element in paragraph.Elements())
            {
                switch (element)
                {
                    case Run run:
                        AppendRun(inlines, run);
                        break;
                    case Hyperlink hyperlink:
                        this.AppendHyperlink(inlines, hyperlink);
                        break;
                }
            }

            return Coalesce(inlines);
        }

        public TableBlock ReadTable(Table table)
        {
            var rows = new List<TableRow>();
            foreach (var row in table.Elements<OoxmlTableRow>())
            {
                var cells = new List<IReadOnlyList<MarkdownInline>>();
                foreach (var cell in row.Elements<TableCell>())
                {
                    var cellInlines = new List<MarkdownInline>();
                    foreach (var paragraph in cell.Elements<Paragraph>())
                    {
                        cellInlines.AddRange(this.ReadInlines(paragraph));
                    }

                    cells.Add(Coalesce(cellInlines));
                }

                rows.Add(new TableRow(cells));
            }

            return new TableBlock(rows);
        }

        private void AppendRun(List<MarkdownInline> inlines, Run run)
        {
            var text = ReadRunText(run);
            if (text.Length == 0)
            {
                return;
            }

            var rPr = run.RunProperties;
            var charStyleId = rPr?.RunStyle?.Val?.Value;
            if (!string.IsNullOrEmpty(charStyleId))
            {
                var local = ResolveStyleName(charStyleId);
                if (styleMap.IsDefaultFor(MarkdownConstruct.CodeSpan, local))
                {
                    inlines.Add(new CodeSpanInline(text));
                }
                else
                {
                    inlines.Add(new StyledSpanInline(text, local!));
                }

                return;
            }

            if (IsBold(rPr))
            {
                inlines.Add(new EmphasisInline(EmphasisKind.Bold, text));
            }
            else if (IsItalic(rPr))
            {
                inlines.Add(new EmphasisInline(EmphasisKind.Italic, text));
            }
            else if (IsUnderline(rPr))
            {
                inlines.Add(new EmphasisInline(EmphasisKind.Underline, text));
            }
            else
            {
                inlines.Add(new TextInline(text));
            }
        }

        private void AppendHyperlink(List<MarkdownInline> inlines, Hyperlink hyperlink)
        {
            var text = new StringBuilder();
            foreach (var run in hyperlink.Elements<Run>())
            {
                text.Append(ReadRunText(run));
            }

            if (text.Length == 0)
            {
                return;
            }

            var url = this.ResolveHyperlinkUrl(hyperlink.Id?.Value);
            if (string.IsNullOrEmpty(url))
            {
                inlines.Add(new TextInline(text.ToString()));
            }
            else
            {
                inlines.Add(new LinkInline(text.ToString(), url));
            }
        }

        private string? ResolveHyperlinkUrl(string? relationshipId)
        {
            if (string.IsNullOrEmpty(relationshipId))
            {
                return null;
            }

            foreach (var relationship in mainPart.HyperlinkRelationships)
            {
                if (string.Equals(relationship.Id, relationshipId, StringComparison.Ordinal))
                {
                    return relationship.Uri?.ToString();
                }
            }

            return null;
        }

        private string ResolveStyleName(string? styleId)
        {
            if (string.IsNullOrEmpty(styleId))
            {
                return styleMap.DefaultLocalName(MarkdownConstruct.Paragraph);
            }

            return styleNames.TryGetValue(styleId, out var name) ? name : styleId;
        }

        private bool ResolveOrdered(ParagraphProperties? pPr)
        {
            var numId = pPr?.NumberingProperties?.NumberingId?.Val;
            if (numId is null)
            {
                return false;
            }

            var levelRef = pPr?.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
            var numbering = mainPart.NumberingDefinitionsPart?.Numbering;
            if (numbering is null)
            {
                return false;
            }

            var instance = numbering.Elements<NumberingInstance>()
                .FirstOrDefault(n => n.NumberID?.Value == numId.Value);
            var abstractId = instance?.AbstractNumId?.Val?.Value;
            if (abstractId is null)
            {
                return false;
            }

            var abstractNum = numbering.Elements<AbstractNum>()
                .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractId.Value);
            var level = abstractNum?.Elements<Level>()
                .FirstOrDefault(l => (l.LevelIndex?.Value ?? 0) == levelRef);
            var format = level?.NumberingFormat?.Val?.Value;

            // Anything that is not a bullet renders as an ordered list.
            return format is not null && format != NumberFormatValues.Bullet;
        }
    }

    private static string ReadPlainText(Paragraph paragraph)
    {
        var builder = new StringBuilder();
        foreach (var run in paragraph.Elements<Run>())
        {
            builder.Append(ReadRunText(run));
        }

        return builder.ToString();
    }

    private static string ReadRunText(Run run)
    {
        var builder = new StringBuilder();
        foreach (var element in run.Elements())
        {
            switch (element)
            {
                case Text text:
                    builder.Append(text.Text);
                    break;
                case TabChar:
                    builder.Append('\t');
                    break;
                case NoBreakHyphen:
                    builder.Append('-');
                    break;
                case Break:
                case CarriageReturn:
                    builder.Append(' ');
                    break;
            }
        }

        return builder.ToString();
    }

    private static bool IsBold(RunProperties? rPr)
    {
        var bold = rPr?.GetFirstChild<Bold>();
        return bold is not null && (bold.Val is null || bold.Val.Value);
    }

    private static bool IsItalic(RunProperties? rPr)
    {
        var italic = rPr?.GetFirstChild<Italic>();
        return italic is not null && (italic.Val is null || italic.Val.Value);
    }

    private static bool IsUnderline(RunProperties? rPr)
    {
        var underline = rPr?.GetFirstChild<Underline>();
        return underline?.Val is not null && underline.Val.Value != UnderlineValues.None;
    }

    private static bool NameEquals(string? a, string? b) => string.Equals(a, b, StringComparison.Ordinal);

    private static bool IsEmpty(IReadOnlyList<MarkdownInline> inlines) =>
        inlines.Count == 0 || inlines.All(i => i is TextInline { Text: var t } && string.IsNullOrWhiteSpace(t));

    /// <summary>Merge adjacent inlines with the same kind + style (Word splits formatted text across runs).</summary>
    private static IReadOnlyList<MarkdownInline> Coalesce(List<MarkdownInline> inlines)
    {
        var result = new List<MarkdownInline>(inlines.Count);
        foreach (var inline in inlines)
        {
            if (result.Count > 0 && TryMerge(result[^1], inline, out var merged))
            {
                result[^1] = merged;
            }
            else
            {
                result.Add(inline);
            }
        }

        return result;
    }

    private static bool TryMerge(MarkdownInline left, MarkdownInline right, out MarkdownInline merged)
    {
        merged = left;
        switch (left, right)
        {
            case (TextInline a, TextInline b):
                merged = new TextInline(a.Text + b.Text);
                return true;
            case (EmphasisInline a, EmphasisInline b) when a.Kind == b.Kind:
                merged = new EmphasisInline(a.Kind, a.Text + b.Text);
                return true;
            case (CodeSpanInline a, CodeSpanInline b) when a.Style == b.Style:
                merged = new CodeSpanInline(a.Code + b.Code, a.Style);
                return true;
            case (StyledSpanInline a, StyledSpanInline b) when a.Style == b.Style:
                merged = new StyledSpanInline(a.Text + b.Text, a.Style);
                return true;
            default:
                return false;
        }
    }
}
