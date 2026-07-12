using System.Text;

namespace WordComMcp.Markdown;

/// <summary>
/// Serializer for the adjusted-Markdown subset (issue 0.14). Emits canonical markdown that
/// round-trips through <see cref="MarkdownParser"/>. A <c>{style=…}</c> annotation is emitted
/// for any paragraph/span whose style differs from the construct's default in the
/// <see cref="StyleMap"/>, so custom Formatvorlagen survive read → edit → write.
/// </summary>
public sealed class MarkdownSerializer
{
    private readonly StyleMap m_styleMap;

    public MarkdownSerializer(StyleMap styleMap)
    {
        this.m_styleMap = styleMap ?? throw new ArgumentNullException(nameof(styleMap));
    }

    public string Serialize(MarkdownDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var parts = new List<string>(document.Blocks.Count);
        foreach (var block in document.Blocks)
        {
            parts.Add(this.SerializeBlock(block));
        }

        return string.Join("\n\n", parts);
    }

    private string SerializeBlock(MarkdownBlock block) => block switch
    {
        HeadingBlock heading => $"{new string('#', heading.Level)} {this.SerializeInlines(heading.Inlines)}",
        ParagraphBlock paragraph => this.SerializeParagraph(paragraph),
        BlockQuoteBlock quote => this.SerializeBlockQuote(quote),
        CodeBlock code => SerializeCode(code),
        ListBlock list => this.SerializeList(list),
        TableBlock table => this.SerializeTable(table),
        _ => throw new ArgumentOutOfRangeException(nameof(block), block.GetType().Name, "Unknown block type"),
    };

    private string SerializeParagraph(ParagraphBlock paragraph)
    {
        var text = this.SerializeInlines(paragraph.Inlines);
        if (this.m_styleMap.IsDefaultFor(MarkdownConstruct.Paragraph, paragraph.Style))
        {
            return text;
        }

        return $"::: {{style=\"{paragraph.Style}\"}}\n{text}\n:::";
    }

    private string SerializeBlockQuote(BlockQuoteBlock quote)
    {
        var text = this.SerializeInlines(quote.Inlines);
        var lines = text.Split('\n');
        var builder = new StringBuilder();
        for (var i = 0; i < lines.Length; i++)
        {
            if (i > 0)
            {
                builder.Append('\n');
            }

            builder.Append("> ").Append(lines[i]);
        }

        return builder.ToString();
    }

    private static string SerializeCode(CodeBlock code)
    {
        var info = string.IsNullOrEmpty(code.InfoString) ? string.Empty : code.InfoString;
        return $"```{info}\n{code.Code}\n```";
    }

    private string SerializeList(ListBlock list)
    {
        var builder = new StringBuilder();
        for (var i = 0; i < list.Items.Count; i++)
        {
            if (i > 0)
            {
                builder.Append('\n');
            }

            var marker = list.Ordered ? $"{i + 1}. " : "- ";
            builder.Append(marker).Append(this.SerializeInlines(list.Items[i].Inlines));
        }

        return builder.ToString();
    }

    private string SerializeTable(TableBlock table)
    {
        if (table.Rows.Count == 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        var columnCount = table.Rows[0].Cells.Count;

        AppendRow(builder, table.Rows[0]);
        builder.Append('\n').Append("| ").Append(string.Join(" | ", Enumerable.Repeat("---", columnCount))).Append(" |");

        for (var r = 1; r < table.Rows.Count; r++)
        {
            builder.Append('\n');
            AppendRow(builder, table.Rows[r]);
        }

        return builder.ToString();

        void AppendRow(StringBuilder sb, TableRow row)
        {
            sb.Append("| ");
            sb.Append(string.Join(" | ", row.Cells.Select(this.SerializeInlines)));
            sb.Append(" |");
        }
    }

    private string SerializeInlines(IReadOnlyList<MarkdownInline> inlines)
    {
        var builder = new StringBuilder();
        foreach (var inline in inlines)
        {
            builder.Append(this.SerializeInline(inline));
        }

        return builder.ToString();
    }

    private string SerializeInline(MarkdownInline inline) => inline switch
    {
        TextInline text => text.Text,
        EmphasisInline emphasis => SerializeEmphasis(emphasis),
        CodeSpanInline code => this.SerializeCodeSpan(code),
        LinkInline link => $"[{link.Text}]({link.Url})",
        StyledSpanInline span => $"[{span.Text}]{{style=\"{span.Style}\"}}",
        _ => throw new ArgumentOutOfRangeException(nameof(inline), inline.GetType().Name, "Unknown inline type"),
    };

    private static string SerializeEmphasis(EmphasisInline emphasis) => emphasis.Kind switch
    {
        EmphasisKind.Bold => $"**{emphasis.Text}**",
        EmphasisKind.Italic => $"*{emphasis.Text}*",
        EmphasisKind.Underline => $"_{emphasis.Text}_",
        _ => throw new ArgumentOutOfRangeException(nameof(emphasis), emphasis.Kind, "Unknown emphasis kind"),
    };

    private string SerializeCodeSpan(CodeSpanInline code)
    {
        if (this.m_styleMap.IsDefaultFor(MarkdownConstruct.CodeSpan, code.Style))
        {
            return $"`{code.Code}`";
        }

        return $"`{code.Code}`{{style=\"{code.Style}\"}}";
    }
}
