namespace WordComMcp.Markdown;

/// <summary>
/// AST for the adjusted-Markdown subset (issue 0.14 / Conventions "interface"). A style
/// (<c>null</c> = the construct's default) captures Pandoc <c>{style=…}</c> annotations so
/// custom Formatvorlagen survive a read → edit → write round-trip.
/// </summary>
public sealed record MarkdownDocument(IReadOnlyList<MarkdownBlock> Blocks);

/// <summary>Base type for block-level nodes.</summary>
public abstract record MarkdownBlock;

/// <summary>A heading (<c>#</c>/<c>##</c>/<c>###</c>), levels 1–3.</summary>
public sealed record HeadingBlock(int Level, IReadOnlyList<MarkdownInline> Inlines, string? Style = null) : MarkdownBlock;

/// <summary>A plain paragraph. A non-null <see cref="Style"/> came from a <c>::: {style=…}</c> div.</summary>
public sealed record ParagraphBlock(IReadOnlyList<MarkdownInline> Inlines, string? Style = null) : MarkdownBlock;

/// <summary>A blockquote (<c>&gt;</c>), default style Quote/Zitat.</summary>
public sealed record BlockQuoteBlock(IReadOnlyList<MarkdownInline> Inlines, string? Style = null) : MarkdownBlock;

/// <summary>A fenced code block (<c>```</c>), default style Code.</summary>
public sealed record CodeBlock(string Code, string? InfoString = null, string? Style = null) : MarkdownBlock;

/// <summary>A bullet (<c>-</c>) or numbered (<c>1.</c>) list.</summary>
public sealed record ListBlock(bool Ordered, IReadOnlyList<ListItem> Items) : MarkdownBlock;

/// <summary>One list item's inline content.</summary>
public sealed record ListItem(IReadOnlyList<MarkdownInline> Inlines, string? Style = null);

/// <summary>A simple table: rows of cells, each cell a list of inlines. The first row is the header.</summary>
public sealed record TableBlock(IReadOnlyList<TableRow> Rows) : MarkdownBlock;

/// <summary>One table row.</summary>
public sealed record TableRow(IReadOnlyList<IReadOnlyList<MarkdownInline>> Cells);

/// <summary>Base type for inline (span-level) nodes.</summary>
public abstract record MarkdownInline;

/// <summary>Literal text.</summary>
public sealed record TextInline(string Text) : MarkdownInline;

/// <summary>Emphasis kinds mapped to Word direct formatting.</summary>
public enum EmphasisKind
{
    Bold,
    Italic,
    Underline,
}

/// <summary><c>**bold**</c> / <c>*italic*</c> / <c>_underline_</c>.</summary>
public sealed record EmphasisInline(EmphasisKind Kind, string Text) : MarkdownInline;

/// <summary>An inline code span (<c>`x`</c>), default character style Code. May carry an explicit <c>{style=…}</c>.</summary>
public sealed record CodeSpanInline(string Code, string? Style = null) : MarkdownInline;

/// <summary>A link <c>[text](url)</c>.</summary>
public sealed record LinkInline(string Text, string Url) : MarkdownInline;

/// <summary>A Pandoc styled span <c>[text]{style="X"}</c> assigning a character Formatvorlage.</summary>
public sealed record StyledSpanInline(string Text, string Style) : MarkdownInline;
