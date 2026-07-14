using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Word = NetOffice.WordApi;

namespace WordComMcp.Markdown;

/// <summary>
/// The OOXML "final text" primitive (issue 0.13): extract the effective text of a range
/// with all tracked revisions applied, <b>without mutating the live document</b>.
///
/// <para>
/// <c>Range.Text</c> cannot be used — it still contains tracked-deleted text. Instead we
/// take the range's <c>WordOpenXML</c> (a Flat-OPC snapshot), load it into
/// <c>DocumentFormat.OpenXml</c> in memory, accept every revision on that throwaway copy,
/// and read the resulting text. The live document is never touched and no revision is
/// created. Shared by <c>get_markdown</c>, verification, and diff-based edits.
/// </para>
/// </summary>
public static class OoxmlFinalText
{
    /// <summary>Get the final (revisions-applied) text of a live Word range.</summary>
    public static string GetFinalText(Word.Range range)
    {
        ArgumentNullException.ThrowIfNull(range);
        return GetFinalTextFromFlatOpc(range.WordOpenXML);
    }

    /// <summary>
    /// Get the final text from a Flat-OPC XML string (the shape returned by
    /// <c>Range.WordOpenXML</c>). Exposed separately so it can be unit-tested without Word.
    /// </summary>
    public static string GetFinalTextFromFlatOpc(string wordOpenXml)
    {
        ArgumentException.ThrowIfNullOrEmpty(wordOpenXml);

        using var doc = WordprocessingDocument.FromFlatOpcString(wordOpenXml);
        return ReadFinalText(doc);
    }

    /// <summary>Get the final text from a full <c>.docx</c> package (byte array). Also test-friendly.</summary>
    public static string GetFinalTextFromPackage(byte[] docxBytes)
    {
        ArgumentNullException.ThrowIfNull(docxBytes);

        using var stream = new MemoryStream();
        stream.Write(docxBytes, 0, docxBytes.Length);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, isEditable: true);
        return ReadFinalText(doc);
    }

    private static string ReadFinalText(WordprocessingDocument doc)
    {
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body is null)
        {
            return string.Empty;
        }

        OoxmlRevisions.Apply(body, RevisionView.Final);
        return ReadParagraphText(body);
    }

    private static string ReadParagraphText(Body body)
    {
        var builder = new StringBuilder();
        var first = true;
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            if (!first)
            {
                builder.Append('\n');
            }

            first = false;
            AppendParagraphContent(paragraph, builder);
        }

        return builder.ToString();
    }

    /// <summary>
    /// Append a paragraph's run content in document order, mapping OOXML text-equivalent
    /// elements to their characters so tabs and soft breaks survive (not just <c>w:t</c>).
    /// </summary>
    private static void AppendParagraphContent(Paragraph paragraph, StringBuilder builder)
    {
        foreach (var element in paragraph.Descendants())
        {
            switch (element)
            {
                case Text text:
                    builder.Append(text.Text);
                    break;
                case TabChar:
                    builder.Append('\t');
                    break;
                case Break:
                case CarriageReturn:
                    builder.Append('\n');
                    break;
                case NoBreakHyphen:
                    builder.Append('-');
                    break;
            }
        }
    }
}
