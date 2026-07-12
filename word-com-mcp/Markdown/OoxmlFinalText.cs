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

        AcceptAllRevisions(body);
        return ReadParagraphText(body);
    }

    /// <summary>
    /// Accept every tracked revision in <paramref name="root"/> in memory: drop deletions
    /// and move-froms, unwrap insertions and move-tos, and remove formatting-change markers.
    /// </summary>
    private static void AcceptAllRevisions(OpenXmlElement root)
    {
        // Deletions and moved-away content are removed outright (the text stays gone).
        RemoveAll(root.Descendants<Deleted>());
        RemoveAll(root.Descendants<DeletedRun>());
        RemoveAll(root.Descendants<DeletedMathControl>());
        RemoveAll(root.Descendants<MoveFrom>());
        RemoveAll(root.Descendants<MoveFromRun>());

        // Insertions and moved-in content are promoted out of their revision wrapper.
        Unwrap(root.Descendants<InsertedRun>());
        Unwrap(root.Descendants<Inserted>());
        Unwrap(root.Descendants<MoveToRun>());
        Unwrap(root.Descendants<MoveTo>());

        // Formatting/property changes: remove the change marker, keep the applied formatting.
        RemoveAll(root.Descendants<ParagraphPropertiesChange>());
        RemoveAll(root.Descendants<RunPropertiesChange>());
        RemoveAll(root.Descendants<SectionPropertiesChange>());
        RemoveAll(root.Descendants<TablePropertiesChange>());
        RemoveAll(root.Descendants<TableRowPropertiesChange>());
        RemoveAll(root.Descendants<TableCellPropertiesChange>());
        RemoveAll(root.Descendants<TablePropertyExceptionsChange>());
        RemoveAll(root.Descendants<NumberingChange>());
    }

    private static void RemoveAll(IEnumerable<OpenXmlElement> elements)
    {
        foreach (var element in elements.ToList())
        {
            element.Remove();
        }
    }

    private static void Unwrap(IEnumerable<OpenXmlElement> wrappers)
    {
        foreach (var wrapper in wrappers.ToList())
        {
            // Move each child up in front of the wrapper, then drop the (now empty) wrapper.
            foreach (var child in wrapper.ChildElements.ToList())
            {
                wrapper.InsertBeforeSelf(child.CloneNode(true));
            }

            wrapper.Remove();
        }
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
            foreach (var text in paragraph.Descendants<Text>())
            {
                builder.Append(text.Text);
            }
        }

        return builder.ToString();
    }
}
