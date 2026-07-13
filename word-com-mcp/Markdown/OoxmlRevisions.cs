using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordComMcp.Markdown;

/// <summary>Which revision view a read produces (issue 1.3 <c>get_markdown view=…</c>).</summary>
public enum RevisionView
{
    /// <summary>Revisions applied: insertions kept, deletions dropped (the default effective text).</summary>
    Final,

    /// <summary>Pre-edit text: insertions dropped, deletions restored.</summary>
    Original,

    /// <summary>All text: both inserted and deleted content shown inline.</summary>
    Markup,
}

/// <summary>
/// In-memory revision resolution over an OOXML tree (issue 0.13 shared walk). Extracted so
/// both <see cref="OoxmlFinalText"/> and <see cref="OoxmlMarkdownReader"/> apply revisions the
/// same way, without ever mutating the live Word document. Property-change markers are always
/// stripped (the applied formatting is kept).
/// </summary>
public static class OoxmlRevisions
{
    /// <summary>Resolve the requested <paramref name="view"/> on <paramref name="root"/> in place.</summary>
    public static void Apply(OpenXmlElement root, RevisionView view)
    {
        ArgumentNullException.ThrowIfNull(root);

        switch (view)
        {
            case RevisionView.Final:
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
                break;

            case RevisionView.Original:
                // Insertions never existed pre-edit; deleted text is restored.
                RemoveAll(root.Descendants<Inserted>());
                RemoveAll(root.Descendants<InsertedRun>());
                RemoveAll(root.Descendants<MoveTo>());
                RemoveAll(root.Descendants<MoveToRun>());
                Unwrap(root.Descendants<DeletedRun>());
                Unwrap(root.Descendants<Deleted>());
                Unwrap(root.Descendants<MoveFromRun>());
                Unwrap(root.Descendants<MoveFrom>());
                ConvertDeletedText(root);
                break;

            case RevisionView.Markup:
                // Show everything: both inserted and deleted content promoted to plain text.
                Unwrap(root.Descendants<InsertedRun>());
                Unwrap(root.Descendants<Inserted>());
                Unwrap(root.Descendants<MoveToRun>());
                Unwrap(root.Descendants<MoveTo>());
                Unwrap(root.Descendants<DeletedRun>());
                Unwrap(root.Descendants<Deleted>());
                Unwrap(root.Descendants<MoveFromRun>());
                Unwrap(root.Descendants<MoveFrom>());
                ConvertDeletedText(root);
                break;
        }

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

    /// <summary>Convert restored <c>w:delText</c> nodes to ordinary <c>w:t</c> so text readers pick them up.</summary>
    private static void ConvertDeletedText(OpenXmlElement root)
    {
        foreach (var deleted in root.Descendants<DeletedText>().ToList())
        {
            var text = new Text(deleted.Text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve };
            deleted.Parent?.ReplaceChild(text, deleted);
        }
    }
}
