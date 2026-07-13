using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordComMcp.Tests.Fixtures;

/// <summary>
/// Builds the reference <c>.docx</c> fixtures used by the tests (issue 0.11) with
/// <see cref="DocumentFormat.OpenXml"/> — no live Word needed to create them.
/// </summary>
public static class FixtureBuilder
{
    /// <summary>
    /// A German, tracking-capable document with the built-in <c>Standard</c>/<c>Überschrift 1</c>
    /// styles plus custom <c>Zitat</c> (paragraph) and <c>Code</c> (character) Formatvorlagen.
    /// </summary>
    public static byte[] BuildGermanFixture()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            AddStyles(mainPart);

            var body = new Body(
                Heading("Überschrift 1", "Projektüberschrift"),
                Normal("Standard", "Dies ist ein deutscher Beispieltext für die Smoke-Tests."),
                Normal("Zitat", "Ein kurzes Zitat in der Formatvorlage Zitat."),
                new SectionProperties());
            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    /// <summary>
    /// A one-paragraph document with overlapping tracked insert + delete revisions.
    /// Final (accepted) text is <c>"Hello beautiful world"</c>; raw text still contains the
    /// deleted <c>"cruel "</c>. Used to prove the 0.13 final-text primitive.
    /// </summary>
    public static byte[] BuildRevisionSample()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();

            var paragraph = new Paragraph(
                PlainRun("Hello "),
                new InsertedRun(PlainRun("beautiful "))
                {
                    Author = "Tester",
                    Id = "1",
                },
                new DeletedRun(
                    new Run(new DeletedText("cruel ") { Space = SpaceProcessingModeValues.Preserve }))
                {
                    Author = "Tester",
                    Id = "2",
                },
                PlainRun("world"));

            mainPart.Document = new Document(new Body(paragraph, new SectionProperties()));
            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    /// <summary>The Flat-OPC string of <see cref="BuildRevisionSample"/> (the shape <c>Range.WordOpenXML</c> returns).</summary>
    public static string BuildRevisionSampleFlatOpc()
    {
        var bytes = BuildRevisionSample();
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        return doc.ToFlatOpcString();
    }

    /// <summary>
    /// A paragraph exercising OOXML text-equivalent elements: <c>"A"</c>, <c>&lt;w:tab/&gt;</c>,
    /// <c>"B"</c>, <c>&lt;w:br/&gt;</c>, <c>"C"</c>. Final text must be <c>"A\tB\nC"</c>.
    /// </summary>
    public static byte[] BuildTabBreakSample()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var paragraph = new Paragraph(
                PlainRun("A"),
                new Run(new TabChar()),
                PlainRun("B"),
                new Run(new Break()),
                PlainRun("C"));

            mainPart.Document = new Document(new Body(paragraph, new SectionProperties()));
            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    /// <summary>
    /// A richer German fixture exercising every read path: a Heading 2, direct bold/italic runs,
    /// a <c>Code</c> character-styled span, a custom <c>Hervorhebung</c> character span, a custom
    /// <c>Merksatz</c> paragraph style, and a two-item bullet list (styled <c>Listenabsatz</c>).
    /// </summary>
    public static byte[] BuildRichFixture()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            AddRichStyles(mainPart);

            var body = new Body(
                Heading("Überschrift 2", "Kapitel Eins"),
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "Standard" }),
                    PlainRun("Ein "),
                    FormattedRun("wichtiger", new Bold()),
                    PlainRun(" und "),
                    FormattedRun("kursiver", new Italic()),
                    PlainRun(" Text.")),
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "Standard" }),
                    PlainRun("Siehe "),
                    FormattedRun("config.json", new RunStyle { Val = "Code" }),
                    PlainRun(" Datei.")),
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "Standard" }),
                    PlainRun("Ein "),
                    FormattedRun("hervorgehobenes", new RunStyle { Val = "Hervorhebung" }),
                    PlainRun(" Wort.")),
                Normal("Merksatz", "Wichtiger Merksatz."),
                Normal("Listenabsatz", "Erster Punkt"),
                Normal("Listenabsatz", "Zweiter Punkt"),
                new SectionProperties());
            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    /// <summary>Write the German fixture to <paramref name="path"/> for the live Word smoke test.</summary>
    public static void WriteGermanFixture(string path) => File.WriteAllBytes(path, BuildGermanFixture());

    private static Run FormattedRun(string text, params OpenXmlElement[] properties) =>
        new(new RunProperties(properties), new Text(text) { Space = SpaceProcessingModeValues.Preserve });

    private static void AddRichStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(
            ParagraphStyle("Standard", "Standard", isDefault: true),
            ParagraphStyle("Überschrift 2", "Überschrift 2"),
            ParagraphStyle("Listenabsatz", "Listenabsatz"),
            ParagraphStyle("Merksatz", "Merksatz"),
            CharacterStyle("Code", "Code"),
            CharacterStyle("Hervorhebung", "Hervorhebung"));
        stylesPart.Styles.Save();
    }

    private static Run PlainRun(string text) =>
        new(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

    private static Paragraph Heading(string styleName, string text) =>
        new(new ParagraphProperties(new ParagraphStyleId { Val = styleName }), PlainRun(text));

    private static Paragraph Normal(string styleName, string text) =>
        new(new ParagraphProperties(new ParagraphStyleId { Val = styleName }), PlainRun(text));

    private static void AddStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(
            ParagraphStyle("Standard", "Standard", isDefault: true),
            ParagraphStyle("Überschrift 1", "Überschrift 1"),
            ParagraphStyle("Zitat", "Zitat"),
            CharacterStyle("Code", "Code"));
        stylesPart.Styles.Save();
    }

    private static Style ParagraphStyle(string styleId, string nameLocal, bool isDefault = false)
    {
        var style = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId,
            Default = isDefault ? OnOffValue.FromBoolean(true) : null,
        };
        style.Append(new StyleName { Val = nameLocal });
        return style;
    }

    private static Style CharacterStyle(string styleId, string nameLocal)
    {
        var style = new Style
        {
            Type = StyleValues.Character,
            StyleId = styleId,
        };
        style.Append(new StyleName { Val = nameLocal });
        return style;
    }
}
