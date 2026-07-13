using WordComMcp.Markdown;
using WordComMcp.Tests.Fixtures;
using Xunit;

namespace WordComMcp.Tests;

/// <summary>
/// Word-free tests for the OOXML → adjusted-Markdown reader (issue 1.3). Fixtures are built
/// in memory with <see cref="FixtureBuilder"/> and read straight from the package, so no live
/// Word instance is needed.
/// </summary>
public class OoxmlMarkdownReaderTests
{
    private readonly StyleMap m_styleMap = StyleMap.Default;

    [Fact]
    public void ReadFromPackage_GermanFixture_ProducesAdjustedMarkdown()
    {
        // Arrange — the German fixture: Überschrift 1 heading, Standard body, Zitat quote.
        var bytes = FixtureBuilder.BuildGermanFixture();

        // Act
        var model = OoxmlMarkdownReader.ReadFromPackage(bytes, this.m_styleMap);
        var markdown = new MarkdownSerializer(this.m_styleMap).Serialize(model);

        // Assert — heading maps to '#', Standard to a plain paragraph, Zitat (Quote default) to '>'.
        const string expected =
            "# Projektüberschrift\n\n" +
            "Dies ist ein deutscher Beispieltext für die Smoke-Tests.\n\n" +
            "> Ein kurzes Zitat in der Formatvorlage Zitat.";
        Assert.Equal(expected, markdown);
    }

    [Fact]
    public void ReadFromPackage_FinalView_ExcludesDeletedIncludesInserted()
    {
        // Arrange — one paragraph with an inserted "beautiful " and a tracked-deleted "cruel ".
        var bytes = FixtureBuilder.BuildRevisionSample();

        // Act
        var model = OoxmlMarkdownReader.ReadFromPackage(bytes, this.m_styleMap, RevisionView.Final);
        var markdown = new MarkdownSerializer(this.m_styleMap).Serialize(model);

        // Assert — deleted text excluded, inserted text included.
        Assert.Equal("Hello beautiful world", markdown);
        Assert.DoesNotContain("cruel", markdown);
    }

    [Fact]
    public void ReadFromPackage_RichFixture_RoundTripsAllConstructs()
    {
        // Arrange — heading, bold/italic, code span, custom char + paragraph styles, bullet list.
        var bytes = FixtureBuilder.BuildRichFixture();

        // Act
        var model = OoxmlMarkdownReader.ReadFromPackage(bytes, this.m_styleMap);
        var markdown = new MarkdownSerializer(this.m_styleMap).Serialize(model);

        // Assert — every construct maps to its canonical adjusted-Markdown form.
        const string expected =
            "## Kapitel Eins\n\n" +
            "Ein **wichtiger** und *kursiver* Text.\n\n" +
            "Siehe `config.json` Datei.\n\n" +
            "Ein [hervorgehobenes]{style=\"Hervorhebung\"} Wort.\n\n" +
            "::: {style=\"Merksatz\"}\nWichtiger Merksatz.\n:::\n\n" +
            "- Erster Punkt\n- Zweiter Punkt";
        Assert.Equal(expected, markdown);
    }

    [Fact]
    public void ReadFromPackage_RichFixture_IsStableThroughParser()
    {
        // Arrange — get_markdown output must be a fixed point through the parser (round-trip).
        var bytes = FixtureBuilder.BuildRichFixture();
        var markdown = new MarkdownSerializer(this.m_styleMap)
            .Serialize(OoxmlMarkdownReader.ReadFromPackage(bytes, this.m_styleMap));

        // Act — parse the produced markdown and serialize again.
        var reserialized = new MarkdownSerializer(this.m_styleMap).Serialize(MarkdownParser.Parse(markdown));

        // Assert
        Assert.Equal(markdown, reserialized);
    }

    [Fact]
    public void ReadFromPackage_OriginalView_ExcludesInsertedRestoresDeleted()
    {
        // Arrange
        var bytes = FixtureBuilder.BuildRevisionSample();

        // Act — original view is the pre-edit text.
        var model = OoxmlMarkdownReader.ReadFromPackage(bytes, this.m_styleMap, RevisionView.Original);
        var markdown = new MarkdownSerializer(this.m_styleMap).Serialize(model);

        // Assert — deleted "cruel " restored, inserted "beautiful " dropped.
        Assert.Equal("Hello cruel world", markdown);
        Assert.DoesNotContain("beautiful", markdown);
    }
}
