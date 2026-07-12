using WordComMcp.Markdown;
using Xunit;

namespace WordComMcp.Tests;

/// <summary>Round-trip and control-char tests for the adjusted-Markdown layer (issue 0.14).</summary>
public class MarkdownRoundTripTests
{
    private readonly MarkdownSerializer m_serializer = new(StyleMap.Default);

    public static TheoryData<string> CanonicalDocuments() => new()
    {
        "# Überschrift eins",
        "## Zwei\n\n### Drei",
        "Ein einfacher Absatz mit Text.",
        "Absatz mit **fett**, *kursiv* und _unterstrichen_.",
        "Ein Absatz mit `Code` und einem [Link](https://example.com).",
        "Ein Absatz mit `k`{style=\"Quellcode\"} custom code span.",
        "Ein Absatz mit [Begriff]{style=\"Zitat\"} styled span.",
        "- Eins\n- Zwei\n- Drei",
        "1. Erstens\n2. Zweitens",
        "> Ein zitierter Satz.",
        "```csharp\nvar x = 1;\n```",
        "::: {style=\"Zitat\"}\nEin Absatz in der Formatvorlage Zitat.\n:::",
        "| A | B |\n| --- | --- |\n| 1 | 2 |",
        "# Titel\n\nEin Absatz.\n\n- Punkt eins\n- Punkt zwei\n\n> Zitat",
    };

    [Theory]
    [MemberData(nameof(CanonicalDocuments))]
    public void SerializeParse_IsStableForCanonicalMarkdown(string canonical)
    {
        // Arrange / Act — serialize(parse(x)) must equal the canonical form exactly.
        var once = this.m_serializer.Serialize(MarkdownParser.Parse(canonical));

        // Assert
        Assert.Equal(canonical, once);
    }

    [Theory]
    [MemberData(nameof(CanonicalDocuments))]
    public void ParseSerialize_RoundTripsLosslessly(string canonical)
    {
        // Arrange
        var first = this.m_serializer.Serialize(MarkdownParser.Parse(canonical));

        // Act — a second pass must be a fixed point.
        var second = this.m_serializer.Serialize(MarkdownParser.Parse(first));

        // Assert
        Assert.Equal(first, second);
    }

    [Fact]
    public void Parse_StripsControlCharacters()
    {
        // Arrange — embed a bell (0x07) cell mark that must never survive.
        var input = "Text mit \x07 Steuerzeichen.";

        // Act
        var output = this.m_serializer.Serialize(MarkdownParser.Parse(input));

        // Assert
        Assert.DoesNotContain('\x07', output);
        Assert.Equal("Text mit  Steuerzeichen.", output);
    }

    [Fact]
    public void Parse_KeepsCustomParagraphStyleAnnotation()
    {
        // Arrange
        var input = "::: {style=\"Zitat\"}\nHallo Welt\n:::";

        // Act
        var doc = MarkdownParser.Parse(input);

        // Assert
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
        Assert.Equal("Zitat", paragraph.Style);
    }

    [Fact]
    public void Serialize_OmitsAnnotationForDefaultStyle()
    {
        // Arrange — a paragraph explicitly tagged with the default "Standard" style.
        var doc = new MarkdownDocument([new ParagraphBlock([new TextInline("Hallo")], "Standard")]);

        // Act
        var markdown = this.m_serializer.Serialize(doc);

        // Assert — no redundant {style=…} for the construct default.
        Assert.Equal("Hallo", markdown);
    }
}
