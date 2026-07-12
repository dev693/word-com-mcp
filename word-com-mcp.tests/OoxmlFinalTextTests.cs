using System.Security.Cryptography;
using WordComMcp.Markdown;
using WordComMcp.Tests.Fixtures;
using Xunit;

namespace WordComMcp.Tests;

/// <summary>Tests for the OOXML final-text primitive (issue 0.13).</summary>
public class OoxmlFinalTextTests
{
    [Fact]
    public void GetFinalTextFromPackage_AppliesRevisions()
    {
        // Arrange — a paragraph with an inserted "beautiful " and a deleted "cruel ".
        var bytes = FixtureBuilder.BuildRevisionSample();

        // Act
        var finalText = OoxmlFinalText.GetFinalTextFromPackage(bytes);

        // Assert — inserted text kept, deleted text excluded.
        Assert.Equal("Hello beautiful world", finalText);
        Assert.DoesNotContain("cruel", finalText);
    }

    [Fact]
    public void GetFinalTextFromFlatOpc_AppliesRevisions()
    {
        // Arrange — the Flat-OPC shape returned by Range.WordOpenXML.
        var flatOpc = FixtureBuilder.BuildRevisionSampleFlatOpc();

        // Act
        var finalText = OoxmlFinalText.GetFinalTextFromFlatOpc(flatOpc);

        // Assert
        Assert.Equal("Hello beautiful world", finalText);
    }

    [Fact]
    public void GetFinalTextFromPackage_PreservesTabsAndLineBreaks()
    {
        // Arrange — a paragraph "A" <tab> "B" <break> "C" (text-equivalent OOXML elements).
        var bytes = FixtureBuilder.BuildTabBreakSample();

        // Act
        var finalText = OoxmlFinalText.GetFinalTextFromPackage(bytes);

        // Assert — w:tab and w:br must survive, not collapse to "ABC".
        Assert.Equal("A\tB\nC", finalText);
    }

    [Fact]
    public void GetFinalTextFromPackage_DoesNotMutateTheSourceBytes()
    {
        // Arrange
        var bytes = FixtureBuilder.BuildRevisionSample();
        var before = SHA256.HashData(bytes);

        // Act
        _ = OoxmlFinalText.GetFinalTextFromPackage(bytes);

        // Assert — the primitive works on a throwaway copy; the input is untouched.
        Assert.Equal(before, SHA256.HashData(bytes));
    }
}
