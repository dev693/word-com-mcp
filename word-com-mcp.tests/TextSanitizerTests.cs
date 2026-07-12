using WordComMcp.Infrastructure;
using Xunit;

namespace WordComMcp.Tests;

/// <summary>Control-char stripping and newline normalization (Conventions "Control-char safety").</summary>
public class TextSanitizerTests
{
    [Theory]
    [InlineData('\x00')] // C0 NUL
    [InlineData('\x07')] // C0 BEL (cell mark)
    [InlineData('\x1f')] // C0 unit separator
    [InlineData('\x7f')] // DEL
    [InlineData('\x80')] // C1 lower bound
    [InlineData('\x85')] // C1 NEL
    [InlineData('\x9f')] // C1 upper bound
    public void StripControlChars_RemovesC0DelAndC1Controls(char control)
    {
        // Arrange
        var input = $"A{control}B";

        // Act
        var output = TextSanitizer.StripControlChars(input);

        // Assert
        Assert.Equal("AB", output);
    }

    [Fact]
    public void StripControlChars_PreservesTabCrLf()
    {
        Assert.Equal("A\tB\r\nC", TextSanitizer.StripControlChars("A\tB\r\nC"));
    }

    [Fact]
    public void StripControlChars_KeepsPrintableUnicodeAboveC1()
    {
        // U+00A0 (NBSP) and beyond are not control chars and must pass through.
        Assert.Equal("A ÖB", TextSanitizer.StripControlChars("A ÖB"));
    }

    [Fact]
    public void NormalizeNewlines_ConvertsToWordParagraphMark()
    {
        Assert.Equal("a\rb\rc", TextSanitizer.NormalizeNewlines("a\r\nb\nc"));
    }
}
