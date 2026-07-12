using WordComMcp.Com;
using WordComMcp.Markdown;
using WdBuiltinStyle = NetOffice.WordApi.Enums.WdBuiltinStyle;
using Xunit;

namespace WordComMcp.Tests;

/// <summary>Style-map resolution and validation tests (issue 0.12).</summary>
public class StyleMapTests
{
    private readonly StyleMap m_map = StyleMap.Default;

    [Theory]
    [InlineData(MarkdownConstruct.Heading1, WdBuiltinStyle.wdStyleHeading1)]
    [InlineData(MarkdownConstruct.Paragraph, WdBuiltinStyle.wdStyleNormal)]
    [InlineData(MarkdownConstruct.Quote, WdBuiltinStyle.wdStyleQuote)]
    public void BuiltInStyle_ResolvesLanguageIndependentEnumId(MarkdownConstruct construct, WdBuiltinStyle expected)
    {
        Assert.Equal(expected, this.m_map.BuiltInStyle(construct));
    }

    [Theory]
    [InlineData(MarkdownConstruct.CodeBlock)]
    [InlineData(MarkdownConstruct.CodeSpan)]
    public void BuiltInStyle_IsNullForCustomFormatvorlagen(MarkdownConstruct construct)
    {
        Assert.Null(this.m_map.BuiltInStyle(construct));
    }

    [Fact]
    public void DefaultLocalName_UsesGermanBuiltInsAndCustomStyles()
    {
        Assert.Equal("Standard", this.m_map.DefaultLocalName(MarkdownConstruct.Paragraph));
        Assert.Equal("Zitat", this.m_map.DefaultLocalName(MarkdownConstruct.Quote));
        Assert.Equal("Code", this.m_map.DefaultLocalName(MarkdownConstruct.CodeSpan));
    }

    [Fact]
    public void IsDefaultFor_TreatsNullAndMatchingNameAsDefault()
    {
        Assert.True(this.m_map.IsDefaultFor(MarkdownConstruct.Paragraph, null));
        Assert.True(this.m_map.IsDefaultFor(MarkdownConstruct.Paragraph, "Standard"));
        Assert.False(this.m_map.IsDefaultFor(MarkdownConstruct.Paragraph, "Zitat"));
    }

    [Fact]
    public void ValidateStyle_PassesForAKnownStyle()
    {
        var available = new[] { "Standard", "Zitat", "Code" };
        StyleMap.ValidateStyle("Zitat", available); // does not throw
    }

    [Fact]
    public void ValidateStyle_ThrowsStructuredErrorListingAvailableStyles()
    {
        // Arrange
        var available = new[] { "Standard", "Zitat", "Code" };

        // Act
        var ex = Assert.Throws<WordConnectionException>(() => StyleMap.ValidateStyle("Nonexistent", available));

        // Assert — structured error code + the available styles are listed for the caller.
        Assert.Equal("unknown style", ex.ErrorCode);
        Assert.Contains("Standard", ex.Message);
        Assert.Contains("Zitat", ex.Message);
        Assert.Contains("Code", ex.Message);
    }
}
