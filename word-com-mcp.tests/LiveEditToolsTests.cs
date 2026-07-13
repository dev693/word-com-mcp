using System.Text.Json;
using WordComMcp.Com;
using WordComMcp.Infrastructure;
using WordComMcp.Markdown;
using WordComMcp.Tests.Fixtures;
using WordComMcp.Tools;
using Xunit;
using Word = NetOffice.WordApi;
using WdAlertLevel = NetOffice.WordApi.Enums.WdAlertLevel;

namespace WordComMcp.Tests;

/// <summary>
/// Tier-1 integration tests (issue 1.10) exercising the real COM tools against a live Word on
/// the German fixture with tracking ON. Gated behind <c>[Trait("Category","Live")]</c> so a
/// machine without Word still passes via <c>dotnet test --filter Category!=Live</c>; run these
/// with <c>dotnet test --filter Category=Live</c> on a Windows box with Word installed.
///
/// <para>xUnit builds a fresh instance per test, so the constructor opens its own Word + fixture
/// copy and <see cref="Dispose"/> tears it down — each test is isolated.</para>
/// </summary>
[Trait("Category", "Live")]
public sealed class LiveEditToolsTests : IDisposable
{
    private readonly StaDispatcher m_dispatcher = new();
    private readonly WordConnection m_connection = new();
    private readonly StyleMap m_styleMap = StyleMap.Default;
    private readonly ServerConfig m_config = new(author: "Tester", authorInitials: "T", styleMapPath: null);
    private readonly string m_path;
    private Word.Application? m_word;

    public LiveEditToolsTests()
    {
        this.m_path = Path.Combine(Path.GetTempPath(), $"wl-mcp-edit-{Guid.NewGuid():N}.docx");
        FixtureBuilder.WriteGermanFixture(this.m_path);

        this.m_dispatcher.RunOnStaAsync(() =>
        {
            this.m_word = new Word.Application { Visible = false, DisplayAlerts = WdAlertLevel.wdAlertsNone };
            var doc = this.m_word.Documents.Open(this.m_path);
            doc.TrackRevisions = true;
            return true;
        }).GetAwaiter().GetResult();
    }

    private string Name => Path.GetFileName(this.m_path);

    [Fact]
    public async Task GetMarkdown_FinalView_ReturnsCleanAdjustedMarkdown()
    {
        // Act
        var json = await ReadTools.GetMarkdown(this.m_dispatcher, this.m_connection, this.m_styleMap, filename: this.Name);

        // Assert — heading, Standard body, and the Zitat quote all present.
        using var scope = new AssertionsScope(json);
        Assert.True(scope.Success);
        var markdown = scope.String("markdown");
        Assert.Contains("# Projektüberschrift", markdown);
        Assert.Contains("> Ein kurzes Zitat", markdown);
    }

    [Fact]
    public async Task ListStyles_IncludesGermanFormatvorlagen()
    {
        // Act
        var json = await StyleTools.ListStyles(this.m_dispatcher, this.m_connection, this.Name);

        // Assert
        Assert.True(new AssertionsScope(json).Success);
        Assert.Contains("Zitat", json);
        Assert.Contains("Code", json);
    }

    [Fact]
    public async Task InsertMarkdown_RealizesGermanStyles_AsTrackedInsertion()
    {
        // Arrange
        var before = await this.RevisionCountAsync();
        const string markdown =
            "## Neue Überschrift\n\n- Punkt eins\n- Punkt zwei\n\n::: {style=\"Zitat\"}\nEin eingefügtes Zitat.\n:::";

        // Act
        var json = await EditTools.InsertMarkdown(
            this.m_dispatcher, this.m_connection, this.m_styleMap, this.m_config,
            markdown, afterText: "Beispieltext", filename: this.Name);

        // Assert — inserted as tracked revisions authored by the configured user, visible in the final view.
        Assert.True(new AssertionsScope(json).Success);
        var (count, author) = await this.FirstRevisionAsync();
        Assert.True(count > before);
        Assert.Equal("Tester", author);

        var readBack = await ReadTools.GetMarkdown(this.m_dispatcher, this.m_connection, this.m_styleMap, filename: this.Name);
        var markdownAfter = new AssertionsScope(readBack).String("markdown");
        Assert.Contains("## Neue Überschrift", markdownAfter);
        Assert.Contains("- Punkt eins", markdownAfter);
        Assert.Contains("> Ein eingefügtes Zitat", markdownAfter);
    }

    [Fact]
    public async Task ReplaceText_ProducesMinimalRedline()
    {
        // Act
        var json = await EditTools.ReplaceText(
            this.m_dispatcher, this.m_connection, this.m_styleMap, this.m_config,
            find: "deutscher", replacement: "englischer", filename: this.Name);

        // Assert — final text swapped; the change is tracked (revision count grew).
        using var scope = new AssertionsScope(json);
        Assert.True(scope.Success);
        Assert.Equal(1, scope.Int("replaced"));

        var final = await this.FinalTextAsync();
        Assert.Contains("englischer", final);
        Assert.DoesNotContain("deutscher", final);
        Assert.True(await this.RevisionCountAsync() > 0);
    }

    [Fact]
    public async Task DeleteText_RemovesFromFinalView_AsTrackedDeletion()
    {
        // Act
        var json = await EditTools.DeleteText(
            this.m_dispatcher, this.m_connection, this.m_styleMap, this.m_config,
            find: "kurzes ", filename: this.Name);

        // Assert — gone from the final view but recorded as a tracked deletion.
        using var scope = new AssertionsScope(json);
        Assert.True(scope.Success);
        Assert.Equal(1, scope.Int("deleted"));

        var final = await this.FinalTextAsync();
        Assert.DoesNotContain("kurzes", final);
        Assert.True(await this.RevisionCountAsync() > 0);
    }

    [Fact]
    public async Task FindText_ReturnsAnnotatedHits()
    {
        // Act
        var json = await ReadTools.FindText(this.m_dispatcher, this.m_connection, "Zitat", filename: this.Name);

        // Assert
        using var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean());
        Assert.True(doc.RootElement.GetProperty("count").GetInt32() >= 1);
        foreach (var match in doc.RootElement.GetProperty("matches").EnumerateArray())
        {
            Assert.False(string.IsNullOrEmpty(match.GetProperty("revisionState").GetString()));
        }
    }

    [Fact]
    public async Task ToggleTrackChanges_ReportsAndFlipsState()
    {
        // Act — tracking starts enabled (set in the constructor); turn it off.
        var json = await DocumentTools.ToggleTrackChanges(this.m_dispatcher, this.m_connection, enabled: false, filename: this.Name);

        // Assert
        using var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean());
        Assert.True(doc.RootElement.GetProperty("previous").GetBoolean());
        Assert.False(doc.RootElement.GetProperty("current").GetBoolean());
    }

    private async Task<string> FinalTextAsync() =>
        await this.m_dispatcher.RunOnStaAsync(() =>
        {
            var app = this.m_connection.GetWordApp();
            var doc = this.m_connection.FindDocument(app, this.Name);
            return OoxmlFinalText.GetFinalText(doc.Content);
        });

    private async Task<int> RevisionCountAsync() =>
        await this.m_dispatcher.RunOnStaAsync(() =>
        {
            var app = this.m_connection.GetWordApp();
            var doc = this.m_connection.FindDocument(app, this.Name);
            return doc.Revisions.Count;
        });

    private async Task<(int Count, string? Author)> FirstRevisionAsync() =>
        await this.m_dispatcher.RunOnStaAsync(() =>
        {
            var app = this.m_connection.GetWordApp();
            var doc = this.m_connection.FindDocument(app, this.Name);
            var revisions = doc.Revisions;
            return (revisions.Count, revisions.Count > 0 ? revisions[1].Author : null);
        });

    public void Dispose()
    {
        try
        {
            this.m_dispatcher.RunOnStaAsync(() =>
            {
                try
                {
                    this.m_word?.Quit(false);
                    this.m_word?.Dispose();
                }
                catch (Exception)
                {
                    // Best-effort teardown.
                }

                return true;
            }).GetAwaiter().GetResult();
        }
        catch (Exception)
        {
            // Ignore teardown failures.
        }

        this.m_connection.InvalidateCache();
        this.m_dispatcher.Dispose();

        try
        {
            File.Delete(this.m_path);
        }
        catch (Exception)
        {
            // Ignore.
        }
    }

    /// <summary>Small helper to read the JSON envelope fields in assertions.</summary>
    private sealed class AssertionsScope : IDisposable
    {
        private readonly JsonDocument m_doc;

        public AssertionsScope(string json)
        {
            this.m_doc = JsonDocument.Parse(json);
        }

        public bool Success => this.m_doc.RootElement.GetProperty("success").GetBoolean();

        public string String(string name) => this.m_doc.RootElement.GetProperty(name).GetString() ?? string.Empty;

        public int Int(string name) => this.m_doc.RootElement.GetProperty(name).GetInt32();

        public void Dispose() => this.m_doc.Dispose();
    }
}
