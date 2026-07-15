using System.Text.Json;
using WordComMcp.Com;
using WordComMcp.Infrastructure;
using WordComMcp.Markdown;
using WordComMcp.Tests.Fixtures;
using WordComMcp.Tools;
using Word = NetOffice.WordApi;
using WdAlertLevel = NetOffice.WordApi.Enums.WdAlertLevel;
using Xunit;

namespace WordComMcp.Tests;

[Trait("Category", "Live")]
[Collection("Live Word")]
public sealed class LiveTierTwoToolsTests : IDisposable
{
    private readonly StaDispatcher m_dispatcher = new();
    private readonly WordConnection m_connection;
    private readonly StyleMap m_styles = StyleMap.Default;
    private readonly ServerConfig m_config = new("Tier Two Tester", "T2", null);
    private readonly string m_path;
    private Word.Application? m_word;
    private Word.Document? m_document;

    public LiveTierTwoToolsTests()
    {
        this.m_connection = new WordConnection(
            () => this.m_word ?? throw new InvalidOperationException("The isolated Word instance is not initialized."));
        this.m_path = Path.Combine(Path.GetTempPath(), $"wl-mcp-tier2-{Guid.NewGuid():N}.docx");
        FixtureBuilder.WriteRichFixture(this.m_path);
        this.m_dispatcher.RunOnStaAsync(() =>
        {
            this.m_word = new Word.Application { Visible = false, DisplayAlerts = WdAlertLevel.wdAlertsNone };
            this.m_document = this.m_word.Documents.Open(this.m_path);
            this.m_document.TrackRevisions = true;
            var anchor = this.m_document.Paragraphs[2].Range;
            this.m_document.Bookmarks.Add("Testmarke", anchor);
            return true;
        }).GetAwaiter().GetResult();
    }

    private string Name => Path.GetFileName(this.m_path);

    [Fact]
    public async Task SetBlockStyle_RecordsCustomAndHeadingFormattingRevisions()
    {
        var before = await this.RevisionCountAsync();

        var custom = await StructureTools.SetBlockStyle(
            this.m_dispatcher, this.m_connection,
            anchorText: "Ein wichtiger", style: "{style=\"Merksatz\"}", filename: this.Name);
        var heading = await StructureTools.SetBlockStyle(
            this.m_dispatcher, this.m_connection,
            anchorText: "Wichtiger Merksatz", level: 3, filename: this.Name);

        AssertSuccess(custom);
        AssertSuccess(heading);
        Assert.True(await this.RevisionCountAsync() > before);
        Assert.Equal("Merksatz", await this.StyleForAsync("Ein wichtiger"));
        Assert.Contains("3", await this.StyleForAsync("Wichtiger Merksatz"));
    }

    [Fact]
    public async Task ApplyMarkdown_UsesMinimalTextRevision_AndRejectsListRenumber()
    {
        var read = await ReadTools.GetMarkdown(
            this.m_dispatcher, this.m_connection, this.m_styles,
            scope: "section", anchor: "Kapitel Eins", filename: this.Name);
        var markdown = Property(read, "markdown");
        var requested = markdown.Replace("wichtiger", "entscheidender", StringComparison.Ordinal);

        var applied = await StructureTools.ApplyMarkdown(
            this.m_dispatcher, this.m_connection, this.m_styles, this.m_config,
            "section", requested, "Kapitel Eins", this.Name);

        AssertSuccess(applied);
        Assert.NotEmpty(JsonDocument.Parse(applied).RootElement.GetProperty("changes").EnumerateArray());
        var finalText = await this.FinalTextAsync();
        Assert.Contains("entscheidender", finalText);
        Assert.DoesNotContain("wichtiger", finalText);

        var beforeWarning = await this.RevisionCountAsync();
        var renumbered = requested
            .Replace("- Erster Punkt", "1. Erster Punkt", StringComparison.Ordinal)
            .Replace("- Zweiter Punkt", "2. Zweiter Punkt", StringComparison.Ordinal);
        var warning = await StructureTools.ApplyMarkdown(
            this.m_dispatcher, this.m_connection, this.m_styles, this.m_config,
            "section", renumbered, "Kapitel Eins", this.Name);

        AssertSuccess(warning);
        Assert.True(JsonDocument.Parse(warning).RootElement.TryGetProperty("warnings", out var warnings));
        Assert.NotEmpty(warnings.EnumerateArray());
        Assert.Equal(beforeWarning, await this.RevisionCountAsync());
    }

    [Fact]
    public async Task CommentWorkflow_AddsReadsRepliesResolvesAndDeletes()
    {
        var added = await ReviewTools.AddComment(
            this.m_dispatcher, this.m_connection, this.m_config,
            "config.json", "Bitte prüfen", filename: this.Name);
        AssertSuccess(added);

        var read = await ReviewTools.GetComments(this.m_dispatcher, this.m_connection, this.Name);
        AssertSuccess(read);
        Assert.Contains("Tier Two Tester", read);
        Assert.Contains("config.json", read);

        var replied = await ReviewTools.ReplyToComment(
            this.m_dispatcher, this.m_connection, this.m_config,
            1, "Ist geprüft", this.Name);
        AssertSuccess(replied);

        var resolved = await ReviewTools.ResolveComment(this.m_dispatcher, this.m_connection, 1, this.Name);
        using (var resolvedJson = JsonDocument.Parse(resolved))
        {
            Assert.True(
                resolvedJson.RootElement.GetProperty("success").GetBoolean() ||
                resolvedJson.RootElement.GetProperty("error").GetString() ==
                    "comment resolve is not supported by this Word comment model",
                resolved);
        }

        var deleted = await ReviewTools.DeleteComment(this.m_dispatcher, this.m_connection, 1, this.Name);
        AssertSuccess(deleted);
        Assert.Equal(0, JsonDocument.Parse(deleted).RootElement.GetProperty("remainingCount").GetInt32());
    }

    [Fact]
    public async Task RevisionHeadingNavigationAndDocumentTools_ReportExpectedState()
    {
        await StructureTools.SetBlockStyle(
            this.m_dispatcher, this.m_connection,
            anchorText: "Wichtiger Merksatz", level: 3, filename: this.Name);

        var revisions = await ReviewTools.ListRevisions(this.m_dispatcher, this.m_connection, this.Name);
        var headings = await NavigationTools.ListHeadings(this.m_dispatcher, this.m_connection, this.Name);
        var headingGoto = await NavigationTools.GoTo(
            this.m_dispatcher, this.m_connection, headingText: "Kapitel Eins", filename: this.Name);
        var bookmarkGoto = await NavigationTools.GoTo(
            this.m_dispatcher, this.m_connection, bookmark: "Testmarke", filename: this.Name);
        var pageGoto = await NavigationTools.GoTo(
            this.m_dispatcher, this.m_connection, page: 1, filename: this.Name);
        var open = await NavigationTools.ListOpen(this.m_dispatcher, this.m_connection);
        var properties = await NavigationTools.SetCoreProperties(
            this.m_dispatcher, this.m_connection,
            title: "Tier 2", author: "MCP", filename: this.Name);

        AssertSuccess(revisions);
        Assert.Contains("start", revisions);
        AssertSuccess(headings);
        Assert.Contains("Kapitel Eins", headings);
        AssertSuccess(headingGoto);
        AssertSuccess(bookmarkGoto);
        AssertSuccess(pageGoto);
        AssertSuccess(open);
        Assert.Contains("\"active\":true", open);
        AssertSuccess(properties);
        Assert.Contains("title", properties);
        Assert.Equal("Tier 2", await this.CorePropertyAsync("Title"));
        Assert.Equal("MCP", await this.CorePropertyAsync("Author"));
    }

    private async Task<int> RevisionCountAsync() =>
        await this.m_dispatcher.RunOnStaAsync(() => this.m_document!.Revisions.Count);

    private async Task<string> FinalTextAsync() =>
        await this.m_dispatcher.RunOnStaAsync(() => OoxmlFinalText.GetFinalText(this.m_document!.Content));

    private async Task<string> StyleForAsync(string text) =>
        await this.m_dispatcher.RunOnStaAsync(() =>
        {
            var match = WordFind.LocateEditable(this.m_document!, text, false).First();
            return DocumentRanges.StyleNameOf(this.m_document!.Range(match.Start, match.End));
        });

    private async Task<string> CorePropertyAsync(string name) =>
        await this.m_dispatcher.RunOnStaAsync(() =>
        {
            dynamic properties = this.m_document!.BuiltInDocumentProperties;
            return (string)properties[name].Value;
        });

    private static void AssertSuccess(string json) =>
        Assert.True(JsonDocument.Parse(json).RootElement.GetProperty("success").GetBoolean(), json);

    private static string Property(string json, string name) =>
        JsonDocument.Parse(json).RootElement.GetProperty(name).GetString() ?? string.Empty;

    private static void Teardown(Action action)
    {
        try
        {
            action();
        }
        catch (Exception)
        {
            // Best effort for isolated test resources.
        }
    }

    public void Dispose()
    {
        try
        {
            this.m_dispatcher.RunOnStaAsync(() =>
            {
                Teardown(() => this.m_document?.Close(false));
                Teardown(() => this.m_document?.Dispose());
                Teardown(() => this.m_word?.Quit(false));
                Teardown(() => this.m_word?.Dispose());
                return true;
            }).GetAwaiter().GetResult();
        }
        catch (Exception)
        {
            // Best effort.
        }

        this.m_dispatcher.Dispose();
        Teardown(() => File.Delete(this.m_path));
    }
}
