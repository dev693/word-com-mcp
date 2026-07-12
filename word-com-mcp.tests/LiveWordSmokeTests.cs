using WordComMcp.Com;
using WordComMcp.Tests.Fixtures;
using Word = NetOffice.WordApi;
using Xunit;

namespace WordComMcp.Tests;

/// <summary>
/// Live smoke test against a running Word instance (issue 0.11). Excluded from the default
/// run (<c>dotnet test --filter Category!=Live</c>); run explicitly on a machine with Word:
/// <c>dotnet test --filter Category=Live</c>. All COM work happens on the shared STA thread.
/// </summary>
[Trait("Category", "Live")]
public class LiveWordSmokeTests
{
    [Fact]
    public async Task ConnectsToWord_AndFindsFixtureByNameAndFullPath()
    {
        // Arrange — write the German fixture to a temp file we can open in Word.
        var path = Path.Combine(Path.GetTempPath(), $"wl-mcp-{Guid.NewGuid():N}.docx");
        FixtureBuilder.WriteGermanFixture(path);

        using var dispatcher = new StaDispatcher();
        try
        {
            var result = await dispatcher.RunOnStaAsync(() => RunSmoke(path));

            // Assert — 0.5: resolves a running Word with documents; finds by basename and full path.
            Assert.True(result.ConnectedHasDocuments);
            Assert.Equal(Path.GetFileName(path), result.FoundByBasename);
            Assert.Equal(Path.GetFileName(path), result.FoundByFullPath);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static SmokeResult RunSmoke(string path)
    {
        var word = new Word.Application { Visible = false, DisplayAlerts = Word.Enums.WdAlertLevel.wdAlertsNone };
        try
        {
            var opened = word.Documents.Open(path);
            var connection = new WordConnection();

            // 0.5 — connect to the running instance and confirm it has documents.
            var app = connection.GetWordApp();
            var connectedHasDocuments = app.Documents.Count > 0;

            // 0.5 — find the document by basename and by normalized full path against our instance.
            var byBasename = connection.FindDocument(word, Path.GetFileName(path)).Name;
            var byFullPath = connection.FindDocument(word, opened.FullName).Name;

            return new SmokeResult(connectedHasDocuments, byBasename, byFullPath);
        }
        finally
        {
            // 0.6 — we started this Word instance, so shut it down; NetOffice releases the proxy tree.
            word.Quit(false);
            word.Dispose();
        }
    }

    private sealed record SmokeResult(bool ConnectedHasDocuments, string FoundByBasename, string FoundByFullPath);
}
