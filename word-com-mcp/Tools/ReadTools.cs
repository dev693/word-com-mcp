using System.ComponentModel;
using ModelContextProtocol.Server;
using WordComMcp.Com;
using WordComMcp.Infrastructure;
using WordComMcp.Markdown;
using Word = NetOffice.WordApi;

namespace WordComMcp.Tools;

/// <summary>
/// Read/query Tier-1 tools: export adjusted-Markdown of the final view (issue 1.3) and
/// search the final text (issue 1.6). Both are read-only — no revision is created.
/// </summary>
[McpServerToolType]
public static class ReadTools
{
    [McpServerTool(Name = "word_live_get_markdown")]
    [Description(
        "Export the document (or a heading section) as adjusted-Markdown. By default returns the " +
        "FINAL view (tracked insertions included, deletions excluded). Custom Formatvorlagen are " +
        "emitted as {style=\"…\"} annotations so they survive a read → edit → write round-trip.")]
    public static Task<string> GetMarkdown(
        StaDispatcher dispatcher,
        WordConnection connection,
        StyleMap styleMap,
        [Description("What to export: \"doc\" (default, whole document) or \"section\" (requires anchor).")]
        string scope = "doc",
        [Description("Heading text to slice a section from (through the next same-or-higher heading). Null = whole document.")]
        string? anchor = null,
        [Description("Revision view: \"final\" (default), \"original\", or \"markup\".")]
        string view = "final",
        [Description("Target document basename or full path. Null/empty uses the active document.")]
        string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (_, doc) =>
        {
            if (!TryParseView(view, out var revisionView))
            {
                return McpResult.Err($"invalid view '{view}' (expected final|original|markup)");
            }

            var range = ResolveRange(doc, scope, anchor, styleMap);
            if (range is null)
            {
                return McpResult.Err($"anchor '{anchor}' not found");
            }

            var model = OoxmlMarkdownReader.Read(range.WordOpenXML, styleMap, revisionView);
            var markdown = new MarkdownSerializer(styleMap).Serialize(model);
            return McpResult.Ok(new
            {
                document = doc.Name,
                scope = string.IsNullOrWhiteSpace(anchor) ? "doc" : "section",
                view = revisionView.ToString().ToLowerInvariant(),
                markdown,
            });
        });

    [McpServerTool(Name = "word_live_find_text")]
    [Description(
        "Search the document's FINAL text for a literal query (max 255 chars). Returns deduped " +
        "hits with character offsets, a short context snippet, and each hit's revision state " +
        "(clean|inside-insertion|inside-deletion). Read-only — creates no revision.")]
    public static Task<string> FindText(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Literal text to search for (max 255 characters).")] string query,
        [Description("Match case exactly. Default false.")] bool matchCase = false,
        [Description("Maximum number of hits to return. Default 50.")] int maxResults = 50,
        [Description("Target document basename or full path. Null/empty uses the active document.")]
        string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (_, doc) =>
        {
            var limit = maxResults <= 0 ? 50 : maxResults;
            var hits = WordFind.Search(doc, query, matchCase, limit);
            return McpResult.Ok(new
            {
                document = doc.Name,
                query,
                count = hits.Count,
                matches = hits.Select(h => new
                {
                    start = h.Start,
                    end = h.End,
                    context = h.Context,
                    revisionState = h.RevisionState,
                }),
            });
        });

    private static bool TryParseView(string view, out RevisionView revisionView)
    {
        switch (view?.Trim().ToLowerInvariant())
        {
            case null:
            case "":
            case "final":
                revisionView = RevisionView.Final;
                return true;
            case "original":
                revisionView = RevisionView.Original;
                return true;
            case "markup":
                revisionView = RevisionView.Markup;
                return true;
            default:
                revisionView = RevisionView.Final;
                return false;
        }
    }

    private static Word.Range? ResolveRange(Word.Document doc, string scope, string? anchor, StyleMap styleMap)
    {
        if (string.IsNullOrWhiteSpace(anchor))
        {
            return doc.Content;
        }

        var headingNames = new[]
        {
            styleMap.DefaultLocalName(MarkdownConstruct.Heading1),
            styleMap.DefaultLocalName(MarkdownConstruct.Heading2),
            styleMap.DefaultLocalName(MarkdownConstruct.Heading3),
        };

        return DocumentRanges.ResolveHeadingSection(doc, anchor!, headingNames);
    }
}
