using System.ComponentModel;
using ModelContextProtocol.Server;
using WordComMcp.Com;
using WordComMcp.Infrastructure;
using Word = NetOffice.WordApi;
using WdGoToDirection = NetOffice.WordApi.Enums.WdGoToDirection;
using WdGoToItem = NetOffice.WordApi.Enums.WdGoToItem;

namespace WordComMcp.Tools;

/// <summary>Tier-2 document enumeration, outline navigation, and core-property tools.</summary>
[McpServerToolType]
public static class NavigationTools
{
    [McpServerTool(Name = "word_live_list_open")]
    [Description("List all documents open in the connected Word application and mark the active document.")]
    public static Task<string> ListOpen(StaDispatcher dispatcher, WordConnection connection) =>
        ToolExecution.RunApplicationAsync(dispatcher, connection, app =>
        {
            string? activePath = null;
            try
            {
                activePath = app.ActiveDocument.FullName;
            }
            catch (Exception)
            {
                // An application with no active document is still a valid empty result.
            }

            var documents = new List<object>();
            for (var i = 1; i <= app.Documents.Count; i++)
            {
                var doc = app.Documents[i];
                var fullName = Safe(() => doc.FullName);
                documents.Add(new
                {
                    index = i,
                    name = Safe(() => doc.Name),
                    fullName,
                    active = activePath is not null && string.Equals(fullName, activePath, StringComparison.OrdinalIgnoreCase),
                });
            }

            return McpResult.Ok(new { count = documents.Count, documents });
        });

    [McpServerTool(Name = "word_live_list_headings")]
    [Description("Return the final-view heading outline using Word OutlineLevel (levels 1-9).")]
    public static Task<string> ListHeadings(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Target document basename or full path. Null uses the active document.")] string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (_, doc) =>
        {
            var headings = TierTwoCom.ReadHeadings(doc);
            return McpResult.Ok(new
            {
                document = doc.Name,
                count = headings.Count,
                headings = headings.Select(h => new { level = h.Level, text = h.Text, start = h.Start }),
            });
        });

    [McpServerTool(Name = "word_live_goto")]
    [Description("Move Word's cursor to exactly one heading, page, or bookmark target.")]
    public static Task<string> GoTo(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Heading text from word_live_list_headings.")] string? headingText = null,
        [Description("1-based page number.")] int? page = null,
        [Description("Bookmark name.")] string? bookmark = null,
        [Description("Target document basename or full path. Null uses the active document.")] string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (app, doc) =>
        {
            var targets = (!string.IsNullOrWhiteSpace(headingText) ? 1 : 0) +
                          (page.HasValue ? 1 : 0) +
                          (!string.IsNullOrWhiteSpace(bookmark) ? 1 : 0);
            if (targets != 1)
            {
                return McpResult.Err("provide exactly one of headingText, page, or bookmark");
            }

            doc.Activate();
            var start = 0;
            string kind;
            string value;
            if (!string.IsNullOrWhiteSpace(headingText))
            {
                var heading = TierTwoCom.ReadHeadings(doc).FirstOrDefault(h =>
                    string.Equals(h.Text, headingText.Trim(), StringComparison.OrdinalIgnoreCase));
                if (heading is null)
                {
                    return McpResult.Err($"heading '{headingText}' not found");
                }

                doc.Range(heading.Start, heading.End).Select();
                start = heading.Start;
                kind = "heading";
                value = heading.Text;
            }
            else if (page.HasValue)
            {
                if (page.Value < 1)
                {
                    return McpResult.Err("page must be at least 1");
                }

                var range = app.Selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, page.Value);
                start = range.Start;
                kind = "page";
                value = page.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            else
            {
                if (!doc.Bookmarks.Exists(bookmark!))
                {
                    return McpResult.Err($"bookmark '{bookmark}' not found");
                }

                var range = doc.Bookmarks[bookmark!].Range;
                range.Select();
                start = range.Start;
                kind = "bookmark";
                value = bookmark!;
            }

            return McpResult.Ok(new { document = doc.Name, target = kind, value, start });
        });

    [McpServerTool(Name = "word_live_set_core_properties")]
    [Description("Update selected built-in document properties. Null leaves a property unchanged; an empty string clears it.")]
    public static Task<string> SetCoreProperties(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Document title; null leaves unchanged.")] string? title = null,
        [Description("Document author; null leaves unchanged.")] string? author = null,
        [Description("Document subject; null leaves unchanged.")] string? subject = null,
        [Description("Document keywords; null leaves unchanged.")] string? keywords = null,
        [Description("Target document basename or full path. Null uses the active document.")] string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (app, doc) =>
        {
            var updated = new List<string>();
            using (new UndoRecordScope(app, "MCP: Set Core Properties"))
            using (new FastEditScope(app))
            {
                dynamic properties = doc.BuiltInDocumentProperties;
                Set(properties, "Title", title, updated);
                Set(properties, "Author", author, updated);
                Set(properties, "Subject", subject, updated);
                Set(properties, "Keywords", keywords, updated);
            }

            return McpResult.Ok(new { document = doc.Name, updated });
        });

    private static void Set(dynamic properties, string name, string? value, List<string> updated)
    {
        if (value is null)
        {
            return;
        }

        properties[name].Value = value;
        updated.Add(name.ToLowerInvariant());
    }

    private static string Safe(Func<string> read)
    {
        try
        {
            return read();
        }
        catch (Exception)
        {
            return string.Empty;
        }
    }
}
