using System.ComponentModel;
using ModelContextProtocol.Server;
using WordComMcp.Com;
using WordComMcp.Infrastructure;
using WordComMcp.Markdown;
using Word = NetOffice.WordApi;

namespace WordComMcp.Tools;

/// <summary>
/// Surgical tracked-edit Tier-1 tools: insert adjusted-Markdown (issue 1.4), and the
/// minimal-redline replace (1.5) and delete (1.7). Every mutation runs inside a single
/// undo record, a FastEdit scope, and an optional tracked-changes override.
/// </summary>
[McpServerToolType]
public static class EditTools
{
    [McpServerTool(Name = "word_live_insert_markdown")]
    [Description(
        "Insert an adjusted-Markdown block into the document as new paragraphs. Anchor with " +
        "afterText/beforeText (matched against the final text) or atEnd. Applies German " +
        "Formatvorlagen (headings, lists, quote, {style=…} spans/paragraphs). Under tracking it " +
        "is one clean insertion revision reverted by a single Ctrl+Z. Unknown style → structured error.")]
    public static Task<string> InsertMarkdown(
        StaDispatcher dispatcher,
        WordConnection connection,
        StyleMap styleMap,
        ServerConfig config,
        [Description("The adjusted-Markdown to realize.")] string markdown,
        [Description("Insert after the paragraph containing this text. Null = unused.")] string? afterText = null,
        [Description("Insert before the paragraph containing this text. Null = unused.")] string? beforeText = null,
        [Description("Insert at the end of the document. Default false.")] bool atEnd = false,
        [Description("Override tracked-changes for this edit (true/false); null keeps the current setting.")]
        bool? trackChanges = null,
        [Description("Target document basename or full path. Null/empty uses the active document.")]
        string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (app, doc) =>
        {
            var model = MarkdownParser.Parse(markdown);
            var insertionPoint = ResolveInsertionPoint(doc, afterText, beforeText, atEnd, out var error);
            if (insertionPoint is null)
            {
                return McpResult.Err(error!);
            }

            using (new UndoRecordScope(app, "MCP: Insert Markdown"))
            using (new FastEditScope(app))
            {
                return TrackChangesScope.With(app, doc, config, trackChanges, () =>
                {
                    var result = MarkdownRealizer.Realize(doc, insertionPoint, model, styleMap);
                    return McpResult.Ok(
                        new
                        {
                            document = doc.Name,
                            paragraphs = result.Paragraphs,
                            tracked = doc.TrackRevisions,
                        },
                        warnings: result.Warnings);
                });
            }
        });

    [McpServerTool(Name = "word_live_replace_text")]
    [Description(
        "Replace occurrences of literal text with a minimal ins/del redline (never a whole-range " +
        "re-insert). occurrence: null = first match, N = the 1-based Nth match, 0 = all matches. " +
        "scope limits the search to a heading section. The replacement may contain inline markdown " +
        "(**bold**, *italic*, `code`, [text]{style=\"…\"}). Long patterns (>255 chars) are chunked.")]
    public static Task<string> ReplaceText(
        StaDispatcher dispatcher,
        WordConnection connection,
        StyleMap styleMap,
        ServerConfig config,
        [Description("Literal text to find (against the final view).")] string find,
        [Description("Replacement text; may contain inline markdown formatting.")] string replacement,
        [Description("null = first match, N = the 1-based Nth match, 0 = all matches.")] int? occurrence = null,
        [Description("Limit the search to the section under this heading text. Null = whole document.")]
        string? scope = null,
        [Description("Override tracked-changes for this edit (true/false); null keeps the current setting.")]
        bool? trackChanges = null,
        [Description("Target document basename or full path. Null/empty uses the active document.")]
        string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (app, doc) =>
        {
            var scopeRange = ResolveScope(doc, scope, styleMap, out var scopeError);
            if (scopeError is not null)
            {
                return McpResult.Err(scopeError);
            }

            var hits = WordFind.LocateEditable(doc, find, matchCase: false, scopeRange);
            var targets = SelectOccurrences(hits, occurrence);
            if (targets.Count == 0)
            {
                return McpResult.Ok(new { document = doc.Name, replaced = 0 });
            }

            using (new UndoRecordScope(app, "MCP: Replace Text"))
            using (new FastEditScope(app))
            {
                return TrackChangesScope.With(app, doc, config, trackChanges, () =>
                {
                    var warnings = new List<string>();

                    // Edit from the end backwards so earlier offsets stay valid.
                    foreach (var target in targets.OrderByDescending(t => t.Start))
                    {
                        var range = doc.Range(target.Start, target.End);
                        warnings.AddRange(MarkdownRealizer.ReplaceRange(doc, range, replacement, styleMap));
                    }

                    return McpResult.Ok(
                        new { document = doc.Name, replaced = targets.Count, tracked = doc.TrackRevisions },
                        warnings: warnings);
                });
            }
        });

    [McpServerTool(Name = "word_live_delete_text")]
    [Description(
        "Delete occurrences of literal text as a clean tracked deletion (w:del under tracking). " +
        "occurrence: null = first match, N = the 1-based Nth match, 0 = all matches. scope limits " +
        "the search to a heading section.")]
    public static Task<string> DeleteText(
        StaDispatcher dispatcher,
        WordConnection connection,
        StyleMap styleMap,
        ServerConfig config,
        [Description("Literal text to find and delete (against the final view).")] string find,
        [Description("null = first match, N = the 1-based Nth match, 0 = all matches.")] int? occurrence = null,
        [Description("Limit the search to the section under this heading text. Null = whole document.")]
        string? scope = null,
        [Description("Override tracked-changes for this edit (true/false); null keeps the current setting.")]
        bool? trackChanges = null,
        [Description("Target document basename or full path. Null/empty uses the active document.")]
        string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (app, doc) =>
        {
            var scopeRange = ResolveScope(doc, scope, styleMap, out var scopeError);
            if (scopeError is not null)
            {
                return McpResult.Err(scopeError);
            }

            var hits = WordFind.LocateEditable(doc, find, matchCase: false, scopeRange);
            var targets = SelectOccurrences(hits, occurrence);
            if (targets.Count == 0)
            {
                return McpResult.Ok(new { document = doc.Name, deleted = 0 });
            }

            using (new UndoRecordScope(app, "MCP: Delete Text"))
            using (new FastEditScope(app))
            {
                return TrackChangesScope.With(app, doc, config, trackChanges, () =>
                {
                    foreach (var target in targets.OrderByDescending(t => t.Start))
                    {
                        doc.Range(target.Start, target.End).Delete();
                    }

                    return McpResult.Ok(new { document = doc.Name, deleted = targets.Count, tracked = doc.TrackRevisions });
                });
            }
        });

    private static IReadOnlyList<WordFind.MatchRange> SelectOccurrences(
        IReadOnlyList<WordFind.MatchRange> hits,
        int? occurrence)
    {
        if (hits.Count == 0)
        {
            return Array.Empty<WordFind.MatchRange>();
        }

        // null → first match; 0 → all matches; N → the 1-based Nth match.
        if (occurrence is null)
        {
            return new[] { hits[0] };
        }

        if (occurrence.Value == 0)
        {
            return hits;
        }

        var index = occurrence.Value - 1;
        return index >= 0 && index < hits.Count ? new[] { hits[index] } : Array.Empty<WordFind.MatchRange>();
    }

    private static Word.Range? ResolveScope(Word.Document doc, string? scope, StyleMap styleMap, out string? error)
    {
        error = null;
        if (string.IsNullOrWhiteSpace(scope))
        {
            return null;
        }

        var headingNames = new[]
        {
            styleMap.DefaultLocalName(MarkdownConstruct.Heading1),
            styleMap.DefaultLocalName(MarkdownConstruct.Heading2),
            styleMap.DefaultLocalName(MarkdownConstruct.Heading3),
        };

        var resolved = DocumentRanges.ResolveHeadingSection(doc, scope, headingNames);
        if (resolved is null)
        {
            error = $"scope '{scope}' not found";
        }

        return resolved;
    }

    private static Word.Range? ResolveInsertionPoint(
        Word.Document doc,
        string? afterText,
        string? beforeText,
        bool atEnd,
        out string? error)
    {
        error = null;

        if (!string.IsNullOrEmpty(afterText))
        {
            var hits = WordFind.LocateEditable(doc, afterText, matchCase: false);
            if (hits.Count == 0)
            {
                error = $"afterText '{afterText}' not found";
                return null;
            }

            var paragraph = doc.Range(hits[0].End, hits[0].End).Paragraphs[1].Range;
            return doc.Range(paragraph.End, paragraph.End);
        }

        if (!string.IsNullOrEmpty(beforeText))
        {
            var hits = WordFind.LocateEditable(doc, beforeText, matchCase: false);
            if (hits.Count == 0)
            {
                error = $"beforeText '{beforeText}' not found";
                return null;
            }

            var paragraph = doc.Range(hits[0].Start, hits[0].Start).Paragraphs[1].Range;
            return doc.Range(paragraph.Start, paragraph.Start);
        }

        // Default (including atEnd): append at the end of the document.
        _ = atEnd;
        var end = doc.Content.End;
        return doc.Range(end, end);
    }
}
