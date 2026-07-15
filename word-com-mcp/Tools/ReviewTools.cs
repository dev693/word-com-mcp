using System.ComponentModel;
using ModelContextProtocol.Server;
using WordComMcp.Com;
using WordComMcp.Infrastructure;
using Word = NetOffice.WordApi;

namespace WordComMcp.Tools;

/// <summary>Tier-2 comment and read-only revision workflow tools.</summary>
[McpServerToolType]
public static class ReviewTools
{
    [McpServerTool(Name = "word_live_add_comment")]
    [Description("Anchor a Word comment to final-view text. By default only the first occurrence is commented.")]
    public static Task<string> AddComment(
        StaDispatcher dispatcher,
        WordConnection connection,
        ServerConfig config,
        [Description("Literal final-view text to anchor.")] string anchorText,
        [Description("Comment text.")] string comment,
        [Description("True comments only the first match; false comments every match.")] bool firstOnly = true,
        [Description("Target document basename or full path. Null uses the active document.")] string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (app, doc) =>
        {
            if (string.IsNullOrEmpty(anchorText) || string.IsNullOrWhiteSpace(comment))
            {
                return McpResult.Err("anchorText and comment are required");
            }

            var matches = WordFind.LocateEditable(doc, anchorText, matchCase: false);
            if (matches.Count == 0)
            {
                return McpResult.Err($"anchorText '{anchorText}' not found");
            }

            var targets = firstOnly ? matches.Take(1) : matches;
            var ids = new List<int>();
            using (new UndoRecordScope(app, "MCP: Add Comment"))
            using (new FastEditScope(app))
            {
                TierTwoCom.WithAuthor(app, config, () =>
                {
                    foreach (var target in targets)
                    {
                        var added = doc.Comments.Add(doc.Range(target.Start, target.End), comment);
                        added.Author = config.Author;
                        added.Initial = config.AuthorInitials;
                        ids.Add(added.Index);
                    }

                    return true;
                });
            }

            return McpResult.Ok(new
            {
                document = doc.Name,
                added = ids.Count,
                commentIds = ids,
                author = config.Author,
            });
        });

    [McpServerTool(Name = "word_live_get_comments")]
    [Description("Return all Word comments, including author/date/text/scope/resolved state and replies.")]
    public static Task<string> GetComments(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Target document basename or full path. Null uses the active document.")] string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (_, doc) =>
            McpResult.Ok(new
            {
                document = doc.Name,
                count = doc.Comments.Count,
                comments = ReadComments(doc),
            }));

    [McpServerTool(Name = "word_live_reply_to_comment")]
    [Description("Reply to a 1-based comment index. Older Word versions fall back to a separate comment on the same scope.")]
    public static Task<string> ReplyToComment(
        StaDispatcher dispatcher,
        WordConnection connection,
        ServerConfig config,
        [Description("Current 1-based comment index.")] int commentId,
        [Description("Reply text.")] string reply,
        [Description("Target document basename or full path. Null uses the active document.")] string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (app, doc) =>
        {
            if (string.IsNullOrWhiteSpace(reply))
            {
                return McpResult.Err("reply is required");
            }

            var parent = TierTwoCom.ResolveComment(doc, commentId, out var error);
            if (parent is null)
            {
                return McpResult.Err(error!);
            }

            var threaded = true;
            var replyId = 0;
            var warnings = new List<string>();
            using (new UndoRecordScope(app, "MCP: Reply to Comment"))
            using (new FastEditScope(app))
            {
                TierTwoCom.WithAuthor(app, config, () =>
                {
                    try
                    {
                        var added = parent.Replies.Add(parent.Scope, reply);
                        added.Author = config.Author;
                        added.Initial = config.AuthorInitials;
                        replyId = added.Index;
                    }
                    catch (Exception)
                    {
                        threaded = false;
                        var added = doc.Comments.Add(parent.Scope, reply);
                        added.Author = config.Author;
                        added.Initial = config.AuthorInitials;
                        replyId = added.Index;
                        warnings.Add("threaded replies are unavailable; added a separate comment on the same scope");
                    }

                    return true;
                });
            }

            return McpResult.Ok(new
            {
                document = doc.Name,
                commentId,
                replyId,
                threaded,
                author = config.Author,
            }, warnings);
        });

    [McpServerTool(Name = "word_live_resolve_comment")]
    [Description("Mark a 1-based Word comment thread as resolved.")]
    public static Task<string> ResolveComment(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Current 1-based comment index.")] int commentId,
        [Description("Target document basename or full path. Null uses the active document.")] string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (app, doc) =>
        {
            var comment = TierTwoCom.ResolveComment(doc, commentId, out var error);
            if (comment is null)
            {
                return McpResult.Err(error!);
            }

            try
            {
                using (new UndoRecordScope(app, "MCP: Resolve Comment"))
                using (new FastEditScope(app))
                {
                    comment.Done = true;
                }
            }
            catch (Exception)
            {
                return McpResult.Err("comment resolve is not supported by this Word comment model");
            }

            return McpResult.Ok(new { document = doc.Name, commentId, resolved = comment.Done });
        });

    [McpServerTool(Name = "word_live_delete_comment")]
    [Description("Delete a 1-based Word comment. Re-read comments afterwards because indices shift.")]
    public static Task<string> DeleteComment(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Current 1-based comment index.")] int commentId,
        [Description("Target document basename or full path. Null uses the active document.")] string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (app, doc) =>
        {
            var comment = TierTwoCom.ResolveComment(doc, commentId, out var error);
            if (comment is null)
            {
                return McpResult.Err(error!);
            }

            var deletedText = TierTwoCom.CleanRangeText(comment.Range, 100);
            using (new UndoRecordScope(app, "MCP: Delete Comment"))
            using (new FastEditScope(app))
            {
                try
                {
                    comment.DeleteRecursively();
                }
                catch (Exception)
                {
                    comment.Delete();
                }
            }

            return McpResult.Ok(new
            {
                document = doc.Name,
                deletedCommentId = commentId,
                deletedText,
                remainingCount = doc.Comments.Count,
                comments = ReadComments(doc),
            });
        });

    [McpServerTool(Name = "word_live_list_revisions")]
    [Description("Read all tracked revisions with type, author, date, text, and effective character offsets. Never accepts or rejects revisions.")]
    public static Task<string> ListRevisions(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Target document basename or full path. Null uses the active document.")] string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (_, doc) =>
        {
            var revisions = new List<object>();
            for (var i = 1; i <= doc.Revisions.Count; i++)
            {
                var revision = doc.Revisions[i];
                var range = revision.Range;
                revisions.Add(new
                {
                    index = i,
                    type = NormalizeRevisionType(revision.Type.ToString()),
                    typeId = (int)revision.Type,
                    author = revision.Author ?? string.Empty,
                    date = revision.Date.ToString("O"),
                    text = TierTwoCom.CleanRangeText(range, 200),
                    start = range.Start,
                    end = range.End,
                });
            }

            return McpResult.Ok(new { document = doc.Name, count = revisions.Count, revisions });
        });

    private static IReadOnlyList<object> ReadComments(Word.Document doc)
    {
        var result = new List<object>();
        for (var i = 1; i <= doc.Comments.Count; i++)
        {
            var comment = doc.Comments[i];
            var replies = new List<object>();
            try
            {
                for (var j = 1; j <= comment.Replies.Count; j++)
                {
                    var reply = comment.Replies[j];
                    replies.Add(new
                    {
                        index = reply.Index,
                        author = reply.Author ?? string.Empty,
                        date = reply.Date.ToString("O"),
                        text = TierTwoCom.CleanRangeText(reply.Range),
                    });
                }
            }
            catch (Exception)
            {
                // Older Word versions expose no threaded reply collection.
            }

            var resolved = false;
            try
            {
                resolved = comment.Done;
            }
            catch (Exception)
            {
                // Modern Comments may not expose Done through COM.
            }

            result.Add(new
            {
                index = i,
                author = comment.Author ?? string.Empty,
                date = comment.Date.ToString("O"),
                text = TierTwoCom.CleanRangeText(comment.Range),
                scope = TierTwoCom.CleanRangeText(comment.Scope),
                resolved,
                replies,
            });
        }

        return result;
    }

    private static string NormalizeRevisionType(string value) =>
        value.StartsWith("wdRevision", StringComparison.Ordinal)
            ? ToSnakeCase(value["wdRevision".Length..])
            : ToSnakeCase(value);

    private static string ToSnakeCase(string value) =>
        string.Concat(value.SelectMany((ch, i) =>
            i > 0 && char.IsUpper(ch) ? new[] { '_', char.ToLowerInvariant(ch) } : new[] { char.ToLowerInvariant(ch) }));
}
