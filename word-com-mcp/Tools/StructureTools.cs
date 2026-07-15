using System.ComponentModel;
using ModelContextProtocol.Server;
using WordComMcp.Com;
using WordComMcp.Infrastructure;
using WordComMcp.Markdown;
using Word = NetOffice.WordApi;
using WdBuiltinStyle = NetOffice.WordApi.Enums.WdBuiltinStyle;

namespace WordComMcp.Tools;

/// <summary>Tier-2 paragraph styling and conservative Markdown restructuring tools.</summary>
[McpServerToolType]
public static class StructureTools
{
    private sealed record LiveBlockRange(int Start, int End);

    [McpServerTool(Name = "word_live_set_block_style")]
    [Description("Change one paragraph's Formatvorlage or heading level as a formatting revision. Provide exactly one target and one of style/level.")]
    public static Task<string> SetBlockStyle(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Text anchoring the target paragraph.")] string? anchorText = null,
        [Description("1-based target paragraph index.")] int? paragraphIndex = null,
        [Description("Localized paragraph style name or {style=\"…\"} annotation.")] string? style = null,
        [Description("Heading level 1, 2, or 3.")] int? level = null,
        [Description("Target document basename or full path. Null uses the active document.")] string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (app, doc) =>
        {
            if ((!string.IsNullOrWhiteSpace(style) ? 1 : 0) + (level.HasValue ? 1 : 0) != 1)
            {
                return McpResult.Err("provide exactly one of style or level");
            }

            if (level is < 1 or > 3)
            {
                return McpResult.Err("level must be between 1 and 3");
            }

            var target = TierTwoCom.ResolveParagraph(doc, anchorText, paragraphIndex, out var error);
            if (target is null)
            {
                return McpResult.Err(error!);
            }

            var requestedStyle = level.HasValue ? null : NormalizeStyle(style!);
            if (requestedStyle is not null)
            {
                StyleMap.ValidateStyle(requestedStyle, DocumentStyles.LocalNames(doc));
            }

            var before = doc.Revisions.Count;
            using (new UndoRecordScope(app, "MCP: Set Block Style"))
            using (new FastEditScope(app))
            {
                if (level.HasValue)
                {
                    target.Style = doc.Styles[HeadingStyle(level.Value)];
                }
                else
                {
                    target.Style = requestedStyle!;
                }
            }

            var actual = DocumentRanges.StyleNameOf(target);
            var expected = level.HasValue ? doc.Styles[HeadingStyle(level.Value)].NameLocal : requestedStyle!;
            var warnings = string.Equals(actual, expected, StringComparison.Ordinal)
                ? Array.Empty<string>()
                : [$"style '{expected}' did not apply (got '{actual}')"];
            return McpResult.Ok(new
            {
                document = doc.Name,
                start = target.Start,
                style = actual,
                level,
                revisionsAdded = Math.Max(0, doc.Revisions.Count - before),
            }, warnings);
        });

    [McpServerTool(Name = "word_live_apply_markdown")]
    [Description("Diff final adjusted-Markdown against newMarkdown and apply only safe, targeted tracked edits. Unsafe list/table/ambiguous structural changes return warnings and apply nothing.")]
    public static Task<string> ApplyMarkdown(
        StaDispatcher dispatcher,
        WordConnection connection,
        StyleMap styleMap,
        ServerConfig config,
        [Description("\"doc\" for the whole document or \"section\" for a heading section.")] string scope,
        [Description("Requested adjusted-Markdown for the scope.")] string newMarkdown,
        [Description("Heading text required when scope is section.")] string? anchor = null,
        [Description("Target document basename or full path. Null uses the active document.")] string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (app, doc) =>
        {
            var scopeRange = ResolveScope(doc, scope, anchor, styleMap, out var scopeError);
            if (scopeRange is null)
            {
                return McpResult.Err(scopeError!);
            }

            var current = OoxmlMarkdownReader.Read(scopeRange.WordOpenXML, styleMap, RevisionView.Final);
            var requested = MarkdownParser.Parse(newMarkdown);
            var plan = MarkdownDiffPlanner.Plan(current, requested, styleMap);
            if (!plan.CanApply)
            {
                return McpResult.Ok(new { document = doc.Name, changes = Array.Empty<object>() }, plan.Warnings);
            }

            var mappingWarnings = new List<string>();
            var ranges = MapBlocks(doc, scopeRange, current, mappingWarnings);
            if (mappingWarnings.Count > 0)
            {
                return McpResult.Ok(new { document = doc.Name, changes = Array.Empty<object>() }, mappingWarnings);
            }

            var applied = new List<object>();
            var warnings = new List<string>();
            using (new UndoRecordScope(app, "MCP: Apply Markdown"))
            using (new FastEditScope(app))
            {
                TierTwoCom.WithAuthor(app, config, () =>
                {
                    foreach (var change in plan.Changes
                        .OrderByDescending(ChangePosition)
                        .ThenByDescending(c => c.NewIndex))
                    {
                        ApplyChange(doc, scopeRange, current, requested, ranges, change, styleMap, applied, warnings);
                    }

                    return true;
                });
            }

            return McpResult.Ok(new { document = doc.Name, changes = applied }, warnings);

            int ChangePosition(MarkdownDiffChange change) =>
                change.OldIndex >= 0 && change.OldIndex < ranges.Count
                    ? ranges[change.OldIndex].Start
                    : scopeRange.End;
        });

    private static void ApplyChange(
        Word.Document doc,
        Word.Range scopeRange,
        MarkdownDocument current,
        MarkdownDocument requested,
        IReadOnlyList<LiveBlockRange> ranges,
        MarkdownDiffChange change,
        StyleMap styleMap,
        List<object> applied,
        List<string> warnings)
    {
        if (change.Kind == "insert")
        {
            var position = change.OldIndex < ranges.Count ? ranges[change.OldIndex].Start : scopeRange.End;
            var insertion = doc.Range(position, position);
            var block = requested.Blocks[change.NewIndex];
            var result = MarkdownRealizer.Realize(doc, insertion, new MarkdownDocument([block]), styleMap);
            warnings.AddRange(result.Warnings);
            applied.Add(new { kind = "insert", newBlock = change.NewIndex + 1, start = position });
            return;
        }

        var live = ranges[change.OldIndex];
        if (change.Kind == "delete")
        {
            doc.Range(live.Start, live.End).Delete();
            applied.Add(new { kind = "delete", oldBlock = change.OldIndex + 1, start = live.Start, end = live.End });
            return;
        }

        var oldBlock = current.Blocks[change.OldIndex];
        var newBlock = requested.Blocks[change.NewIndex];
        if (oldBlock is ListBlock oldList && newBlock is ListBlock newList)
        {
            ApplyListChange(doc, live, oldList, newList, styleMap, warnings);
        }
        else
        {
            var target = doc.Range(live.Start, live.End);
            if (change.Kind == "replace")
            {
                ApplyMinimalTextChange(doc, target, MarkdownDiffPlanner.PlainText(oldBlock), MarkdownDiffPlanner.PlainText(newBlock), styleMap);
            }

            warnings.AddRange(MarkdownRealizer.ApplyBlockFormatting(doc, target, newBlock, styleMap));
        }

        applied.Add(new
        {
            kind = change.Kind,
            oldBlock = change.OldIndex + 1,
            newBlock = change.NewIndex + 1,
            start = live.Start,
            end = live.End,
        });
    }

    private static void ApplyListChange(
        Word.Document doc,
        LiveBlockRange live,
        ListBlock oldList,
        ListBlock newList,
        StyleMap styleMap,
        List<string> warnings)
    {
        var paragraphs = doc.Range(live.Start, live.End).Paragraphs;
        for (var i = oldList.Items.Count - 1; i >= 0; i--)
        {
            var paragraph = paragraphs[i + 1].Range;
            ApplyMinimalTextChange(
                doc,
                paragraph,
                MarkdownDiffPlanner.InlineText(oldList.Items[i].Inlines),
                MarkdownDiffPlanner.InlineText(newList.Items[i].Inlines),
                styleMap);
            warnings.AddRange(MarkdownRealizer.ApplyBlockFormatting(
                doc,
                paragraph,
                new ParagraphBlock(newList.Items[i].Inlines),
                styleMap));
            paragraph.Style = doc.Styles[newList.Ordered ? WdBuiltinStyle.wdStyleListNumber : WdBuiltinStyle.wdStyleListBullet];
        }
    }

    private static void ApplyMinimalTextChange(
        Word.Document doc,
        Word.Range paragraphRange,
        string oldText,
        string newText,
        StyleMap styleMap)
    {
        if (oldText == newText)
        {
            return;
        }

        var prefix = 0;
        while (prefix < oldText.Length && prefix < newText.Length && oldText[prefix] == newText[prefix])
        {
            prefix++;
        }

        var suffix = 0;
        while (suffix < oldText.Length - prefix && suffix < newText.Length - prefix &&
               oldText[oldText.Length - suffix - 1] == newText[newText.Length - suffix - 1])
        {
            suffix++;
        }

        var content = doc.Range(paragraphRange.Start, Math.Max(paragraphRange.Start, paragraphRange.End - 1));
        var hits = WordFind.LocateEditable(doc, oldText, matchCase: true, content);
        var start = hits.Count > 0 ? hits[0].Start : content.Start;
        var end = hits.Count > 0 ? hits[0].End : content.End;
        var changed = doc.Range(start + prefix, Math.Max(start + prefix, end - suffix));
        var replacement = newText.Substring(prefix, newText.Length - prefix - suffix);
        _ = MarkdownRealizer.ReplaceRange(doc, changed, replacement, styleMap);
    }

    private static IReadOnlyList<LiveBlockRange> MapBlocks(
        Word.Document doc,
        Word.Range scope,
        MarkdownDocument model,
        List<string> warnings)
    {
        var result = new List<LiveBlockRange>();
        var cursor = scope.Start;
        foreach (var block in model.Blocks)
        {
            var pieces = block switch
            {
                ListBlock list => list.Items.Select(i => MarkdownDiffPlanner.InlineText(i.Inlines)).ToArray(),
                TableBlock table => table.Rows.SelectMany(row => row.Cells)
                    .Select(MarkdownDiffPlanner.InlineText).ToArray(),
                _ => new[] { MarkdownDiffPlanner.PlainText(block).Split('\n')[0] },
            };
            var start = -1;
            var end = -1;
            foreach (var piece in pieces)
            {
                if (string.IsNullOrEmpty(piece))
                {
                    continue;
                }

                var search = doc.Range(cursor, scope.End);
                var match = WordFind.LocateEditable(doc, piece, matchCase: true, search).FirstOrDefault();
                if (match is null)
                {
                    warnings.Add($"could not map Markdown block {result.Count + 1} to a live Word range");
                    return result;
                }

                var paragraph = doc.Range(match.Start, match.End).Paragraphs[1].Range;
                start = start < 0 ? paragraph.Start : start;
                end = paragraph.End;
                cursor = paragraph.End;
            }

            if (start < 0)
            {
                warnings.Add($"could not map empty Markdown block {result.Count + 1}");
                return result;
            }

            result.Add(new LiveBlockRange(start, end));
        }

        return result;
    }

    private static Word.Range? ResolveScope(
        Word.Document doc,
        string scope,
        string? anchor,
        StyleMap styleMap,
        out string? error)
    {
        error = null;
        if (string.Equals(scope, "doc", StringComparison.OrdinalIgnoreCase))
        {
            if (!string.IsNullOrWhiteSpace(anchor))
            {
                error = "anchor is only valid when scope is section";
                return null;
            }

            return doc.Content;
        }

        if (!string.Equals(scope, "section", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(anchor))
        {
            error = "scope must be doc or section; section requires anchor";
            return null;
        }

        var headings = new[]
        {
            styleMap.DefaultLocalName(MarkdownConstruct.Heading1),
            styleMap.DefaultLocalName(MarkdownConstruct.Heading2),
            styleMap.DefaultLocalName(MarkdownConstruct.Heading3),
        };
        var range = DocumentRanges.ResolveHeadingSection(doc, anchor, headings);
        if (range is null)
        {
            error = $"anchor '{anchor}' not found";
        }

        return range;
    }

    private static string NormalizeStyle(string style)
    {
        var value = style.Trim();
        const string prefix = "{style=\"";
        if (value.StartsWith(prefix, StringComparison.Ordinal) && value.EndsWith("\"}", StringComparison.Ordinal))
        {
            return value[prefix.Length..^2];
        }

        return value;
    }

    private static WdBuiltinStyle HeadingStyle(int level) => level switch
    {
        1 => WdBuiltinStyle.wdStyleHeading1,
        2 => WdBuiltinStyle.wdStyleHeading2,
        _ => WdBuiltinStyle.wdStyleHeading3,
    };
}
