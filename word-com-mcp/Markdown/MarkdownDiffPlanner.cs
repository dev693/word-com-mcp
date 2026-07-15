namespace WordComMcp.Markdown;

/// <summary>One block-level operation produced by <see cref="MarkdownDiffPlanner"/>.</summary>
public sealed record MarkdownDiffChange(string Kind, int OldIndex, int NewIndex);

/// <summary>A preflighted Markdown edit plan. Warnings make the plan non-executable.</summary>
public sealed record MarkdownDiffPlan(
    IReadOnlyList<MarkdownDiffChange> Changes,
    IReadOnlyList<string> Warnings)
{
    public bool CanApply => this.Warnings.Count == 0;
}

/// <summary>
/// Conservative, Word-free block differ for <c>word_live_apply_markdown</c>. Exact blocks are
/// aligned with an LCS; one-for-one replacements are editable, while ambiguous multi-block
/// replacements and structural edits to existing lists/tables are rejected during preflight.
/// </summary>
public static class MarkdownDiffPlanner
{
    public static MarkdownDiffPlan Plan(MarkdownDocument current, MarkdownDocument requested, StyleMap styleMap)
    {
        ArgumentNullException.ThrowIfNull(current);
        ArgumentNullException.ThrowIfNull(requested);
        ArgumentNullException.ThrowIfNull(styleMap);

        var serializer = new MarkdownSerializer(styleMap);
        var oldBlocks = current.Blocks.Select(b => Serialize(serializer, b)).ToArray();
        var newBlocks = requested.Blocks.Select(b => Serialize(serializer, b)).ToArray();
        var matches = LongestCommonSubsequence(oldBlocks, newBlocks);
        var changes = new List<MarkdownDiffChange>();
        var warnings = new List<string>();

        var oldCursor = 0;
        var newCursor = 0;
        foreach (var (oldMatch, newMatch) in matches.Append((oldBlocks.Length, newBlocks.Length)))
        {
            AddSegment(current, requested, oldCursor, oldMatch, newCursor, newMatch, changes, warnings);
            oldCursor = oldMatch + 1;
            newCursor = newMatch + 1;
        }

        return new MarkdownDiffPlan(warnings.Count == 0 ? changes : Array.Empty<MarkdownDiffChange>(), warnings);
    }

    public static string PlainText(MarkdownBlock block) => block switch
    {
        HeadingBlock h => InlineText(h.Inlines),
        ParagraphBlock p => InlineText(p.Inlines),
        BlockQuoteBlock q => InlineText(q.Inlines),
        CodeBlock c => c.Code,
        ListBlock l => string.Join("\r", l.Items.Select(i => InlineText(i.Inlines))),
        TableBlock t => string.Join("\r", t.Rows.SelectMany(r => r.Cells).Select(InlineText)),
        _ => string.Empty,
    };

    public static string InlineText(IReadOnlyList<MarkdownInline> inlines) =>
        string.Concat(inlines.Select(i => i switch
        {
            TextInline t => t.Text,
            EmphasisInline e => e.Text,
            CodeSpanInline c => c.Code,
            LinkInline l => l.Text,
            StyledSpanInline s => s.Text,
            _ => string.Empty,
        }));

    private static void AddSegment(
        MarkdownDocument current,
        MarkdownDocument requested,
        int oldStart,
        int oldEnd,
        int newStart,
        int newEnd,
        List<MarkdownDiffChange> changes,
        List<string> warnings)
    {
        var oldCount = oldEnd - oldStart;
        var newCount = newEnd - newStart;
        if (oldCount == 0 && newCount == 0)
        {
            return;
        }

        if (oldCount == 1 && newCount == 1)
        {
            var oldBlock = current.Blocks[oldStart];
            var newBlock = requested.Blocks[newStart];
            if (oldBlock is TableBlock || newBlock is TableBlock)
            {
                warnings.Add($"table change at block {oldStart + 1} is too structural to apply safely");
                return;
            }

            if (oldBlock is ListBlock oldList && newBlock is ListBlock newList)
            {
                if (oldList.Ordered != newList.Ordered || oldList.Items.Count != newList.Items.Count)
                {
                    warnings.Add($"existing list at block {oldStart + 1} would be renumbered or re-leveled");
                    return;
                }

                var oldItems = oldList.Items.Select(i => InlineText(i.Inlines)).ToArray();
                var newItems = newList.Items.Select(i => InlineText(i.Inlines)).ToArray();
                if (SameItemsDifferentOrder(oldItems, newItems))
                {
                    warnings.Add($"existing list at block {oldStart + 1} would be reordered");
                    return;
                }
            }
            else if (oldBlock is ListBlock || newBlock is ListBlock)
            {
                warnings.Add($"existing list at block {oldStart + 1} would be re-leveled");
                return;
            }

            var kind = PlainText(oldBlock) == PlainText(newBlock) ? "style" : "replace";
            changes.Add(new MarkdownDiffChange(kind, oldStart, newStart));
            return;
        }

        if (oldCount > 0 && newCount > 0)
        {
            warnings.Add(
                $"blocks {oldStart + 1}-{oldEnd} form an ambiguous multi-block replacement; no changes were applied");
            return;
        }

        if (oldCount > 0)
        {
            for (var i = oldStart; i < oldEnd; i++)
            {
                if (current.Blocks[i] is ListBlock or TableBlock)
                {
                    warnings.Add($"deleting structural block {i + 1} is not safe under tracked changes");
                }
                else
                {
                    changes.Add(new MarkdownDiffChange("delete", i, -1));
                }
            }

            return;
        }

        for (var i = newStart; i < newEnd; i++)
        {
            if (requested.Blocks[i] is TableBlock)
            {
                warnings.Add($"table insertion at block {i + 1} is not supported");
            }
            else
            {
                changes.Add(new MarkdownDiffChange("insert", oldStart, i));
            }
        }
    }

    private static bool SameItemsDifferentOrder(string[] oldItems, string[] newItems) =>
        !oldItems.SequenceEqual(newItems) &&
        oldItems.OrderBy(x => x, StringComparer.Ordinal).SequenceEqual(
            newItems.OrderBy(x => x, StringComparer.Ordinal), StringComparer.Ordinal);

    private static string Serialize(MarkdownSerializer serializer, MarkdownBlock block) =>
        serializer.Serialize(new MarkdownDocument([block]));

    private static IReadOnlyList<(int Old, int New)> LongestCommonSubsequence(string[] oldBlocks, string[] newBlocks)
    {
        var lengths = new int[oldBlocks.Length + 1, newBlocks.Length + 1];
        for (var i = oldBlocks.Length - 1; i >= 0; i--)
        {
            for (var j = newBlocks.Length - 1; j >= 0; j--)
            {
                lengths[i, j] = oldBlocks[i] == newBlocks[j]
                    ? lengths[i + 1, j + 1] + 1
                    : Math.Max(lengths[i + 1, j], lengths[i, j + 1]);
            }
        }

        var result = new List<(int, int)>();
        var oldIndex = 0;
        var newIndex = 0;
        while (oldIndex < oldBlocks.Length && newIndex < newBlocks.Length)
        {
            if (oldBlocks[oldIndex] == newBlocks[newIndex])
            {
                result.Add((oldIndex++, newIndex++));
            }
            else if (lengths[oldIndex + 1, newIndex] >= lengths[oldIndex, newIndex + 1])
            {
                oldIndex++;
            }
            else
            {
                newIndex++;
            }
        }

        return result;
    }
}
