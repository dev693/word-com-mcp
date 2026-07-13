using WordComMcp.Infrastructure;
using Word = NetOffice.WordApi;
using WdFindWrap = NetOffice.WordApi.Enums.WdFindWrap;
using WdRevisionType = NetOffice.WordApi.Enums.WdRevisionType;

namespace WordComMcp.Com;

/// <summary>
/// The single <c>Word.Find</c> wrapper (issue 1.6 + Conventions). Locates literal text against
/// the <b>final</b> view and annotates each hit's revision state. Runs <c>Find.Execute</c>
/// in-process (perf rule), never <c>wdReplaceAll</c> under tracking, and enforces Word's
/// 255-character <c>FindText</c> limit (guarded for search, chunked for edit anchors). STA-only.
/// </summary>
public static class WordFind
{
    /// <summary>Word's hard limit on the length of a single <c>Find.Execute</c> pattern.</summary>
    public const int MaxFindLength = 255;

    public const string StateClean = "clean";
    public const string StateInsideInsertion = "inside-insertion";
    public const string StateInsideDeletion = "inside-deletion";

    /// <summary>A search hit: character offsets, a short context snippet, and its revision state.</summary>
    public sealed record FindHit(int Start, int End, string Context, string RevisionState);

    /// <summary>A resolved editable match range (final-view; deletions excluded).</summary>
    public sealed record MatchRange(int Start, int End, string RevisionState);

    /// <summary>
    /// Search for <paramref name="query"/> and return up to <paramref name="maxResults"/> deduped
    /// hits, each annotated with its revision state (for <c>find_text</c>). Throws
    /// <see cref="WordConnectionException"/> with <see cref="McpResult.Errors.FindLimitExceeded"/>
    /// when the query exceeds 255 characters.
    /// </summary>
    public static IReadOnlyList<FindHit> Search(
        Word.Document doc,
        string query,
        bool matchCase,
        int maxResults,
        Word.Range? scope = null)
    {
        ArgumentNullException.ThrowIfNull(doc);
        if (string.IsNullOrEmpty(query))
        {
            return Array.Empty<FindHit>();
        }

        if (query.Length > MaxFindLength)
        {
            throw new WordConnectionException(McpResult.Errors.FindLimitExceeded, McpResult.Errors.FindLimitExceeded);
        }

        var hits = new List<FindHit>();
        var seen = new HashSet<(int, int)>();
        foreach (var (start, end) in EnumerateMatches(doc, scope, query, matchCase))
        {
            if (!seen.Add((start, end)))
            {
                continue;
            }

            var range = doc.Range(start, end);
            hits.Add(new FindHit(start, end, Context(doc, start, end), ClassifyRevision(range)));
            if (hits.Count >= maxResults)
            {
                break;
            }
        }

        return hits;
    }

    /// <summary>
    /// Locate final-view matches of <paramref name="find"/> (deletions excluded) as editable
    /// ranges, ordered by position, for <c>replace_text</c>/<c>delete_text</c>. Patterns longer
    /// than 255 chars are chunked and stitched into a contiguous range.
    /// </summary>
    public static IReadOnlyList<MatchRange> LocateEditable(
        Word.Document doc,
        string find,
        bool matchCase,
        Word.Range? scope = null)
    {
        ArgumentNullException.ThrowIfNull(doc);
        if (string.IsNullOrEmpty(find))
        {
            return Array.Empty<MatchRange>();
        }

        var matches = find.Length <= MaxFindLength
            ? LocateShort(doc, find, matchCase, scope)
            : LocateChunked(doc, find, matchCase, scope);

        // Only clean / inserted text is part of the final view; deleted text is not editable here.
        return matches.Where(m => m.RevisionState != StateInsideDeletion).ToList();
    }

    private static List<MatchRange> LocateShort(Word.Document doc, string find, bool matchCase, Word.Range? scope)
    {
        var result = new List<MatchRange>();
        foreach (var (start, end) in EnumerateMatches(doc, scope, find, matchCase))
        {
            result.Add(new MatchRange(start, end, ClassifyRevision(doc.Range(start, end))));
        }

        return result;
    }

    private static List<MatchRange> LocateChunked(Word.Document doc, string find, bool matchCase, Word.Range? scope)
    {
        var head = find[..MaxFindLength];
        var tail = find[MaxFindLength..];
        var scopeEnd = (scope ?? doc.Content).End;

        var result = new List<MatchRange>();
        foreach (var (start, end) in EnumerateMatches(doc, scope, head, matchCase))
        {
            var tailEnd = Math.Min(end + tail.Length, scopeEnd);
            var tailRange = doc.Range(end, tailEnd);
            var tailText = tailRange.Text ?? string.Empty;
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            if (tailText.Equals(tail, comparison))
            {
                result.Add(new MatchRange(start, tailEnd, ClassifyRevision(doc.Range(start, tailEnd))));
            }
        }

        return result;
    }

    private static IEnumerable<(int Start, int End)> EnumerateMatches(
        Word.Document doc,
        Word.Range? scope,
        string query,
        bool matchCase)
    {
        var searchRange = (scope ?? doc.Content).Duplicate;
        var boundStart = searchRange.Start;
        var boundEnd = searchRange.End;

        // Configure the search via Find's properties and drive it with the no-arg Execute() —
        // the reliable NetOffice pattern (avoids the many-optional-parameter Execute overload).
        var find = searchRange.Find;
        find.ClearFormatting();
        find.Text = query;
        find.Forward = true;
        find.Wrap = WdFindWrap.wdFindStop;
        find.MatchCase = matchCase;
        find.MatchWholeWord = false;
        find.MatchWildcards = false;
        find.MatchSoundsLike = false;
        find.MatchAllWordForms = false;

        var matches = new List<(int, int)>();
        while (find.Execute())
        {
            var start = searchRange.Start;
            var end = searchRange.End;
            if (end <= start || start < boundStart || end > boundEnd)
            {
                break;
            }

            matches.Add((start, end));
            if (end >= boundEnd)
            {
                break;
            }

            // Continue after this match, staying inside the scope.
            searchRange.Start = end;
            searchRange.End = boundEnd;
        }

        return matches;
    }

    private static string ClassifyRevision(Word.Range range)
    {
        var revisions = range.Revisions;
        if (revisions.Count == 0)
        {
            return StateClean;
        }

        var sawInsertion = false;
        for (var i = 1; i <= revisions.Count; i++)
        {
            var type = revisions[i].Type;
            if (type is WdRevisionType.wdRevisionDelete or WdRevisionType.wdRevisionMovedFrom)
            {
                return StateInsideDeletion;
            }

            if (type is WdRevisionType.wdRevisionInsert or WdRevisionType.wdRevisionMovedTo)
            {
                sawInsertion = true;
            }
        }

        return sawInsertion ? StateInsideInsertion : StateClean;
    }

    private static string Context(Word.Document doc, int start, int end)
    {
        const int pad = 30;
        var content = doc.Content;
        var from = Math.Max(content.Start, start - pad);
        var to = Math.Min(content.End, end + pad);
        var text = doc.Range(from, to).Text ?? string.Empty;
        return text.Replace('\r', ' ').Replace('\n', ' ').Replace('\a', ' ').Trim();
    }
}
