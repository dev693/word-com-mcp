using WordComMcp.Infrastructure;
using WordComMcp.Markdown;
using Word = NetOffice.WordApi;
using WdOutlineLevel = NetOffice.WordApi.Enums.WdOutlineLevel;

namespace WordComMcp.Com;

/// <summary>Shared range, heading, author, and comment helpers for Tier-2 tools.</summary>
public static class TierTwoCom
{
    public sealed record HeadingInfo(int Level, string Text, int Start, int End);

    public static Word.Range? ResolveParagraph(
        Word.Document doc,
        string? anchorText,
        int? paragraphIndex,
        out string? error)
    {
        error = null;
        if ((!string.IsNullOrWhiteSpace(anchorText) ? 1 : 0) + (paragraphIndex.HasValue ? 1 : 0) != 1)
        {
            error = "provide exactly one of anchorText or paragraphIndex";
            return null;
        }

        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value < 1 || paragraphIndex.Value > doc.Paragraphs.Count)
            {
                error = $"paragraphIndex {paragraphIndex.Value} is out of range";
                return null;
            }

            return doc.Paragraphs[paragraphIndex.Value].Range;
        }

        var matches = WordFind.LocateEditable(doc, anchorText!, matchCase: false);
        if (matches.Count == 0)
        {
            error = $"anchorText '{anchorText}' not found";
            return null;
        }

        return doc.Range(matches[0].Start, matches[0].End).Paragraphs[1].Range;
    }

    public static IReadOnlyList<HeadingInfo> ReadHeadings(Word.Document doc)
    {
        var result = new List<HeadingInfo>();
        var paragraphs = doc.Paragraphs;
        for (var i = 1; i <= paragraphs.Count; i++)
        {
            var paragraph = paragraphs[i];
            var level = (int)paragraph.OutlineLevel;
            if (level is < 1 or > 9 || paragraph.OutlineLevel == WdOutlineLevel.wdOutlineLevelBodyText)
            {
                continue;
            }

            var range = paragraph.Range;
            result.Add(new HeadingInfo(
                level,
                OoxmlFinalText.GetFinalText(range).TrimEnd('\r', '\n', '\a'),
                range.Start,
                range.End));
        }

        return result;
    }

    public static T WithAuthor<T>(Word.Application app, ServerConfig config, Func<T> action)
    {
        var previousName = app.UserName;
        var previousInitials = app.UserInitials;
        app.UserName = config.Author;
        if (!string.IsNullOrEmpty(config.AuthorInitials))
        {
            app.UserInitials = config.AuthorInitials;
        }

        try
        {
            return action();
        }
        finally
        {
            app.UserName = previousName;
            app.UserInitials = previousInitials;
        }
    }

    public static Word.Comment? ResolveComment(Word.Document doc, int commentId, out string? error)
    {
        error = null;
        if (commentId < 1 || commentId > doc.Comments.Count)
        {
            error = $"commentId {commentId} is out of range";
            return null;
        }

        return doc.Comments[commentId];
    }

    public static string CleanRangeText(Word.Range? range, int maxLength = 0)
    {
        var text = range?.Text ?? string.Empty;
        text = text.TrimEnd('\r', '\n', '\a');
        return maxLength > 0 && text.Length > maxLength ? text[..maxLength] : text;
    }
}
