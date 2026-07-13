using System.ComponentModel;
using ModelContextProtocol.Server;
using WordComMcp.Com;
using WordComMcp.Infrastructure;

namespace WordComMcp.Tools;

/// <summary>
/// Formatvorlagen discovery (issue 1.1). Lets the LLM see which paragraph/character styles
/// it may target in <c>{style=…}</c> Markdown annotations before editing.
/// </summary>
[McpServerToolType]
public static class StyleTools
{
    [McpServerTool(Name = "word_live_list_styles")]
    [Description(
        "List the document's styles (Formatvorlagen): name, localized name, type " +
        "(paragraph|character|linked|table|list), whether it is a Word built-in, and whether " +
        "it is in use. Use the localized names in {style=\"…\"} annotations for insert/replace.")]
    public static Task<string> ListStyles(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Target document basename or full path. Null/empty uses the active document.")]
        string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (_, doc) =>
        {
            var styles = DocumentStyles.Enumerate(doc);
            return McpResult.Ok(new
            {
                document = doc.Name,
                count = styles.Count,
                styles = styles.Select(s => new
                {
                    name = s.Name,
                    nameLocal = s.NameLocal,
                    type = s.Type,
                    builtin = s.BuiltIn,
                    inUse = s.InUse,
                }),
            });
        });
}
