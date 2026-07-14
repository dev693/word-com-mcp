using System.ComponentModel;
using ModelContextProtocol.Server;
using WordComMcp.Com;
using WordComMcp.Infrastructure;
using Word = NetOffice.WordApi;
using WdSaveFormat = NetOffice.WordApi.Enums.WdSaveFormat;

namespace WordComMcp.Tools;

/// <summary>
/// Document-level Tier-1 tools: toggle tracked changes, save/save-as, and metadata
/// (issues 1.2, 1.8, 1.9). Each runs through <see cref="ToolExecution"/> so it is a
/// single STA hop returning a <see cref="McpResult"/> envelope.
/// </summary>
[McpServerToolType]
public static class DocumentTools
{
    [McpServerTool(Name = "word_live_toggle_track_changes")]
    [Description(
        "Turn Word's tracked-changes (revisions) mode on or off for a document. " +
        "This is a document-global switch and persists; it is not restored afterwards. " +
        "Returns the previous and current state.")]
    public static Task<string> ToggleTrackChanges(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("True to enable tracked changes, false to disable.")] bool enabled,
        [Description("Target document basename or full path. Null/empty uses the active document.")]
        string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (_, doc) =>
        {
            var previous = doc.TrackRevisions;
            doc.TrackRevisions = enabled;
            return McpResult.Ok(new
            {
                document = doc.Name,
                previous,
                current = enabled,
            });
        });

    [McpServerTool(Name = "word_live_save")]
    [Description("Save the document in place. Returns a read-only/protected error if it cannot be saved.")]
    public static Task<string> Save(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Target document basename or full path. Null/empty uses the active document.")]
        string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (_, doc) =>
        {
            if (doc.ReadOnly)
            {
                return McpResult.Err(McpResult.Errors.ReadOnlyOrProtected);
            }

            doc.Save();
            return McpResult.Ok(new { document = doc.Name, saved = true, path = SafeFullName(doc) });
        });

    [McpServerTool(Name = "word_live_save_as")]
    [Description(
        "Save the document to a new .docx path (Word XML format). Note this changes the active " +
        "document's path. Returns the written path.")]
    public static Task<string> SaveAs(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Destination .docx path.")] string path,
        [Description("Target document basename or full path. Null/empty uses the active document.")]
        string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (_, doc) =>
        {
            doc.SaveAs2(path, WdSaveFormat.wdFormatXMLDocument);
            return McpResult.Ok(new { document = doc.Name, saved = true, path });
        });

    [McpServerTool(Name = "word_live_get_info")]
    [Description(
        "Report document metadata: name, full path, saved/track-revisions/protection state, and " +
        "paragraph / revision / comment counts.")]
    public static Task<string> GetInfo(
        StaDispatcher dispatcher,
        WordConnection connection,
        [Description("Target document basename or full path. Null/empty uses the active document.")]
        string? filename = null) =>
        ToolExecution.RunAsync(dispatcher, connection, filename, (_, doc) =>
            McpResult.Ok(new
            {
                document = doc.Name,
                fullName = SafeFullName(doc),
                saved = doc.Saved,
                trackRevisions = doc.TrackRevisions,
                protectionType = doc.ProtectionType.ToString(),
                paragraphs = doc.Paragraphs.Count,
                revisions = doc.Revisions.Count,
                comments = doc.Comments.Count,
            }));

    private static string SafeFullName(Word.Document doc)
    {
        try
        {
            return doc.FullName;
        }
        catch (Exception)
        {
            return doc.Name;
        }
    }
}
