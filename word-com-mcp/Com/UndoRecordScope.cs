using Word = NetOffice.WordApi;

namespace WordComMcp.Com;

/// <summary>
/// Groups all COM mutations of one tool call into a single custom undo record
/// (issue 0.7) so the user undoes an MCP edit with one Ctrl+Z:
/// <code>using (new UndoRecordScope(app, "MCP: Insert Text")) { ... }</code>
///
/// A stale record left by a previously crashed call is closed first. On Word 2007
/// or earlier the <c>UndoRecord</c> API is unavailable, so the scope degrades to a
/// no-op and the body still runs.
/// </summary>
public sealed class UndoRecordScope : IDisposable
{
    private const int MaxNameLength = 64;

    private Word.UndoRecord? m_record;

    public UndoRecordScope(Word.Application app, string name)
    {
        ArgumentNullException.ThrowIfNull(app);

        try
        {
            var record = app.UndoRecord;

            // Clean up a stale custom record from a prior crashed/interrupted call.
            if (record.IsRecordingCustomRecord)
            {
                TryEnd(record);
            }

            record.StartCustomRecord(Truncate(name));
            this.m_record = record;
        }
        catch (Exception)
        {
            // Word <= 2007: no UndoRecord support. Proceed without grouping.
            this.m_record = null;
        }
    }

    public void Dispose()
    {
        var record = this.m_record;
        this.m_record = null;
        if (record is not null)
        {
            TryEnd(record);
        }
    }

    private static void TryEnd(Word.UndoRecord record)
    {
        try
        {
            record.EndCustomRecord();
        }
        catch (Exception)
        {
            // Best-effort close; never surface an exception from cleanup.
        }
    }

    private static string Truncate(string name)
    {
        name ??= string.Empty;
        return name.Length <= MaxNameLength ? name : name[..MaxNameLength];
    }
}
