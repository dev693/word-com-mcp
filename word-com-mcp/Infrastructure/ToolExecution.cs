using WordComMcp.Com;
using Word = NetOffice.WordApi;

namespace WordComMcp.Infrastructure;

/// <summary>
/// Standard shell every COM tool runs through so the 10 Tier-1 tools stay uniform
/// (Conventions Q3/Q4 + #6 "one coarse STA hop per tool"). It performs the whole
/// tool body in a single <see cref="StaDispatcher.RunOnStaAsync{T}"/> lambda,
/// resolves the Word application + target document once, and maps any
/// <see cref="WordConnectionException"/> to its structured <see cref="McpResult.Errors"/>
/// code (unknown failures degrade to a structured message, never a raw stack trace).
/// </summary>
public static class ToolExecution
{
    /// <summary>Run an application-only tool such as <c>list_open</c> on the STA thread.</summary>
    public static async Task<string> RunApplicationAsync(
        StaDispatcher dispatcher,
        WordConnection connection,
        Func<Word.Application, string> body)
    {
        ArgumentNullException.ThrowIfNull(dispatcher);
        ArgumentNullException.ThrowIfNull(connection);
        ArgumentNullException.ThrowIfNull(body);

        try
        {
            return await dispatcher.RunOnStaAsync(() => body(connection.GetWordApp()));
        }
        catch (WordConnectionException ex)
        {
            return McpResult.Err(ex.ErrorCode);
        }
        catch (Exception ex)
        {
            return McpResult.Err(ex.Message);
        }
    }

    /// <summary>
    /// Resolve the app + document and run <paramref name="body"/> on the STA thread.
    /// <paramref name="body"/> returns a serialized <see cref="McpResult"/> envelope.
    /// </summary>
    public static async Task<string> RunAsync(
        StaDispatcher dispatcher,
        WordConnection connection,
        string? filename,
        Func<Word.Application, Word.Document, string> body)
    {
        ArgumentNullException.ThrowIfNull(dispatcher);
        ArgumentNullException.ThrowIfNull(connection);
        ArgumentNullException.ThrowIfNull(body);

        try
        {
            return await dispatcher.RunOnStaAsync(() =>
            {
                var app = connection.GetWordApp();
                var doc = connection.FindDocument(app, filename);
                return body(app, doc);
            });
        }
        catch (WordConnectionException ex)
        {
            return McpResult.Err(ex.ErrorCode);
        }
        catch (Exception ex)
        {
            // Structured message, not a stack trace (Conventions Q3).
            return McpResult.Err(ex.Message);
        }
    }
}
