using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using WordComMcp.Infrastructure;
using Word = NetOffice.WordApi;

namespace WordComMcp.Com;

/// <summary>
/// Thrown when no running Word instance / requested document can be resolved.
/// Carries a stable <see cref="McpResult.Errors"/> code so tools surface a
/// structured error rather than a raw COM stack trace (issue 0.9).
/// </summary>
public sealed class WordConnectionException : Exception
{
    public WordConnectionException(string errorCode, string message)
        : base(message)
    {
        this.ErrorCode = errorCode;
    }

    public string ErrorCode { get; }
}

/// <summary>
/// COM connection layer (issue 0.5): attach to a running Word, find a document, and
/// cache the resolved application/document so the expensive ROT scan runs at most once
/// per connection. All members must be called on the STA thread
/// (<see cref="StaDispatcher"/>); the cache holds no locks by design.
/// </summary>
public sealed class WordConnection
{
    private const string WordProgId = "Word.Application";

    // Disconnect HRESULTs that invalidate the cache and force a re-resolve (Conventions #1).
    private const int RpcEDisconnected = unchecked((int)0x80010108);
    private const int RpcSServerUnavailable = unchecked((int)0x800706BA);
    private const int CoEObjNotConnected = unchecked((int)0x800401FD);

    private readonly Func<Word.Application>? m_appResolver;
    private Word.Application? m_cachedApp;

    public WordConnection()
    {
    }

    /// <summary>
    /// Create a connection that resolves a caller-owned Word application. Intended for live tests
    /// that must remain isolated from other running Word instances.
    /// </summary>
    public WordConnection(Func<Word.Application> appResolver)
    {
        ArgumentNullException.ThrowIfNull(appResolver);
        this.m_appResolver = appResolver;
    }

    /// <summary>
    /// Resolve a running Word application that has at least one document, self-healing
    /// on a disconnected cache. Mirrors the reference <c>get_word_app</c>.
    /// </summary>
    public Word.Application GetWordApp()
    {
        if (this.m_appResolver is not null)
        {
            return this.m_appResolver();
        }

        if (this.m_cachedApp is not null)
        {
            if (this.IsCacheAlive())
            {
                return this.m_cachedApp;
            }

            this.InvalidateCache();
        }

        var resolved = ResolveWordApp();
        this.m_cachedApp = resolved;
        return resolved;
    }

    /// <summary>
    /// Find a document in <paramref name="app"/> (issue 0.5). <c>null</c>/empty
    /// <paramref name="filename"/> returns the active document; otherwise match by
    /// basename (<c>Name</c>) and — only for absolute inputs — normalized full path
    /// (<c>FullName</c>), both compared NFC + case-insensitively.
    /// </summary>
    public Word.Document FindDocument(Word.Application app, string? filename)
    {
        ArgumentNullException.ThrowIfNull(app);

        var documents = app.Documents;
        if (documents.Count == 0)
        {
            throw new WordConnectionException(McpResult.Errors.DocumentNotFound, "No documents are open in Word.");
        }

        if (string.IsNullOrWhiteSpace(filename))
        {
            return app.ActiveDocument;
        }

        var targetBasename = Normalize(Path.GetFileName(filename));
        var targetFullPath = Path.IsPathRooted(filename) ? Normalize(Path.GetFullPath(filename)) : null;

        var openNames = new List<string>();
        for (var i = 1; i <= documents.Count; i++)
        {
            var doc = documents[i];
            openNames.Add(doc.Name);

            if (Normalize(doc.Name) == targetBasename)
            {
                return doc;
            }

            if (targetFullPath is not null && Normalize(Path.GetFullPath(doc.FullName)) == targetFullPath)
            {
                return doc;
            }
        }

        throw new WordConnectionException(
            McpResult.Errors.DocumentNotFound,
            $"Document '{filename}' is not open in Word. Open documents: {string.Join(", ", openNames)}");
    }

    /// <summary>Drop the cached application (e.g. after a detected disconnect or on shutdown).</summary>
    public void InvalidateCache()
    {
        if (this.m_appResolver is not null)
        {
            return;
        }

        var cached = this.m_cachedApp;
        this.m_cachedApp = null;
        cached?.Dispose();
    }

    private bool IsCacheAlive()
    {
        try
        {
            // Cheap round-trip; throws a disconnect HRESULT when the proxy is dead.
            _ = this.m_cachedApp!.Documents.Count;
            return true;
        }
        catch (COMException ex) when (IsDisconnected(ex))
        {
            return false;
        }
    }

    private static Word.Application ResolveWordApp()
    {
        object? active = null;
        try
        {
            active = RunningObjectTable.GetActiveObject(WordProgId);
            if (active is not null && CountDocuments(active) > 0)
            {
                return Wrap(active);
            }

            // Active proxy is empty (common under O365/OneDrive) — scan the ROT.
            var withDocs = FindWordWithDocs();
            if (withDocs is not null)
            {
                return Wrap(withDocs);
            }

            // No instance has documents; hand back the empty active one if we have it.
            if (active is not null)
            {
                return Wrap(active);
            }
        }
        catch (WordConnectionException)
        {
            throw;
        }
        catch (Exception)
        {
            var rescued = FindWordWithDocs();
            if (rescued is not null)
            {
                return Wrap(rescued);
            }
        }

        throw new WordConnectionException(
            McpResult.Errors.WordNotRunning,
            "Microsoft Word is not running. Please open Word first.");
    }

    /// <summary>
    /// Two-pass ROT scan (port of the reference <c>_find_word_with_docs</c>): pass 1
    /// returns any <c>Word.Application</c> entry that has documents; pass 2 rescues the
    /// O365 case by binding a <c>.docx</c>/<c>.doc</c> moniker → Document → Application.
    /// </summary>
    private static object? FindWordWithDocs()
    {
        var fileMonikers = new List<object>();

        foreach (var (displayName, instance) in RunningObjectTable.Enumerate())
        {
            try
            {
                if (LooksLikeWordApplication(instance) && CountDocuments(instance) > 0)
                {
                    return instance;
                }

                if (IsWordFileName(displayName))
                {
                    fileMonikers.Add(instance);
                }
            }
            catch (Exception)
            {
                if (IsWordFileName(displayName))
                {
                    fileMonikers.Add(instance);
                }
            }
        }

        foreach (var candidate in fileMonikers)
        {
            try
            {
                dynamic doc = candidate;
                object app = doc.Application;
                if (CountDocuments(app) > 0)
                {
                    return app;
                }
            }
            catch (Exception)
            {
                // Stale/foreign moniker — skip.
            }
        }

        return null;
    }

    private static bool LooksLikeWordApplication(object instance)
    {
        try
        {
            dynamic candidate = instance;
            _ = candidate.Documents;
            _ = candidate.ActiveDocument;
            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }

    private static int CountDocuments(object appProxy)
    {
        dynamic app = appProxy;
        return (int)app.Documents.Count;
    }

    private static bool IsWordFileName(string displayName)
    {
        if (string.IsNullOrEmpty(displayName))
        {
            return false;
        }

        return displayName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
            || displayName.EndsWith(".doc", StringComparison.OrdinalIgnoreCase);
    }

    private static Word.Application Wrap(object comProxy) => new(null, comProxy);

    private static string Normalize(string value) =>
        value.Normalize(NormalizationForm.FormC).ToLower(CultureInfo.InvariantCulture);

    private static bool IsDisconnected(COMException ex) =>
        ex.HResult is RpcEDisconnected or RpcSServerUnavailable or CoEObjNotConnected;
}
