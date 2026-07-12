using Word = NetOffice.WordApi;

namespace WordComMcp.Com;

/// <summary>
/// Suppresses Word's reactive recompute for the duration of a mutating tool call
/// (Conventions "COM performance" #2). Caches the current settings, disables screen
/// updating / background pagination / as-you-type proofing, and restores them on
/// dispose — even when the body throws.
/// <code>using (new FastEditScope(app)) { ... mutate ... }</code>
/// </summary>
public sealed class FastEditScope : IDisposable
{
    private readonly Word.Application m_app;
    private readonly bool m_prevScreenUpdating;
    private readonly bool m_prevPagination;
    private readonly bool m_prevCheckSpelling;
    private readonly bool m_prevCheckGrammar;
    private bool m_disposed;

    public FastEditScope(Word.Application app)
    {
        this.m_app = app ?? throw new ArgumentNullException(nameof(app));

        var options = app.Options;
        this.m_prevScreenUpdating = app.ScreenUpdating;
        this.m_prevPagination = options.Pagination;
        this.m_prevCheckSpelling = options.CheckSpellingAsYouType;
        this.m_prevCheckGrammar = options.CheckGrammarAsYouType;

        app.ScreenUpdating = false;
        options.Pagination = false;
        options.CheckSpellingAsYouType = false;
        options.CheckGrammarAsYouType = false;
    }

    public void Dispose()
    {
        if (this.m_disposed)
        {
            return;
        }

        this.m_disposed = true;

        try
        {
            var options = this.m_app.Options;
            options.Pagination = this.m_prevPagination;
            options.CheckSpellingAsYouType = this.m_prevCheckSpelling;
            options.CheckGrammarAsYouType = this.m_prevCheckGrammar;
            this.m_app.ScreenUpdating = this.m_prevScreenUpdating;
        }
        catch (Exception)
        {
            // Restoring performance toggles must never mask the tool's own outcome.
        }
    }
}
