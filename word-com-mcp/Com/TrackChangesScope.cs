using WordComMcp.Infrastructure;
using Word = NetOffice.WordApi;

namespace WordComMcp.Com;

/// <summary>
/// Saves and restores <c>Document.TrackRevisions</c> and <c>Application.UserName</c>/
/// <c>UserInitials</c> around an action (Conventions Q2 "WithTrackChanges"). When a
/// tracked edit is requested, tracking is turned on and Word's identity is swapped to
/// the configured author so the revision carries the right name; the original state is
/// always restored in <c>finally</c>, even on exception.
/// </summary>
public static class TrackChangesScope
{
    /// <summary>
    /// Run <paramref name="action"/> with an optional tracked-changes override.
    /// <paramref name="trackChanges"/>: <c>true</c>/<c>false</c> forces the state,
    /// <c>null</c> leaves the document's current setting untouched.
    /// </summary>
    public static T With<T>(Word.Application app, Word.Document doc, ServerConfig config, bool? trackChanges, Func<T> action)
    {
        ArgumentNullException.ThrowIfNull(app);
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(config);
        ArgumentNullException.ThrowIfNull(action);

        var prevTracking = doc.TrackRevisions;
        var prevUserName = app.UserName;
        var prevInitials = app.UserInitials;

        if (trackChanges is bool desired)
        {
            doc.TrackRevisions = desired;
            if (desired)
            {
                app.UserName = config.Author;
                if (!string.IsNullOrEmpty(config.AuthorInitials))
                {
                    app.UserInitials = config.AuthorInitials;
                }
            }
        }

        try
        {
            return action();
        }
        finally
        {
            doc.TrackRevisions = prevTracking;
            app.UserName = prevUserName;
            app.UserInitials = prevInitials;
        }
    }

    /// <summary>Void overload of <see cref="With{T}"/>.</summary>
    public static void With(Word.Application app, Word.Document doc, ServerConfig config, bool? trackChanges, Action action)
    {
        ArgumentNullException.ThrowIfNull(action);
        With(app, doc, config, trackChanges, () =>
        {
            action();
            return true;
        });
    }
}
