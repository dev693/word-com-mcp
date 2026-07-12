namespace WordComMcp.Infrastructure;

/// <summary>
/// Configuration read from environment variables (issue 0.8).
/// <list type="bullet">
///   <item><c>MCP_AUTHOR</c> — author name for tracked edits/comments (default <c>"Author"</c>).</item>
///   <item><c>MCP_AUTHOR_INITIALS</c> — author initials (default empty).</item>
///   <item><c>MCP_STYLE_MAP</c> — optional path overriding the default style map (0.12).</item>
/// </list>
/// </summary>
public sealed class ServerConfig
{
    public const string AuthorEnv = "MCP_AUTHOR";
    public const string AuthorInitialsEnv = "MCP_AUTHOR_INITIALS";
    public const string StyleMapEnv = "MCP_STYLE_MAP";

    public const string DefaultAuthor = "Author";

    public ServerConfig(string author, string authorInitials, string? styleMapPath)
    {
        this.Author = author;
        this.AuthorInitials = authorInitials;
        this.StyleMapPath = styleMapPath;
    }

    /// <summary>Author name applied to tracked changes and comments.</summary>
    public string Author { get; }

    /// <summary>Author initials applied to comments.</summary>
    public string AuthorInitials { get; }

    /// <summary>Optional path to a JSON style-map override (0.12); null when unset.</summary>
    public string? StyleMapPath { get; }

    /// <summary>Read the configuration from the current process environment, applying defaults.</summary>
    public static ServerConfig FromEnvironment()
    {
        var author = Environment.GetEnvironmentVariable(AuthorEnv);
        var initials = Environment.GetEnvironmentVariable(AuthorInitialsEnv);
        var styleMap = Environment.GetEnvironmentVariable(StyleMapEnv);

        return new ServerConfig(
            author: string.IsNullOrWhiteSpace(author) ? DefaultAuthor : author,
            authorInitials: initials ?? string.Empty,
            styleMapPath: string.IsNullOrWhiteSpace(styleMap) ? null : styleMap);
    }
}
