using System.Runtime.InteropServices;

namespace WordComMcp.Com;

/// <summary>
/// Disposable wrapper that releases a raw COM proxy via
/// <see cref="Marshal.ReleaseComObject(object)"/> on dispose (issue 0.6).
///
/// <para><b>Release policy (documented decision for tool authors):</b></para>
/// <list type="bullet">
///   <item>
///     NetOffice objects (<c>NetOffice.WordApi.*</c>) track their own proxy tree and
///     are released by calling <c>Dispose()</c> / <c>DisposeChildInstances()</c> on
///     them — do <b>not</b> call <see cref="Marshal.ReleaseComObject(object)"/> on a
///     NetOffice object. Prefer disposing the NetOffice <c>Application</c> to release
///     the whole tree.
///   </item>
///   <item>
///     Use <see cref="ComRef{T}"/> only for <i>raw</i> COM proxies obtained outside
///     NetOffice's tracking — e.g. the objects returned by the Running Object Table
///     scan in <see cref="RunningObjectTable"/>, or intermediate
///     <c>Range</c>/<c>Selection</c>/<c>Table</c> proxies you fetch by casting.
///   </item>
///   <item>
///     Fetch-once: avoid multi-dot chains (each dot is a marshalled round-trip that
///     also leaks an intermediate proxy). Capture the intermediate in a
///     <see cref="ComRef{T}"/> and reuse it.
///   </item>
/// </list>
/// </summary>
/// <typeparam name="T">The proxy's static type (usually <see cref="object"/> for late binding).</typeparam>
public sealed class ComRef<T> : IDisposable
    where T : class
{
    private T? m_value;

    public ComRef(T value)
    {
        this.m_value = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>The wrapped proxy. Throws once disposed.</summary>
    public T Value =>
        this.m_value ?? throw new ObjectDisposedException(nameof(ComRef<T>));

    /// <summary>Detach the proxy so dispose will not release it (ownership transferred to the caller).</summary>
    public T Detach()
    {
        var value = this.Value;
        this.m_value = null;
        return value;
    }

    public void Dispose()
    {
        var value = this.m_value;
        this.m_value = null;
        if (value is not null && Marshal.IsComObject(value))
        {
            Marshal.ReleaseComObject(value);
        }
    }
}

/// <summary>Factory helpers so <c>ComRef.Of(proxy)</c> can infer the type argument.</summary>
public static class ComRef
{
    public static ComRef<T> Of<T>(T value)
        where T : class => new(value);
}
