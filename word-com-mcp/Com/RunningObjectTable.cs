using System.Runtime.InteropServices;
using COMTypes = System.Runtime.InteropServices.ComTypes;

namespace WordComMcp.Com;

/// <summary>
/// Raw Running Object Table (ROT) interop used by <see cref="WordConnection"/> (issue 0.5).
/// NetOffice offers no ROT helper, and .NET (Core) dropped <c>Marshal.GetActiveObject</c>,
/// so both are provided here via P/Invoke into <c>ole32</c>/<c>oleaut32</c>.
///
/// This mirrors the reference server's <c>get_word_app</c>/<c>_find_word_with_docs</c>:
/// <c>GetActiveObject</c> first, then a full ROT enumeration to rescue the O365/OneDrive
/// case where the active proxy reports zero documents.
///
/// All callers must be on the STA thread (<see cref="StaDispatcher"/>).
/// </summary>
internal static class RunningObjectTable
{
    /// <summary>
    /// Return the active object registered for <paramref name="progId"/> (e.g.
    /// <c>"Word.Application"</c>), or <c>null</c> when none is registered.
    /// Replacement for the removed <c>Marshal.GetActiveObject</c>.
    /// </summary>
    public static object? GetActiveObject(string progId)
    {
        try
        {
            CLSIDFromProgID(progId, out var clsid);
            GetActiveObject(ref clsid, IntPtr.Zero, out var instance);
            return instance;
        }
        catch (COMException)
        {
            // MK_E_UNAVAILABLE / no registered active object.
            return null;
        }
    }

    /// <summary>
    /// Enumerate every entry in the Running Object Table as
    /// (displayName, boundObject) pairs. Binding a moniker can throw for stale or
    /// foreign entries, so each is wrapped defensively and skipped on failure.
    /// </summary>
    public static IEnumerable<(string DisplayName, object Instance)> Enumerate()
    {
        var hr = GetRunningObjectTable(0, out var rot);
        if (hr != 0 || rot is null)
        {
            yield break;
        }

        COMTypes.IEnumMoniker? enumMoniker = null;
        try
        {
            rot.EnumRunning(out enumMoniker);
            if (enumMoniker is null)
            {
                yield break;
            }

            enumMoniker.Reset();
            var monikers = new COMTypes.IMoniker[1];
            var fetched = IntPtr.Zero;

            while (enumMoniker.Next(1, monikers, fetched) == 0)
            {
                var moniker = monikers[0];
                var entry = TryBind(rot, moniker);
                if (moniker is not null)
                {
                    Marshal.ReleaseComObject(moniker);
                }

                if (entry is not null)
                {
                    yield return entry.Value;
                }
            }
        }
        finally
        {
            if (enumMoniker is not null)
            {
                Marshal.ReleaseComObject(enumMoniker);
            }

            Marshal.ReleaseComObject(rot);
        }
    }

    private static (string DisplayName, object Instance)? TryBind(
        COMTypes.IRunningObjectTable rot, COMTypes.IMoniker moniker)
    {
        COMTypes.IBindCtx? bindCtx = null;
        try
        {
            if (CreateBindCtx(0, out bindCtx) != 0 || bindCtx is null)
            {
                return null;
            }

            moniker.GetDisplayName(bindCtx, null, out var displayName);
            if (rot.GetObject(moniker, out var instance) != 0 || instance is null)
            {
                return null;
            }

            return (displayName ?? string.Empty, instance);
        }
        catch (COMException)
        {
            return null;
        }
        finally
        {
            if (bindCtx is not null)
            {
                Marshal.ReleaseComObject(bindCtx);
            }
        }
    }

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out COMTypes.IRunningObjectTable prot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out COMTypes.IBindCtx ppbc);

    [DllImport("ole32.dll", PreserveSig = false)]
    private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        ref Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
}
