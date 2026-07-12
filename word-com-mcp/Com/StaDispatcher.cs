using System.Collections.Concurrent;

namespace WordComMcp.Com;

/// <summary>
/// Serializes <b>every</b> COM call onto a single dedicated STA background thread
/// (issue 0.4). Word's automation objects are apartment-threaded; touching them
/// from arbitrary thread-pool threads yields cross-apartment errors
/// (<c>0x8001010E RPC_E_WRONGTHREAD</c>). A single STA thread fed by a work queue
/// guarantees affinity and serialization.
///
/// Performance rule (Conventions #6): do <i>all</i> of a tool's COM work inside one
/// <see cref="RunOnStaAsync{T}"/> lambda so each tool costs one coarse STA hop.
/// </summary>
public sealed class StaDispatcher : IDisposable
{
    private readonly BlockingCollection<Action> m_queue = new(new ConcurrentQueue<Action>());
    private readonly Thread m_thread;
    private bool m_disposed;

    public StaDispatcher()
    {
        this.m_thread = new Thread(this.PumpQueue)
        {
            IsBackground = true,
            Name = "WordComMcp-STA",
        };
        this.m_thread.SetApartmentState(ApartmentState.STA);
        this.m_thread.Start();
    }

    /// <summary>
    /// Run <paramref name="work"/> on the STA thread and await its result. Exceptions
    /// thrown by <paramref name="work"/> are marshalled back to the awaiter.
    /// </summary>
    public Task<T> RunOnStaAsync<T>(Func<T> work)
    {
        ArgumentNullException.ThrowIfNull(work);
        ObjectDisposedException.ThrowIf(this.m_disposed, this);

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        this.m_queue.Add(() =>
        {
            try
            {
                tcs.SetResult(work());
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });
        return tcs.Task;
    }

    /// <summary>Run a void COM sequence on the STA thread and await completion.</summary>
    public Task RunOnStaAsync(Action work)
    {
        ArgumentNullException.ThrowIfNull(work);
        return this.RunOnStaAsync(() =>
        {
            work();
            return true;
        });
    }

    private void PumpQueue()
    {
        foreach (var action in this.m_queue.GetConsumingEnumerable())
        {
            action();
        }
    }

    public void Dispose()
    {
        if (this.m_disposed)
        {
            return;
        }

        this.m_disposed = true;
        this.m_queue.CompleteAdding();

        // Give the pump a moment to drain; it is a background thread either way.
        if (!this.m_thread.Join(TimeSpan.FromSeconds(5)))
        {
            return;
        }

        this.m_queue.Dispose();
    }
}
