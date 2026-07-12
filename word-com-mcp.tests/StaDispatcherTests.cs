using System.Collections.Concurrent;
using WordComMcp.Com;
using Xunit;

namespace WordComMcp.Tests;

/// <summary>STA dispatcher behavior (issue 0.4) — no Word required.</summary>
public class StaDispatcherTests
{
    [Fact]
    public async Task RunOnStaAsync_ExecutesOnASingleStaThread()
    {
        // Arrange
        using var dispatcher = new StaDispatcher();

        // Act
        var apartment = await dispatcher.RunOnStaAsync(() => Thread.CurrentThread.GetApartmentState());

        // Assert
        Assert.Equal(ApartmentState.STA, apartment);
    }

    [Fact]
    public async Task RunOnStaAsync_SerializesConcurrentCallsOntoTheSameThread()
    {
        // Arrange
        using var dispatcher = new StaDispatcher();
        var threadIds = new ConcurrentBag<int>();

        // Act — fire many calls concurrently; all must land on the one STA thread.
        var tasks = Enumerable.Range(0, 32).Select(_ => dispatcher.RunOnStaAsync(() =>
        {
            threadIds.Add(Environment.CurrentManagedThreadId);
            return 0;
        }));
        await Task.WhenAll(tasks);

        // Assert
        Assert.Single(threadIds.Distinct());
    }

    [Fact]
    public async Task RunOnStaAsync_MarshalsExceptionsBackToTheCaller()
    {
        // Arrange
        using var dispatcher = new StaDispatcher();

        // Act / Assert
        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            dispatcher.RunOnStaAsync<int>(() => throw new InvalidOperationException("boom")));
    }
}
