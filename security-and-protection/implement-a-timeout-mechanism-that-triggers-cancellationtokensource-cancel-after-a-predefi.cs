using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    // Simulated long‑running operation that respects a cancellation token.
    private static async Task PerformWorkAsync(CancellationToken token)
    {
        Console.WriteLine("Work started.");
        // Simulate work in 500 ms increments.
        for (int i = 0; i < 20; i++)
        {
            // Throw if cancellation was requested.
            token.ThrowIfCancellationRequested();

            await Task.Delay(500, token); // Respect the token during the delay.
            Console.WriteLine($"Progress: {(i + 1) * 5}%");
        }

        Console.WriteLine("Work completed successfully.");
    }

    public static async Task Main()
    {
        // Define a timeout after which the operation should be cancelled.
        TimeSpan timeout = TimeSpan.FromSeconds(3);

        using var cts = new CancellationTokenSource();

        // Automatically cancel after the specified timeout.
        cts.CancelAfter(timeout);
        Console.WriteLine($"Cancellation will be triggered after {timeout.TotalSeconds} seconds.");

        try
        {
            // Run the work task and await its completion.
            await PerformWorkAsync(cts.Token);
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Operation was cancelled due to timeout.");
        }
        finally
        {
            // Ensure the token source is disposed.
            cts.Dispose();
        }

        // Give a brief moment for console output to flush before exiting.
        await Task.Delay(500);
    }
}
