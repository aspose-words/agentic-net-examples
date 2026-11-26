using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    // Entry point of the console application.
    public static async Task Main()
    {
        // Create a CancellationTokenSource that will be cancelled automatically after a timeout.
        using var cts = new CancellationTokenSource();

        // Define the timeout duration (e.g., 3 seconds).
        TimeSpan timeout = TimeSpan.FromSeconds(3);
        cts.CancelAfter(timeout); // Schedules automatic cancellation.

        try
        {
            // Run a sample long‑running operation that observes the cancellation token.
            await PerformLongRunningOperationAsync(cts.Token);
        }
        catch (OperationCanceledException)
        {
            // This block is reached when the operation is cancelled due to the timeout.
            Console.WriteLine("Operation cancelled after timeout.");
        }
    }

    // Simulates a long‑running task that periodically checks the cancellation token.
    private static async Task PerformLongRunningOperationAsync(CancellationToken token)
    {
        // Loop to represent work steps.
        for (int i = 1; i <= 10; i++)
        {
            // Throw if cancellation has been requested.
            token.ThrowIfCancellationRequested();

            Console.WriteLine($"Processing step {i}...");

            // Simulate work by delaying for 1 second, also observing the token.
            await Task.Delay(TimeSpan.FromSeconds(1), token);
        }

        Console.WriteLine("Operation completed without cancellation.");
    }
}
