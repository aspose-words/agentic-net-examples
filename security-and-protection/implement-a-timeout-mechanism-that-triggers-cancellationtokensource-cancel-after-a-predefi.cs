using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    // Simulates a long‑running operation that respects cancellation.
    private static async Task LongRunningOperationAsync(CancellationToken token)
    {
        for (int i = 1; i <= 10; i++)
        {
            // Throw if cancellation has been requested.
            token.ThrowIfCancellationRequested();

            // Simulate work.
            await Task.Delay(1000, token);
            Console.WriteLine($"Iteration {i}");
        }
    }

    public static async Task Main(string[] args)
    {
        // Define a timeout duration.
        TimeSpan timeout = TimeSpan.FromSeconds(3);

        // Create a CancellationTokenSource that will cancel after the timeout.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(timeout);

        try
        {
            // Run the operation with the cancellation token.
            await LongRunningOperationAsync(cts.Token);
            Console.WriteLine("Operation completed without cancellation.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Operation cancelled due to timeout.");
        }
    }
}
