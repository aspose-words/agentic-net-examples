using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    // Predefined timeout duration (in milliseconds).
    private const int TimeoutMilliseconds = 2000;

    public static void Main(string[] args)
    {
        // Run the asynchronous workflow synchronously.
        RunAsync().GetAwaiter().GetResult();
    }

    private static async Task RunAsync()
    {
        // Create a CancellationTokenSource that will be cancelled after the timeout.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(TimeoutMilliseconds);

        try
        {
            // Simulate a long-running operation that respects cancellation.
            await PerformLongRunningOperationAsync(cts.Token);
        }
        catch (OperationCanceledException)
        {
            // The operation was cancelled due to the timeout.
            // No further action needed; the program will exit gracefully.
        }
    }

    private static async Task PerformLongRunningOperationAsync(CancellationToken token)
    {
        // Example loop that performs work in small increments.
        for (int i = 0; i < 10; i++)
        {
            // Throw if cancellation has been requested.
            token.ThrowIfCancellationRequested();

            // Simulate work by delaying for 500 ms.
            await Task.Delay(500, token);
        }
    }
}
