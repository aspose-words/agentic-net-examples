using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    public static async Task Main()
    {
        // Create two independent cancellation token sources.
        using var cts1 = new CancellationTokenSource();
        using var cts2 = new CancellationTokenSource();

        // Create a linked token source that combines cts1 and cts2.
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cts1.Token, cts2.Token);
        CancellationToken linkedToken = linkedCts.Token;

        // Start a long-running operation that observes the linked token.
        Task operation = LongRunningOperationAsync(linkedToken);

        // Cancel the first source after 500 ms.
        _ = Task.Run(async () =>
        {
            await Task.Delay(500);
            Console.WriteLine("Cancelling first token source (cts1).");
            cts1.Cancel();
        });

        // Cancel the second source after 1000 ms (if not already cancelled).
        _ = Task.Run(async () =>
        {
            await Task.Delay(1000);
            Console.WriteLine("Cancelling second token source (cts2).");
            cts2.Cancel();
        });

        try
        {
            await operation;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Operation was cancelled via linked token.");
        }

        Console.WriteLine("Program completed.");
    }

    private static async Task LongRunningOperationAsync(CancellationToken token)
    {
        Console.WriteLine("Long-running operation started.");
        // Simulate work in a loop, checking for cancellation.
        for (int i = 0; i < 10; i++)
        {
            token.ThrowIfCancellationRequested();
            Console.WriteLine($"Working... step {i + 1}");
            await Task.Delay(300, token); // Pass token to delay to observe cancellation promptly.
        }
        Console.WriteLine("Long-running operation completed successfully.");
    }
}
