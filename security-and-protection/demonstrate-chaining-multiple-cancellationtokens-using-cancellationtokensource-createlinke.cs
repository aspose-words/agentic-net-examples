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

        // Link the two tokens into a single token source.
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cts1.Token, cts2.Token);
        CancellationToken linkedToken = linkedCts.Token;

        // Start a long‑running operation that observes the linked token.
        Task workTask = DoWorkAsync(linkedToken);

        // Cancel the first source after a short delay.
        _ = Task.Run(async () =>
        {
            await Task.Delay(500);
            Console.WriteLine("Cancelling cts1");
            cts1.Cancel();
        });

        // Cancel the second source after a longer delay (won't fire if the first already cancelled).
        _ = Task.Run(async () =>
        {
            await Task.Delay(1500);
            Console.WriteLine("Cancelling cts2");
            cts2.Cancel();
        });

        try
        {
            await workTask;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Operation was cancelled via linked token.");
        }

        Console.WriteLine("Program completed.");
    }

    private static async Task DoWorkAsync(CancellationToken token)
    {
        for (int i = 0; i < 5; i++)
        {
            token.ThrowIfCancellationRequested();
            Console.WriteLine($"Working... step {i + 1}");
            await Task.Delay(400, token); // Simulate work.
        }

        Console.WriteLine("Work completed successfully.");
    }
}
