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

        // Create a linked token source that combines both tokens.
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cts1.Token, cts2.Token);
        CancellationToken token = linkedCts.Token;

        // Start a background operation that respects the linked token.
        Task worker = Task.Run(() => DoWork(token), token);

        // Simulate some work, then cancel the first source.
        await Task.Delay(500);
        Console.WriteLine("Cancelling first token source (cts1).");
        cts1.Cancel();

        // Give the worker a moment to observe the cancellation.
        await Task.Delay(200);
        Console.WriteLine($"Is cancellation requested? {token.IsCancellationRequested}");

        // Cancel the second source as well.
        Console.WriteLine("Cancelling second token source (cts2).");
        cts2.Cancel();

        // Wait for the worker to finish.
        await worker;

        Console.WriteLine("All done.");
    }

    private static void DoWork(CancellationToken token)
    {
        int iteration = 0;
        while (true)
        {
            if (token.IsCancellationRequested)
            {
                Console.WriteLine($"Cancellation detected after {iteration} iterations.");
                break;
            }

            // Simulate work.
            Thread.Sleep(100);
            iteration++;
            Console.WriteLine($"Working... iteration {iteration}");
        }
    }
}
