using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    public static async Task Main()
    {
        // Create a CancellationTokenSource that we can cancel manually.
        using var manualCts = new CancellationTokenSource();

        // Create a second CancellationTokenSource that will cancel automatically after 3 seconds.
        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(3));

        // Link the two tokens so that cancellation of either source will cancel the linked token.
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(manualCts.Token, timeoutCts.Token);
        CancellationToken token = linkedCts.Token;

        // Start a background operation that respects the linked cancellation token.
        Task operation = DoWorkAsync(token);

        // Simulate some work in the main thread, then cancel manually after 1 second.
        await Task.Delay(1000);
        Console.WriteLine("Main thread: requesting manual cancellation.");
        manualCts.Cancel();

        // Wait for the operation to finish, handling cancellation gracefully.
        try
        {
            await operation;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Main thread: operation was cancelled.");
        }

        Console.WriteLine("Main thread: operation completed. Exiting.");
        // Brief pause to ensure output is visible before the program ends.
        await Task.Delay(500);
    }

    private static async Task DoWorkAsync(CancellationToken token)
    {
        Console.WriteLine("Operation started.");

        for (int i = 0; i < 10; i++)
        {
            // Check for cancellation before each iteration.
            if (token.IsCancellationRequested)
            {
                Console.WriteLine($"Operation cancelled after {i} iteration(s).");
                token.ThrowIfCancellationRequested();
            }

            Console.WriteLine($"Operation working... iteration {i + 1}");
            // Simulate work.
            await Task.Delay(500, token);
        }

        Console.WriteLine("Operation completed successfully.");
    }
}
