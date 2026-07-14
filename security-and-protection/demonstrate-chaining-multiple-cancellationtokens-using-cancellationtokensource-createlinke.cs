using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    // Simulated long‑running operation that respects cancellation.
    private static async Task PerformWorkAsync(string name, CancellationToken token)
    {
        Console.WriteLine($"{name}: started.");
        try
        {
            // Loop with short delays to periodically check the token.
            for (int i = 1; i <= 10; i++)
            {
                token.ThrowIfCancellationRequested();
                await Task.Delay(200, token); // 200 ms per step
                Console.WriteLine($"{name}: progress {i * 10}%");
            }

            Console.WriteLine($"{name}: completed successfully.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine($"{name}: cancelled.");
            throw;
        }
    }

    public static async Task Main(string[] args)
    {
        // First cancellation source – could represent a user‑initiated cancel.
        using var ctsUser = new CancellationTokenSource();

        // Second cancellation source – could represent a timeout.
        using var ctsTimeout = new CancellationTokenSource();

        // Create a linked token source that observes both tokens.
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(ctsUser.Token, ctsTimeout.Token);

        // Start a task that uses the linked token.
        var workTask = PerformWorkAsync("LinkedTask", linkedCts.Token);

        // Schedule cancellation of the timeout token after 1 second.
        _ = Task.Run(async () =>
        {
            await Task.Delay(1000);
            Console.WriteLine("Timeout token: cancelling.");
            ctsTimeout.Cancel();
        });

        // Optionally, simulate a user cancel after 1.5 seconds (comment out to test only timeout).
        _ = Task.Run(async () =>
        {
            await Task.Delay(1500);
            Console.WriteLine("User token: cancelling.");
            ctsUser.Cancel();
        });

        try
        {
            await workTask;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Main: work was cancelled via linked token.");
        }

        // Demonstrate chaining a third token.
        using var ctsExternal = new CancellationTokenSource();
        using var secondLinked = CancellationTokenSource.CreateLinkedTokenSource(linkedCts.Token, ctsExternal.Token);

        var secondTask = PerformWorkAsync("SecondLinkedTask", secondLinked.Token);

        // Cancel the external token after a short delay.
        _ = Task.Run(async () =>
        {
            await Task.Delay(500);
            Console.WriteLine("External token: cancelling.");
            ctsExternal.Cancel();
        });

        try
        {
            await secondTask;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Main: second work was cancelled via second linked token.");
        }

        Console.WriteLine("Program finished.");
    }
}
