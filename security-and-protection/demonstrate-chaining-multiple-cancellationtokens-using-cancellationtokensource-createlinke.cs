using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    // Simulated workload that respects cancellation.
    private static async Task SimulateWorkAsync(string name, CancellationToken token)
    {
        try
        {
            int iteration = 0;
            while (true)
            {
                token.ThrowIfCancellationRequested();
                Console.WriteLine($"{name}: iteration {++iteration}");
                await Task.Delay(500, token); // Simulate work.
            }
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine($"{name}: cancelled.");
        }
    }

    public static async Task Main(string[] args)
    {
        // Create two independent cancellation sources.
        using var cts1 = new CancellationTokenSource();
        using var cts2 = new CancellationTokenSource();

        // Link them so that cancellation of either source cancels the linked token.
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cts1.Token, cts2.Token);
        CancellationToken linkedToken = linkedCts.Token;

        // Start a task that observes the linked token.
        Task workTask = SimulateWorkAsync("LinkedWork", linkedToken);

        // Cancel the first source after 2 seconds.
        await Task.Delay(2000);
        Console.WriteLine("Cancelling first token source (cts1).");
        cts1.Cancel();

        // Give the task a moment to observe cancellation.
        await Task.Delay(1000);

        // Restart the work with a fresh linked token to demonstrate second cancellation.
        using var cts3 = new CancellationTokenSource();
        using var cts4 = new CancellationTokenSource();
        using var linkedCts2 = CancellationTokenSource.CreateLinkedTokenSource(cts3.Token, cts4.Token);
        Task workTask2 = SimulateWorkAsync("LinkedWork2", linkedCts2.Token);

        // Cancel the second source after 1.5 seconds.
        await Task.Delay(1500);
        Console.WriteLine("Cancelling second token source (cts4).");
        cts4.Cancel();

        // Wait for both tasks to finish.
        await Task.WhenAll(workTask, workTask2);

        Console.WriteLine("All work completed. Program exiting.");
    }
}
