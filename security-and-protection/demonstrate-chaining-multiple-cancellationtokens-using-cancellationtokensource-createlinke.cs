using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    // Simulates a long‑running operation that observes cancellation.
    private static async Task SimulateWorkAsync(string operationName, CancellationToken token)
    {
        Console.WriteLine($"{operationName} started.");
        try
        {
            // Perform work in 5 steps, checking for cancellation between each step.
            for (int i = 1; i <= 5; i++)
            {
                token.ThrowIfCancellationRequested();
                await Task.Delay(TimeSpan.FromSeconds(1), token); // Simulated work.
                Console.WriteLine($"{operationName} progress: step {i}/5");
            }

            Console.WriteLine($"{operationName} completed successfully.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine($"{operationName} was cancelled.");
        }
    }

    public static async Task Main(string[] args)
    {
        // Individual cancellation sources representing different reasons to cancel.
        using var ctsOverall = new CancellationTokenSource();   // Overall application shutdown.
        using var ctsUser = new CancellationTokenSource();      // User‑initiated cancellation.
        using var ctsTimeout = new CancellationTokenSource();   // Timeout‑based cancellation.

        // Create a linked token that is cancelled when any of the sources is cancelled.
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(
            ctsOverall.Token, ctsUser.Token, ctsTimeout.Token);

        // Start two independent operations that share the same linked token.
        Task op1 = SimulateWorkAsync("Operation 1", linkedCts.Token);
        Task op2 = SimulateWorkAsync("Operation 2", linkedCts.Token);

        // Simulate external events that trigger cancellation.
        // After 2.5 seconds, cancel the user token.
        await Task.Delay(TimeSpan.FromSeconds(2.5));
        Console.WriteLine("User requested cancellation.");
        ctsUser.Cancel();

        // After 5 seconds, trigger a timeout cancellation (won't have effect because already cancelled).
        await Task.Delay(TimeSpan.FromSeconds(2.5));
        Console.WriteLine("Timeout reached.");
        ctsTimeout.Cancel();

        // Wait for both operations to finish handling cancellation.
        await Task.WhenAll(op1, op2);

        // Finally, cancel the overall token (clean‑up scenario).
        ctsOverall.Cancel();

        Console.WriteLine("Program finished.");
    }
}
