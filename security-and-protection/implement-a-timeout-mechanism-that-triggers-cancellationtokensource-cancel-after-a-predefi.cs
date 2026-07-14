using System;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;

public class Program
{
    public static void Main()
    {
        // Create a cancellation token source that will be cancelled after a timeout.
        using var cts = new CancellationTokenSource();

        // Define the timeout duration (e.g., 3 seconds).
        const double timeoutMilliseconds = 3000;

        // Set up a timer that triggers the cancellation.
        // Use the fully qualified System.Timers.Timer to avoid ambiguity with System.Threading.Timer.
        using var timer = new System.Timers.Timer(timeoutMilliseconds) { AutoReset = false };
        timer.Elapsed += (sender, e) => cts.Cancel();
        timer.Start();

        // Simulate a long‑running operation that observes the cancellation token.
        try
        {
            RunLongOperationAsync(cts.Token).GetAwaiter().GetResult();
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Operation was cancelled due to timeout.");
        }

        // Ensure the timer is stopped before exiting.
        timer.Stop();
    }

    private static async Task RunLongOperationAsync(CancellationToken token)
    {
        // Perform work in a loop, checking for cancellation.
        while (true)
        {
            // Throw if cancellation has been requested.
            token.ThrowIfCancellationRequested();

            // Simulated work (e.g., processing).
            Console.WriteLine("Working...");
            await Task.Delay(500, token); // Delay respects the token.
        }
    }
}
