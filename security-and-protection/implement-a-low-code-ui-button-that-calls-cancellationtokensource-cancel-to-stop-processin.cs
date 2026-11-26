using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    // Simulated UI button that cancels the operation.
    private static void CancelButtonClick(CancellationTokenSource cts)
    {
        Console.WriteLine("Cancel button clicked.");
        cts.Cancel();
    }

    // Example long‑running process that respects cancellation.
    private static async Task LongRunningOperationAsync(CancellationToken token)
    {
        Console.WriteLine("Operation started.");
        for (int i = 0; i < 10; i++)
        {
            token.ThrowIfCancellationRequested();
            await Task.Delay(300, token); // Simulate work.
            Console.WriteLine($"Step {i + 1} completed.");
        }
        Console.WriteLine("Operation completed successfully.");
    }

    public static async Task Main(string[] args)
    {
        using var cts = new CancellationTokenSource();

        // Start the operation.
        var operationTask = LongRunningOperationAsync(cts.Token);

        // Simulate UI button press after a short delay.
        _ = Task.Run(async () =>
        {
            await Task.Delay(800);
            CancelButtonClick(cts);
        });

        try
        {
            await operationTask;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Operation was cancelled.");
        }
    }
}
