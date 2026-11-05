using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    private static readonly CancellationTokenSource _cts = new CancellationTokenSource();

    // Simulated UI button click that cancels the operation.
    private static void CancelButton_Click()
    {
        Console.WriteLine("Cancel button clicked.");
        _cts.Cancel();
    }

    // Example long‑running operation that respects cancellation.
    private static async Task LongRunningOperation(CancellationToken token)
    {
        Console.WriteLine("Operation started.");
        try
        {
            for (int i = 1; i <= 10; i++)
            {
                token.ThrowIfCancellationRequested();
                Console.WriteLine($"Processing step {i}");
                await Task.Delay(500, token); // Simulate work.
            }

            Console.WriteLine("Operation completed.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Operation was cancelled.");
        }
    }

    public static async Task Main()
    {
        // Start the long‑running task.
        Task operationTask = LongRunningOperation(_cts.Token);

        // Simulate a user pressing the cancel button after a short delay.
        await Task.Delay(1500);
        CancelButton_Click();

        // Wait for the operation to finish handling cancellation.
        await operationTask;
    }
}
