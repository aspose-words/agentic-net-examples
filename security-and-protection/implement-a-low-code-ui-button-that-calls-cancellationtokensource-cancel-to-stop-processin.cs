using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    // Simulated long‑running operation that respects cancellation.
    private static async Task ProcessAsync(CancellationToken token)
    {
        int i = 0;
        try
        {
            while (true)
            {
                token.ThrowIfCancellationRequested();
                Console.WriteLine($"Processing step {i++}");
                // Simulate work.
                await Task.Delay(200, token);
            }
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Processing was cancelled.");
        }
    }

    // Simulated low‑code UI button click handler.
    private static void OnCancelButtonClick(CancellationTokenSource cts)
    {
        // Immediately request cancellation.
        cts.Cancel();
        Console.WriteLine("Cancel button clicked.");
    }

    public static async Task Main(string[] args)
    {
        using var cts = new CancellationTokenSource();

        // Start the processing task.
        Task processingTask = ProcessAsync(cts.Token);

        // Simulate a short delay before the user clicks the cancel button.
        await Task.Delay(500);
        OnCancelButtonClick(cts);

        // Wait for the processing task to finish cleanly.
        await processingTask;
    }
}
