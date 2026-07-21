using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    public static void Main()
    {
        // Create a cancellation token source.
        var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Simulate a low‑code UI button that triggers cancellation after a short delay.
        Task.Run(async () =>
        {
            await Task.Delay(500); // wait before "clicking" the button
            Console.WriteLine("Button clicked: cancelling operation.");
            cts.Cancel(); // stop the processing immediately
        });

        // Example long‑running processing loop that observes the token.
        try
        {
            for (int i = 0; i < 10; i++)
            {
                token.ThrowIfCancellationRequested(); // abort if cancelled
                Console.WriteLine($"Processing step {i + 1}");
                Thread.Sleep(300); // simulate work
            }

            Console.WriteLine("Processing completed successfully.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Processing was cancelled.");
        }
    }
}
