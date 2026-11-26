using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    public static void Main()
    {
        // Create a cancellation token source that cancels after a short delay.
        using var cts = new CancellationTokenSource(100);
        try
        {
            // Execute a cancellable operation synchronously.
            PerformCancellableOperation(cts.Token).Wait();
        }
        catch (AggregateException aggEx)
        {
            // Unwrap and handle OperationCanceledException.
            foreach (var ex in aggEx.InnerExceptions)
            {
                if (ex is OperationCanceledException oce)
                {
                    LogException(oce);
                }
                else
                {
                    // Re‑throw any other unexpected exceptions.
                    throw;
                }
            }
        }
    }

    private static async Task PerformCancellableOperation(CancellationToken token)
    {
        // Simulate work that periodically checks for cancellation.
        for (int i = 0; i < 10; i++)
        {
            token.ThrowIfCancellationRequested();
            await Task.Delay(50, token);
        }
    }

    private static void LogException(OperationCanceledException ex)
    {
        // Capture the stack trace.
        string stackTrace = ex.StackTrace ?? "No stack trace available";

        // Log details to the console.
        Console.WriteLine("OperationCanceledException caught.");
        Console.WriteLine("Message: " + ex.Message);
        Console.WriteLine("StackTrace:");
        Console.WriteLine(stackTrace);

        // Also write the information to a file for later analysis.
        File.WriteAllText(
            "cancellation_log.txt",
            $"Message: {ex.Message}{Environment.NewLine}StackTrace:{Environment.NewLine}{stackTrace}"
        );
    }
}
