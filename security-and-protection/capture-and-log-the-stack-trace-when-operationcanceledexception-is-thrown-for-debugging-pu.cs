using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    public static void Main()
    {
        var cts = new CancellationTokenSource();
        var token = cts.Token;

        // Cancel the token shortly after start.
        Task.Run(async () =>
        {
            await Task.Delay(100);
            cts.Cancel();
        });

        try
        {
            PerformOperation(token);
        }
        catch (OperationCanceledException ex)
        {
            // Capture the stack trace.
            string stackTrace = ex.StackTrace ?? "No stack trace available.";

            // Log details to a file.
            string logPath = "error.log";
            string logContent = $"[{DateTime.Now:O}] OperationCanceledException caught.{Environment.NewLine}" +
                                $"Message: {ex.Message}{Environment.NewLine}" +
                                $"StackTrace:{Environment.NewLine}{stackTrace}{Environment.NewLine}";
            File.WriteAllText(logPath, logContent);

            // Also write a brief notice to the console.
            Console.WriteLine($"Exception logged to {logPath}");
        }
    }

    private static void PerformOperation(CancellationToken token)
    {
        // Simulate work that periodically checks for cancellation.
        for (int i = 0; i < 10; i++)
        {
            token.ThrowIfCancellationRequested();
            Thread.Sleep(50); // Simulated work.
        }
    }
}
