using System;
using System.IO;
using System.Threading;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a cancellation token source and cancel it immediately.
        var cts = new CancellationTokenSource();
        cts.Cancel();

        try
        {
            // Perform an operation that checks the token and throws if cancelled.
            PerformOperation(cts.Token);
        }
        catch (OperationCanceledException ex)
        {
            // Capture and log the stack trace for debugging.
            string logMessage = $"OperationCanceledException caught at {DateTime.Now}{Environment.NewLine}" +
                                $"Message: {ex.Message}{Environment.NewLine}" +
                                $"Stack Trace:{Environment.NewLine}{ex.StackTrace}{Environment.NewLine}";

            // Write to console.
            Console.WriteLine(logMessage);

            // Also write to a log file in the current directory.
            string logFilePath = Path.Combine(Directory.GetCurrentDirectory(), "exception_log.txt");
            File.WriteAllText(logFilePath, logMessage);
        }
    }

    private static void PerformOperation(CancellationToken token)
    {
        // This will throw OperationCanceledException if the token is cancelled.
        token.ThrowIfCancellationRequested();

        // Simulate work (won't be reached because token is already cancelled).
        Console.WriteLine("Operation completed.");
    }
}
