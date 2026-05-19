using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    private const string AuditFileName = "audit.log";

    private static void LogCancellation(string message)
    {
        string entry = $"{DateTime.UtcNow:O} - {message}{Environment.NewLine}";
        File.AppendAllText(AuditFileName, entry);
    }

    public static void Main()
    {
        // Ensure the audit file starts empty for this run.
        if (File.Exists(AuditFileName))
            File.Delete(AuditFileName);

        // Example 1: Cancellation triggered by a timeout.
        var ctsTimeout = new CancellationTokenSource();
        ctsTimeout.CancelAfter(100); // Cancel after 100 ms.

        try
        {
            // This will be cancelled before the delay completes.
            Task.Delay(1000, ctsTimeout.Token).Wait();
        }
        catch (AggregateException ae) when (ae.InnerException is OperationCanceledException)
        {
            LogCancellation("Operation cancelled due to timeout.");
        }
        catch (OperationCanceledException)
        {
            LogCancellation("Operation cancelled due to timeout.");
        }

        // Example 2: Manual cancellation.
        var ctsManual = new CancellationTokenSource();

        var task = Task.Run(async () =>
        {
            // Simulate work that checks for cancellation.
            for (int i = 0; i < 5; i++)
            {
                ctsManual.Token.ThrowIfCancellationRequested();
                await Task.Delay(200, ctsManual.Token);
            }
        }, ctsManual.Token);

        // Cancel the operation after a short delay.
        Task.Delay(350).ContinueWith(_ => ctsManual.Cancel());

        try
        {
            task.Wait();
        }
        catch (AggregateException ae) when (ae.InnerException is OperationCanceledException)
        {
            LogCancellation("Operation manually cancelled.");
        }
        catch (OperationCanceledException)
        {
            LogCancellation("Operation manually cancelled.");
        }

        // Verify that the audit file was created.
        if (!File.Exists(AuditFileName))
            throw new InvalidOperationException("Audit log was not created.");

        // The program ends here; no user interaction required.
    }
}
