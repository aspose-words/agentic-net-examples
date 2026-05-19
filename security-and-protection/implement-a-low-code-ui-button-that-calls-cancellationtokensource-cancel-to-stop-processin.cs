using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    public static void Main()
    {
        // Create a cancellation token source that will be used to cancel the operation.
        using var cts = new CancellationTokenSource();

        // Start a long‑running task that checks the token periodically.
        Task processingTask = Task.Run(() => DoWork(cts.Token), cts.Token);

        // Simulate a low‑code UI button press by cancelling after a short delay.
        // In a real UI this would be wired to a button click event.
        Task.Delay(500).ContinueWith(_ => cts.Cancel());

        try
        {
            // Wait for the task to complete or be cancelled.
            processingTask.Wait();
        }
        catch (AggregateException ae)
        {
            // The task throws an OperationCanceledException when cancelled.
            ae.Handle(e => e is OperationCanceledException);
        }

        Console.WriteLine("Processing finished.");
    }

    private static void DoWork(CancellationToken token)
    {
        int iteration = 0;
        while (true)
        {
            // Throw if cancellation has been requested.
            token.ThrowIfCancellationRequested();

            // Simulate work.
            Console.WriteLine($"Working... iteration {++iteration}");
            Thread.Sleep(200); // Short pause to mimic work.
        }
    }
}
