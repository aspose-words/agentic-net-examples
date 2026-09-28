using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    public static async Task Main()
    {
        var cts = new CancellationTokenSource();

        // Simulate a UI button that cancels the operation after a short delay
        Task buttonTask = Task.Run(async () =>
        {
            await Task.Delay(1000); // wait 1 second before "click"
            cts.Cancel(); // button click triggers cancellation
        });

        // Long-running work that observes the cancellation token
        Task workTask = Task.Run(() => DoWork(cts.Token));

        await Task.WhenAll(buttonTask, workTask);
    }

    private static void DoWork(CancellationToken token)
    {
        try
        {
            for (int i = 0; i < 10; i++)
            {
                token.ThrowIfCancellationRequested();
                // Simulate work
                Thread.Sleep(500);
                Console.WriteLine($"Working... step {i + 1}");
            }

            Console.WriteLine("Work completed.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Work was cancelled.");
        }
    }
}
