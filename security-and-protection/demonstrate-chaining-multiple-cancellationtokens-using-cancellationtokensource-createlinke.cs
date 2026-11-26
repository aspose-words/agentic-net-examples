using System;
using System.Threading;
using System.Threading.Tasks;

public class Program
{
    public static async Task Main()
    {
        // First demonstration: cancel via the first source
        using var cts1 = new CancellationTokenSource();
        using var cts2 = new CancellationTokenSource();
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cts1.Token, cts2.Token);
        CancellationToken linkedToken = linkedCts.Token;

        Task operation = LongRunningOperationAsync(linkedToken);

        // Cancel cts1 after a short delay; this propagates to the linked token
        _ = Task.Run(async () =>
        {
            await Task.Delay(500);
            cts1.Cancel();
        });

        try
        {
            await operation;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Operation was cancelled via linked token (first source).");
        }

        // Second demonstration: cancel via the second source
        using var cts3 = new CancellationTokenSource();
        using var cts4 = new CancellationTokenSource();
        using var linkedCts2 = CancellationTokenSource.CreateLinkedTokenSource(cts3.Token, cts4.Token);
        CancellationToken linkedToken2 = linkedCts2.Token;

        Task operation2 = LongRunningOperationAsync(linkedToken2);

        // Cancel cts4 after a short delay; this also propagates to the linked token
        _ = Task.Run(async () =>
        {
            await Task.Delay(300);
            cts4.Cancel();
        });

        try
        {
            await operation2;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Operation was cancelled via linked token (second source).");
        }
    }

    private static async Task LongRunningOperationAsync(CancellationToken token)
    {
        // Simulate work in steps, checking for cancellation
        for (int i = 0; i < 10; i++)
        {
            token.ThrowIfCancellationRequested();
            Console.WriteLine($"Working... step {i + 1}");
            await Task.Delay(200, token);
        }

        Console.WriteLine("Operation completed successfully.");
    }
}
