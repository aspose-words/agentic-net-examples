using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

public class Program
{
    public static void Main()
    {
        // Create a cancellation token that will be triggered after a short delay
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(200); // cancel after 200 ms

        // Simulate a large data source
        IEnumerable<int> data = Enumerable.Range(1, 5_000_000);

        try
        {
            // Custom LINQ query that checks token.IsCancellationRequested
            List<int> report = data
                .AsParallel()
                .Select(item =>
                {
                    // Abort if cancellation was requested
                    if (cts.Token.IsCancellationRequested)
                        throw new OperationCanceledException(cts.Token);

                    // Simulate some processing work
                    Thread.SpinWait(10_000);
                    return item * 2;
                })
                .Take(100) // limit the result size for the demo
                .ToList();

            Console.WriteLine($"Report generated with {report.Count} items.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Report generation was cancelled.");
        }
    }
}
