using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a large document to make layout generation take noticeable time.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Set up a cancellation token that will be triggered shortly.
        var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Run layout generation on a background task.
        Task layoutTask = Task.Run(() =>
        {
            // Aspose.Words does not accept a cancellation token for layout generation,
            // so we simply call the method. The token can be used to cancel the task
            // before it starts, but once the method is executing it will run to completion.
            doc.UpdatePageLayout();
        }, token);

        // Give the layout a moment to start, then request cancellation.
        Thread.Sleep(10);
        cts.Cancel();

        try
        {
            // Wait for the task to complete.
            layoutTask.Wait();

            // If the task completed without throwing, check whether cancellation was requested.
            if (token.IsCancellationRequested)
            {
                Console.WriteLine("Cancellation was requested, but layout generation completed (no built‑in cancellation support).");
            }
            else
            {
                Console.WriteLine("Layout generation completed without cancellation request.");
            }
        }
        catch (AggregateException ae)
        {
            // If the task was cancelled before it started, an OperationCanceledException may be thrown.
            if (ae.InnerExceptions.Count == 1 && ae.InnerException is OperationCanceledException)
            {
                Console.WriteLine("Layout generation was cancelled before it started.");
            }
            else
            {
                // Re‑throw unexpected exceptions.
                throw;
            }
        }

        // Verify that the document can still be saved after the layout task.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new Exception("Failed to save the document after layout task.");

        Console.WriteLine("Document saved successfully after layout test.");
    }
}
