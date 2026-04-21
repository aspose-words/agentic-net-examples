using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");

        // Save the initial document.
        doc.Save(outputPath);

        // Define a timeout duration.
        TimeSpan timeout = TimeSpan.FromSeconds(2);

        // Set up a CancellationTokenSource that will be cancelled after the timeout.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Schedule cancellation.
            Task.Delay(timeout).ContinueWith(_ => cts.Cancel());

            try
            {
                // Perform a simulated long-running operation that respects the token.
                PerformLongOperation(cts.Token);

                // If the operation completes without cancellation, add a note.
                builder.Writeln("Operation completed without cancellation.");
            }
            catch (OperationCanceledException)
            {
                // If cancelled, add a note indicating the timeout.
                builder.Writeln("Operation was cancelled due to timeout.");
            }

            // Save the final document.
            doc.Save(outputPath);
        }

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }

    // Simulates work that periodically checks for cancellation.
    private static void PerformLongOperation(CancellationToken token)
    {
        // Loop longer than the timeout to ensure cancellation occurs.
        for (int i = 0; i < 10; i++)
        {
            token.ThrowIfCancellationRequested();
            Thread.Sleep(500); // Simulate work (total ~5 seconds).
        }
    }
}
