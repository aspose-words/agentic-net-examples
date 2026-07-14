using System;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare cancellation token that will be triggered after a short delay.
        var cts = new CancellationTokenSource();
        // Cancel after 100 milliseconds.
        var timer = new Timer(_ => cts.Cancel(), null, 100, Timeout.Infinite);

        // Create a new empty document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        try
        {
            // Simulate complex document construction with many steps.
            for (int i = 1; i <= 20; i++)
            {
                // Check for cancellation before each operation.
                if (cts.Token.IsCancellationRequested)
                    throw new OperationCanceledException();

                builder.Writeln($"Paragraph {i}");
                // Simulate work.
                Thread.Sleep(30);
            }
        }
        catch (OperationCanceledException)
        {
            // Interruption occurred – continue with whatever has been built so far.
            Console.WriteLine("Document construction was interrupted.");
        }
        finally
        {
            // Ensure timer resources are released.
            timer.Dispose();
        }

        // Save the (potentially partially built) document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
