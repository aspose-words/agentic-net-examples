using System;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    // Builds a document while periodically checking a cancellation token.
    private static void BuildDocument(Document doc, CancellationToken token)
    {
        // DocumentBuilder resides in the Aspose.Words namespace.
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 1000; i++)
        {
            // Simulate a long‑running build operation.
            builder.Writeln($"Paragraph {i}");

            // Check for cancellation request and stop building if requested.
            if (token.IsCancellationRequested)
            {
                // Optionally add a note about interruption.
                builder.Writeln("[Document construction was interrupted]");
                break;
            }
        }
    }

    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "InterruptedDocument.docx");

        // Create a blank document.
        Document doc = new Document();

        // Set up a cancellation token that will be triggered after a short delay.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel after 10 milliseconds to simulate an early interruption.
            cts.CancelAfter(10);
            BuildDocument(doc, cts.Token);
        }

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Output the result path (no interactive prompts).
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
