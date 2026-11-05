using System;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a cancellation token source that will be used to interrupt the building process.
        CancellationTokenSource cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        try
        {
            // Build a complex document with many paragraphs.
            for (int i = 1; i <= 1000; i++)
            {
                // Simulate a condition to cancel the operation (e.g., after 500 paragraphs).
                if (i == 501)
                {
                    cts.Cancel(); // Request cancellation.
                }

                // Throw if cancellation has been requested.
                token.ThrowIfCancellationRequested();

                // Insert a paragraph using DocumentBuilder.
                builder.Writeln($"Paragraph {i}");
            }
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document building was interrupted by the cancellation token.");
        }

        // Save the partially built document.
        string outputPath = "ComplexDocument.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
