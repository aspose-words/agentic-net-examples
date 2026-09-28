using System;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a cancellation token that will be triggered after a short delay.
        using var cts = new CancellationTokenSource();
        // Cancel after 100 milliseconds to simulate an interruption.
        cts.CancelAfter(100);
        CancellationToken token = cts.Token;

        // Build the document with the ability to be interrupted.
        Document doc = BuildComplexDocument(token);

        // Save the resulting document (partial if cancelled).
        string outputPath = "ComplexDocument.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }

    private static Document BuildComplexDocument(CancellationToken token)
    {
        // Start with an empty document.
        Document document = new Document();
        DocumentBuilder builder = new DocumentBuilder(document);

        // Simulate building a complex document with many paragraphs.
        for (int i = 1; i <= 1000; i++)
        {
            // Check for cancellation before each operation.
            if (token.IsCancellationRequested)
                break; // Stop building and return the partially built document.

            builder.Writeln($"Paragraph {i}: This is a sample line of text.");
        }

        return document;
    }
}
