using System;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputPath);
        string docPath = Path.Combine(outputPath, "InterruptedDocument.docx");

        // Create a blank document.
        Document doc = new Document();

        // Set up a cancellation token source that will be used to interrupt the building process.
        CancellationTokenSource cts = new CancellationTokenSource();

        // Create a DocumentBuilder (no interruption options are needed because we will check the token manually).
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the document in a loop.
        for (int i = 1; i <= 10; i++)
        {
            // Simulate a condition that triggers cancellation after a few paragraphs.
            if (i == 5)
                cts.Cancel(); // Request interruption.

            // Check for cancellation before performing the builder operation.
            if (cts.Token.IsCancellationRequested)
            {
                Console.WriteLine($"Document building was interrupted at paragraph {i}.");
                break;
            }

            try
            {
                // Write a paragraph.
                builder.Writeln($"Paragraph {i}");
            }
            catch (OperationCanceledException)
            {
                // This catch is retained for completeness, although the manual check prevents the exception.
                Console.WriteLine($"Document building was interrupted at paragraph {i}.");
                break;
            }
        }

        // Save the (potentially partially) built document.
        doc.Save(docPath);

        // Validate that the file was created.
        if (!File.Exists(docPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        Console.WriteLine($"Document saved to: {docPath}");
    }
}
