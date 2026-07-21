using System;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a cancellation token that will be cancelled after a few operations.
        using var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Build the document in a separate method, passing the token.
        Document doc = BuildDocument(token, cts);

        // Save the document.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        // Indicate success (no interactive output required, but console write is allowed).
        Console.WriteLine("Document built and saved successfully.");
    }

    private static Document BuildDocument(CancellationToken token, CancellationTokenSource cts)
    {
        // Create a new empty document.
        Document document = new Document();
        DocumentBuilder builder = new DocumentBuilder(document);

        // Simulate a complex document construction with many sections.
        for (int i = 1; i <= 20; i++)
        {
            // Check for cancellation request.
            if (token.IsCancellationRequested)
            {
                Console.WriteLine($"Construction interrupted after {i - 1} sections.");
                break;
            }

            // Add a heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Section {i}");

            // Add some body text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            // Simulate a condition to cancel after 5 sections.
            if (i == 5)
            {
                // Trigger cancellation.
                cts.Cancel();
            }
        }

        return document;
    }
}
