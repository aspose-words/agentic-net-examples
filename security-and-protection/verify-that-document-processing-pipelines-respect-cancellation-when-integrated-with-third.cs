using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    // Entry point of the console application.
    public static async Task Main(string[] args)
    {
        // Create a sample document with many paragraphs.
        const string sourcePath = "sample.docx";
        CreateSampleDocument(sourcePath);

        // Set up a cancellation token that will be triggered shortly.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(50); // Cancel after 50 milliseconds.

        try
        {
            // Process the document while respecting cancellation.
            await ProcessDocumentAsync(sourcePath, cts.Token);

            // If processing completes without cancellation, the test fails.
            throw new Exception("Cancellation was not respected by the processing pipeline.");
        }
        catch (OperationCanceledException)
        {
            // Expected path: cancellation was observed.
            Console.WriteLine("Cancellation respected.");
        }
        catch (Exception ex)
        {
            // Any other exception indicates a failure in the test.
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Creates a simple DOCX file containing a large number of paragraphs.
    private static void CreateSampleDocument(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        for (int i = 0; i < 200; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        doc.Save(path);
    }

    // Simulates a document processing pipeline that checks for cancellation between steps.
    private static async Task ProcessDocumentAsync(string path, CancellationToken token)
    {
        // Load the document (no password needed for this example).
        var loadOptions = new LoadOptions();
        var doc = new Document(path, loadOptions);

        // Iterate over each paragraph, simulating work and checking for cancellation.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            token.ThrowIfCancellationRequested();

            // Simulated work per paragraph (e.g., integration with a third‑party reporting tool).
            await Task.Delay(5, token);
        }

        // Save the processed document to verify that the pipeline can complete when not cancelled.
        const string outputPath = "processed.docx";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("Processed document was not saved as expected.");
        }
    }
}
