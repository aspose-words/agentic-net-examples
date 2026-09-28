using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Tables; // For NodeType enum

public class Program
{
    // Entry point of the console application.
    public static async Task Main(string[] args)
    {
        // Paths for temporary source and output documents.
        string sourcePath = Path.Combine(Path.GetTempPath(), "sample.docx");
        string outputPath = Path.Combine(Path.GetTempPath(), "processed.pdf");

        // Ensure any previous files are removed.
        if (File.Exists(sourcePath)) File.Delete(sourcePath);
        if (File.Exists(outputPath)) File.Delete(outputPath);

        // 1. Create a sample document with many paragraphs to simulate a long‑running operation.
        Document sourceDoc = new Document();
        for (int i = 0; i < 500; i++)
        {
            sourceDoc.FirstSection.Body.AppendChild(new Paragraph(sourceDoc));
            sourceDoc.FirstSection.Body.LastParagraph.AppendChild(new Run(sourceDoc, $"Paragraph {i + 1}"));
        }
        sourceDoc.Save(sourcePath); // Save the source document.

        // 2. Load the document (bootstrap step as required by the rules).
        Document doc = new Document(sourcePath);

        // 3. Set up cancellation.
        using CancellationTokenSource cts = new CancellationTokenSource();

        // Start the processing pipeline.
        Task processingTask = ProcessDocumentAsync(doc, outputPath, cts.Token);

        // Cancel after a short delay to simulate user‑initiated cancellation.
        _ = Task.Delay(100).ContinueWith(_ => cts.Cancel());

        try
        {
            await processingTask;
            // If we reach here, processing completed without cancellation – this is unexpected for the test.
            throw new Exception("Processing completed despite cancellation request.");
        }
        catch (OperationCanceledException)
        {
            // Expected path: processing was cancelled.
        }

        // 4. Verify that the output file was not created due to cancellation.
        if (File.Exists(outputPath))
        {
            throw new Exception("Output file was created even though processing was cancelled.");
        }

        // Clean up source file.
        if (File.Exists(sourcePath)) File.Delete(sourcePath);
    }

    // Simulated document processing pipeline that respects cancellation.
    private static async Task ProcessDocumentAsync(Document doc, string outputPath, CancellationToken token)
    {
        // Simulate work per paragraph; check cancellation token regularly.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Simulate a small amount of work.
            await Task.Delay(5, token);
            // Throw if cancellation was requested.
            token.ThrowIfCancellationRequested();
        }

        // If not cancelled, save the document to PDF.
        doc.Save(outputPath);
    }
}
