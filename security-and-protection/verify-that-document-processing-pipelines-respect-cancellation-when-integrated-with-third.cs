using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    // Simulates a third‑party reporting tool that processes each page of a document.
    private static async Task ProcessDocumentAsync(Document doc, CancellationToken token)
    {
        // Ensure layout is up to date so PageCount is accurate.
        doc.UpdatePageLayout();
        int pageCount = doc.PageCount;

        for (int i = 0; i < pageCount; i++)
        {
            // Respect cancellation request.
            token.ThrowIfCancellationRequested();

            // Simulate time‑consuming work for each page.
            await Task.Delay(50, token);
        }
    }

    public static void Main()
    {
        // Prepare a sample document with many pages.
        string sourcePath = "source.docx";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert enough paragraphs to generate multiple pages.
        for (int i = 0; i < 200; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
            // Insert a page break every 5 paragraphs to increase page count.
            if ((i + 1) % 5 == 0)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the source document locally.
        sourceDoc.Save(sourcePath);

        // Validate that the source file was created.
        if (!File.Exists(sourcePath))
            throw new Exception($"Failed to create source document at '{sourcePath}'.");

        // Load the document for processing.
        Document docToProcess = new Document(sourcePath);

        // Set up cancellation: cancel after a short delay.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Start the processing task.
            Task processingTask = ProcessDocumentAsync(docToProcess, cts.Token);

            // Cancel after 200 milliseconds.
            Task.Delay(200).ContinueWith(_ => cts.Cancel());

            try
            {
                // Wait for the processing to finish (or be cancelled).
                processingTask.Wait();

                // If we reach here, processing completed without cancellation, which is unexpected.
                throw new Exception("Processing completed without respecting cancellation.");
            }
            catch (AggregateException ae)
            {
                // Expect an OperationCanceledException.
                if (!ae.InnerExceptions.Any(e => e is OperationCanceledException))
                    throw; // Re‑throw unexpected exceptions.
                // Cancellation was respected – continue.
            }
        }

        // Optionally, save the (partially) processed document to demonstrate normal workflow.
        string outputPath = "processed.docx";
        docToProcess.Save(outputPath);

        // Validate that the output file exists.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to save processed document at '{outputPath}'.");

        // Indicate successful completion.
        Console.WriteLine("Cancellation was respected and output document saved.");
    }
}
