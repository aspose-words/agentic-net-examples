using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Callback that aborts the save operation after a short time.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _startTime;
        private const double MaxDurationSeconds = 0.001; // Very short to trigger cancellation quickly.

        public SavingProgressCallback()
        {
            _startTime = DateTime.Now;
        }

        public void Notify(DocumentSavingArgs args)
        {
            // If the elapsed time exceeds the limit, abort the operation.
            if ((DateTime.Now - _startTime).TotalSeconds > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }

    public static void Main()
    {
        // Prepare a temporary output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Canceled.docx");
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        // Create a document with enough content to make layout processing noticeable.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 200; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure save options with the progress callback that will cancel the operation.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback()
        };

        bool cancellationCaught = false;
        try
        {
            // Attempt to save; the callback should abort the process.
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            // Expected cancellation.
            Console.WriteLine($"Save operation canceled as expected: {ex.Message}");
            cancellationCaught = true;
        }

        // Validate that cancellation was detected.
        if (!cancellationCaught)
            throw new Exception("The save operation was not canceled as expected.");

        // Ensure the partially saved file does not exist (Aspose.Words cleans up on cancellation).
        if (File.Exists(outputPath))
            throw new Exception("Output file should not exist after cancellation.");

        Console.WriteLine("Cancellation behavior test completed successfully.");
    }
}
