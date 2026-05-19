using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Callback that aborts the save operation after a short duration.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly Stopwatch _timer = Stopwatch.StartNew();
        private const double MaxDurationSeconds = 0.001; // Abort almost immediately.

        public void Notify(DocumentSavingArgs args)
        {
            if (_timer.Elapsed.TotalSeconds > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }

    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "CancelledDocument.docx");

        // Ensure any previous file is removed.
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        // Build a document with many paragraphs to make layout take noticeable time.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 2000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Set up save options with the progress callback that will cancel the operation.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback()
        };

        // Attempt to save and expect an OperationCanceledException.
        bool canceled = false;
        try
        {
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException)
        {
            canceled = true;
        }

        // Validate that cancellation occurred.
        if (!canceled)
            throw new Exception("The save operation was not canceled as expected.");

        // Validate that the partially saved file does not exist.
        if (File.Exists(outputPath))
            throw new Exception("The output file should not exist after cancellation.");

        // If we reach this point, the test succeeded.
        // No interactive output is required.
    }
}
