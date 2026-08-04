using System;
using System.IO;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Saving;

public class SavingProgressCallback : IDocumentSavingCallback
{
    private readonly Stopwatch _stopwatch = Stopwatch.StartNew();
    private const double MaxDurationSeconds = 0.01; // Abort quickly for the test

    public void Notify(DocumentSavingArgs args)
    {
        if (_stopwatch.Elapsed.TotalSeconds > MaxDurationSeconds)
            throw new OperationCanceledException(
                $"EstimatedProgress = {args.EstimatedProgress}; Canceled after {_stopwatch.Elapsed.TotalSeconds:F3}s");
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Path.GetTempPath(), "AsposeWordsCancellationTest");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "CancelledDocument.docx");

        // Create a large document to ensure layout takes noticeable time
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 2000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Configure save options with a progress callback that aborts after a short duration
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback()
        };

        bool cancellationOccurred = false;

        try
        {
            // Saving triggers layout generation; the callback will cancel the operation
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            cancellationOccurred = true;
            // Output the cancellation message (no interactive input required)
            Console.WriteLine($"Save operation was canceled: {ex.Message}");
        }

        // Validate that cancellation was detected
        if (!cancellationOccurred)
            throw new Exception("Expected the save operation to be canceled, but it completed successfully.");

        // Validate that the partially saved file does not exist
        if (File.Exists(outputPath))
            throw new Exception("The output file should not exist after a canceled save operation.");

        // Clean up temporary directory (optional)
        try { Directory.Delete(outputDir, true); } catch { /* ignore cleanup errors */ }
    }
}
