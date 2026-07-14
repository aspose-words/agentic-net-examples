using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Entry point.
    public static void Main()
    {
        // Path for the output document.
        const string outputPath = "CanceledOutput.docx";

        // Ensure a clean start.
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document is used to test cancellation of the saving pipeline.");

        // Attempt to save with a progress callback that cancels immediately.
        bool cancellationCaught = false;
        try
        {
            OoxmlSaveOptions cancelOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback()
            };
            doc.Save(outputPath, cancelOptions);
        }
        catch (OperationCanceledException)
        {
            cancellationCaught = true;
            Console.WriteLine("Saving was cancelled as expected.");
        }

        // Verify that cancellation was detected.
        if (!cancellationCaught)
            throw new Exception("Expected cancellation was not triggered.");

        // Verify that the partially saved file does not exist (or is empty).
        if (File.Exists(outputPath) && new FileInfo(outputPath).Length > 0)
            throw new Exception("File should not exist or should be empty after cancellation.");

        // Now save without cancellation to confirm normal pipeline works.
        const string successPath = "SuccessfulOutput.docx";
        if (File.Exists(successPath))
            File.Delete(successPath);

        OoxmlSaveOptions normalOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        doc.Save(successPath, normalOptions);
        Console.WriteLine("Document saved successfully without cancellation.");

        // Load the saved document and verify its content.
        Document loaded = new Document(successPath);
        string text = loaded.GetText().Trim();
        if (!text.Contains("This document is used to test cancellation"))
            throw new Exception("Saved document content verification failed.");

        Console.WriteLine("Loaded document content verified successfully.");
    }

    // Implementation of the progress callback that aborts the save operation.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _startTime;
        // Set to zero to cancel on the first progress notification.
        private const double MaxDurationSeconds = 0.0;

        public SavingProgressCallback()
        {
            _startTime = DateTime.Now;
        }

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _startTime).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; Cancellation triggered after {elapsed} seconds.");
        }
    }
}
