using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeCancellationDemo
{
    // Callback that aborts the saving process after a short time.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _savingStartedAt;
        private const double MaxDurationSeconds = 0.01; // Cancel quickly.

        public SavingProgressCallback()
        {
            _savingStartedAt = DateTime.Now;
        }

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _savingStartedAt).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"Saving canceled after {elapsed:F3}s (EstimatedProgress = {args.EstimatedProgress}).");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string outputDir = Path.Combine(Path.GetTempPath(), "AsposeCancellationDemo");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "CanceledDocument.docx");

            // Ensure any previous file is removed.
            if (File.Exists(outputPath))
                File.Delete(outputPath);

            // Create a simple document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This document will be saved with a progress callback that cancels the operation.");

            // Configure save options with the cancellation callback.
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
                // Expected path: saving was canceled.
                Console.WriteLine($"Cancellation confirmed: {ex.Message}");
                cancellationCaught = true;
            }

            // Verify that cancellation was observed.
            if (!cancellationCaught)
                throw new Exception("The saving operation was not canceled as expected.");

            // Verify that the output file was not created (or is empty).
            if (File.Exists(outputPath) && new FileInfo(outputPath).Length > 0)
                throw new Exception("Output file exists despite cancellation.");

            Console.WriteLine("Document processing pipeline correctly respected cancellation.");
        }
    }
}
