using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeCancellationDemo
{
    // Callback that aborts the save operation after a short time.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _startTime;
        // Maximum allowed duration in seconds before cancellation.
        private const double MaxDuration = 0.01; // 10 ms

        public SavingProgressCallback()
        {
            _startTime = DateTime.Now;
        }

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _startTime).TotalSeconds;
            if (elapsed > MaxDuration)
                throw new OperationCanceledException(
                    $"Saving cancelled after {elapsed:F3}s (estimated progress = {args.EstimatedProgress}).");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a folder for output files.
            string artifactsDir = Path.Combine(Path.GetTempPath(), "AsposeCancellationDemo");
            Directory.CreateDirectory(artifactsDir);

            // Create a sample document with enough content to make saving take noticeable time.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            for (int i = 0; i < 5000; i++)
                builder.Writeln($"Line {i + 1}");

            // Configure save options with the progress‑cancellation callback.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback()
            };

            string outPath = Path.Combine(artifactsDir, "Canceled.docx");
            bool cancelled = false;

            try
            {
                // Attempt to save; the callback should abort the operation.
                doc.Save(outPath, saveOptions);
            }
            catch (OperationCanceledException)
            {
                cancelled = true;
            }

            // Verify that cancellation was observed.
            if (!cancelled)
                throw new Exception("The document save operation was not cancelled as expected.");

            // The file should not exist or be incomplete; delete it if present.
            if (File.Exists(outPath))
                File.Delete(outPath);

            // Indicate success (no interactive input required).
            Console.WriteLine("Cancellation verified successfully.");
        }
    }
}
