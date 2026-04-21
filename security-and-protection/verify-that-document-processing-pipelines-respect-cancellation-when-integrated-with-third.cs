using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Callback that aborts the saving process as soon as it is invoked.
    class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _startTime;
        // Set to zero so that any progress notification triggers cancellation.
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
                    $"EstimatedProgress = {args.EstimatedProgress}; Canceled after {elapsed:F4}s");
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare a temporary file path for the output document.
            string outputPath = Path.Combine(Path.GetTempPath(), "CancellationDemo.docx");

            // Ensure any previous file is removed.
            if (File.Exists(outputPath))
                File.Delete(outputPath);

            // Create a sample document with enough content to make saving take measurable time.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            for (int i = 0; i < 2000; i++)
            {
                builder.Writeln($"Paragraph {i + 1}");
            }

            // Configure save options with the cancellation callback.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback()
            };

            // Attempt to save the document and expect cancellation.
            bool canceled = false;
            try
            {
                doc.Save(outputPath, saveOptions);
            }
            catch (OperationCanceledException ex)
            {
                canceled = true;
                Console.WriteLine($"Saving was canceled as expected: {ex.Message}");
            }

            // Verify that cancellation was reported and no file was created.
            if (!canceled)
                throw new Exception("The saving operation was not canceled as expected.");

            if (File.Exists(outputPath))
                throw new Exception("Output file should not exist after cancellation.");

            Console.WriteLine("Cancellation behavior verified successfully.");
        }
    }
}
