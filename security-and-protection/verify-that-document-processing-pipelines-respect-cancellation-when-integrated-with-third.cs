using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeCancellationDemo
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a temporary folder for the output file.
            string outputDir = Path.Combine(Path.GetTempPath(), "AsposeCancellationDemo");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "CanceledDocument.docx");

            // Create a simple document with some content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This document is used to test cancellation of the saving pipeline.");

            // Configure save options with a progress callback that aborts quickly.
            var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback()
            };

            bool canceled = false;
            try
            {
                // Attempt to save; the callback should throw OperationCanceledException.
                doc.Save(outputPath, saveOptions);
            }
            catch (OperationCanceledException)
            {
                canceled = true;
                Console.WriteLine("Saving was canceled as expected.");
            }

            // Verify that cancellation was observed.
            if (!canceled)
                throw new InvalidOperationException("The saving operation was not canceled.");

            // Clean up any partially written file.
            if (File.Exists(outputPath))
                File.Delete(outputPath);
        }

        // Implements the progress callback used by Aspose.Words during saving.
        private class SavingProgressCallback : IDocumentSavingCallback
        {
            private readonly DateTime _startTime;
            // Cancel after a very short duration (e.g., 0.001 seconds).
            private const double MaxDurationSeconds = 0.001;

            public SavingProgressCallback()
            {
                _startTime = DateTime.Now;
            }

            public void Notify(DocumentSavingArgs args)
            {
                double elapsed = (DateTime.Now - _startTime).TotalSeconds;
                if (elapsed > MaxDurationSeconds)
                    throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}");
            }
        }
    }
}
