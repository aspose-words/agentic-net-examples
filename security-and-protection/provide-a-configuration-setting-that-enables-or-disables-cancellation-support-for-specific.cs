using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Simple configuration class to enable or disable cancellation for specific stages.
    public class ProcessingConfig
    {
        // When true, the saving process will be cancelled once progress exceeds the threshold.
        public bool CancelDuringSaving { get; set; } = false;
    }

    // Implements the saving progress callback. Throws an exception to cancel saving
    // when the configuration requests cancellation and the progress threshold is met.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly ProcessingConfig _config;
        private const double CancelThreshold = 0.5; // Cancel after 50% progress.

        public SavingProgressCallback(ProcessingConfig config)
        {
            _config = config;
        }

        public void Notify(DocumentSavingArgs args)
        {
            if (_config.CancelDuringSaving && args.EstimatedProgress >= CancelThreshold)
                throw new OperationCanceledException($"Saving canceled at {args.EstimatedProgress:P0} progress.");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a simple document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world! This is a test document for cancellation support.");

            // Define output paths.
            string baseDir = Directory.GetCurrentDirectory();
            string cancelPath = Path.Combine(baseDir, "CancelledOutput.docx");
            string successPath = Path.Combine(baseDir, "SuccessfulOutput.docx");

            // ---------- Demonstration with cancellation enabled ----------
            var config = new ProcessingConfig { CancelDuringSaving = true };
            var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback(config)
            };

            try
            {
                doc.Save(cancelPath, saveOptions);
                Console.WriteLine("Document saved (cancellation was expected but did not occur).");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Save operation was cancelled as configured: {ex.Message}");
            }

            // ---------- Demonstration with cancellation disabled ----------
            config.CancelDuringSaving = false; // Disable cancellation.
            // Reuse the same document; optionally modify it.
            builder.Writeln("Additional line after cancellation test.");

            var saveOptionsNoCancel = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback(config)
            };

            try
            {
                doc.Save(successPath, saveOptionsNoCancel);
                Console.WriteLine($"Document saved successfully to '{successPath}'.");
            }
            catch (OperationCanceledException ex)
            {
                // This block should not be reached when cancellation is disabled.
                Console.WriteLine($"Unexpected cancellation: {ex.Message}");
            }

            // Verify that the output files exist.
            if (File.Exists(cancelPath))
                Console.WriteLine($"Cancelled file exists (partial save may be present): {cancelPath}");
            else
                Console.WriteLine("Cancelled file was not created, as expected.");

            if (File.Exists(successPath))
                Console.WriteLine($"Successful file exists: {successPath}");
            else
                Console.WriteLine("Successful file was not created, which indicates an error.");
        }
    }
}
