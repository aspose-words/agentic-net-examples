using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeCancellationDemo
{
    // Simple configuration class to enable or disable cancellation.
    public class Config
    {
        public bool EnableCancellation { get; set; }
    }

    // Callback that checks the configuration and throws if cancellation is enabled.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly Config _config;

        public SavingProgressCallback(Config config)
        {
            _config = config;
        }

        public void Notify(DocumentSavingArgs args)
        {
            // If cancellation is enabled, abort the save operation.
            if (_config.EnableCancellation)
                throw new OperationCanceledException(
                    $"Saving canceled at estimated progress {args.EstimatedProgress}%.");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a temporary folder for output files.
            string outputDir = Path.Combine(Path.GetTempPath(), "AsposeCancellationDemo");
            Directory.CreateDirectory(outputDir);

            // Create a simple document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document used to demonstrate cancellation support.");

            // Configuration instance.
            Config config = new Config();

            // ---------- Example 1: Cancellation enabled ----------
            config.EnableCancellation = true;
            string canceledPath = Path.Combine(outputDir, "CanceledOutput.pdf");
            PdfSaveOptions cancelOptions = new PdfSaveOptions
            {
                ProgressCallback = new SavingProgressCallback(config)
            };

            try
            {
                doc.Save(canceledPath, cancelOptions);
                Console.WriteLine("Document saved (cancellation disabled).");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Save operation was canceled as expected: {ex.Message}");
            }

            // Verify that the file was not created.
            if (!File.Exists(canceledPath))
                Console.WriteLine("No file was created due to cancellation.");

            // ---------- Example 2: Cancellation disabled ----------
            config.EnableCancellation = false;
            string savedPath = Path.Combine(outputDir, "SavedOutput.pdf");
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                ProgressCallback = new SavingProgressCallback(config)
            };

            try
            {
                doc.Save(savedPath, saveOptions);
                Console.WriteLine("Document saved successfully.");
            }
            catch (OperationCanceledException ex)
            {
                // This block should not be reached in this example.
                Console.WriteLine($"Unexpected cancellation: {ex.Message}");
            }

            // Validate that the file exists.
            if (File.Exists(savedPath))
                Console.WriteLine($"Output file exists: {savedPath}");
            else
                Console.WriteLine("Failed to create the output file.");

            // Clean up (optional).
            // Directory.Delete(outputDir, true);
        }
    }
}
