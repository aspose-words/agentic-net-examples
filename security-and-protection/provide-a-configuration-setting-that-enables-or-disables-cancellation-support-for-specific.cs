using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Callback that can cancel the saving process based on progress.
    public class CancelSavingCallback : IDocumentSavingCallback
    {
        private readonly double _cancelAfterProgress;

        public CancelSavingCallback(double cancelAfterProgress = 0.5)
        {
            _cancelAfterProgress = cancelAfterProgress;
        }

        public void Notify(DocumentSavingArgs args)
        {
            // Cancel when the estimated progress reaches the configured threshold.
            if (args.EstimatedProgress >= _cancelAfterProgress)
                throw new OperationCanceledException(
                    $"Saving canceled at {args.EstimatedProgress:P0} progress.");
        }
    }

    public class Program
    {
        // Configuration setting: enable or disable cancellation support.
        private static readonly bool EnableCancellation = true;

        public static void Main()
        {
            // Prepare a simple document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document used to demonstrate cancellation support.");

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");

            try
            {
                if (EnableCancellation)
                {
                    // Attach the cancellation callback via SaveOptions.
                    OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
                    {
                        ProgressCallback = new CancelSavingCallback()
                    };
                    doc.Save(outputPath, saveOptions);
                }
                else
                {
                    // Save without any cancellation logic.
                    doc.Save(outputPath);
                }

                // Verify that the file was created.
                if (File.Exists(outputPath))
                    Console.WriteLine($"Document saved successfully to '{outputPath}'.");
                else
                    throw new InvalidOperationException("The document was not saved as expected.");
            }
            catch (OperationCanceledException ex)
            {
                // Expected when cancellation is enabled and the callback aborts the save.
                Console.WriteLine($"Saving was canceled: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Any other unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
                throw;
            }
        }
    }
}
