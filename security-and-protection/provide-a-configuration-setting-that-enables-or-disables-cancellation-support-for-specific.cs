using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Callback that can cancel a save operation based on a configuration flag.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly bool _enableCancellation;
        private readonly double _cancelAfterProgress;

        public SavingProgressCallback(bool enableCancellation, double cancelAfterProgress = 0.5)
        {
            _enableCancellation = enableCancellation;
            _cancelAfterProgress = cancelAfterProgress;
        }

        public void Notify(DocumentSavingArgs args)
        {
            // If cancellation is enabled and the estimated progress exceeds the threshold,
            // abort the save by throwing an exception.
            if (_enableCancellation && args.EstimatedProgress >= _cancelAfterProgress)
                throw new OperationCanceledException(
                    $"Saving canceled at {args.EstimatedProgress:P0} progress.");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Configuration setting: turn cancellation on or off for the PDF save stage.
            bool cancelDuringPdfSave = true;

            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string pdfPath = Path.Combine(outputDir, "Sample.pdf");

            // Create a simple document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello Aspose.Words cancellation demo.");

            // Attempt to save with cancellation enabled.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ProgressCallback = new SavingProgressCallback(cancelDuringPdfSave, 0.1)
            };

            try
            {
                doc.Save(pdfPath, pdfOptions);
                Console.WriteLine("PDF saved successfully (cancellation disabled).");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Save operation canceled: {ex.Message}");
                // Ensure no partial file remains.
                if (File.Exists(pdfPath))
                    File.Delete(pdfPath);
            }

            // Disable cancellation and save again.
            cancelDuringPdfSave = false;
            pdfOptions.ProgressCallback = new SavingProgressCallback(cancelDuringPdfSave);
            doc.Save(pdfPath, pdfOptions);
            Console.WriteLine("PDF saved without cancellation.");

            // Validate that the file was created.
            if (!File.Exists(pdfPath))
                throw new Exception("The PDF file was not created as expected.");
        }
    }
}
