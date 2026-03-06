using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Implements the saving progress callback that checks a CancellationToken.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public SavingProgressCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        // This method is called periodically during document saving.
        public void Notify(DocumentSavingArgs args)
        {
            // If the token has been cancelled, abort the save operation.
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException(
                    $"Saving cancelled at estimated progress {args.EstimatedProgress}%.");
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            const string inputPath = @"C:\Docs\BigDocument.docx";
            // Path where the (partial) output would be written if not cancelled.
            const string outputPath = @"C:\Docs\BigDocument_Cancelled.docx";

            // Load the document using the standard constructor.
            Document doc = new Document(inputPath);

            // Create a CancellationTokenSource that will cancel after a short delay.
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                // Cancel after 200 milliseconds (adjust as needed for testing).
                cts.CancelAfter(200);

                // Configure save options and attach the progress callback.
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
                {
                    ProgressCallback = new SavingProgressCallback(cts.Token)
                };

                try
                {
                    // Attempt to save the document. The callback will throw if cancellation occurs.
                    doc.Save(outputPath, saveOptions);
                    Console.WriteLine("Document saved successfully.");
                }
                catch (OperationCanceledException ex)
                {
                    // Expected path when the operation is cancelled.
                    Console.WriteLine($"Save operation was cancelled: {ex.Message}");
                }
                catch (Exception ex)
                {
                    // Any other unexpected errors.
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
            }
        }
    }
}
