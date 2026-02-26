using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Implements the saving progress callback.
    // Throws OperationCanceledException when the supplied CancellationToken is cancelled.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public SavingProgressCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        public void Notify(DocumentSavingArgs args)
        {
            // If cancellation has been requested, abort the save operation.
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
            string sourcePath = @"C:\Docs\BigDocument.docx";

            // Path where the (partial) output would be written if not cancelled.
            string outputPath = @"C:\Docs\BigDocument_Cancelled.docx";

            // Create a CancellationTokenSource that will cancel after a short delay.
            using var cts = new CancellationTokenSource();

            // For demonstration, cancel after 100 milliseconds.
            cts.CancelAfter(TimeSpan.FromMilliseconds(100));

            // Load the document (no custom loading callback needed for this demo).
            Document doc = new Document(sourcePath);

            // Configure save options and attach the progress callback.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback(cts.Token)
            };

            try
            {
                // Attempt to save the document. The callback will abort if cancellation occurs.
                doc.Save(outputPath, saveOptions);
                Console.WriteLine("Document saved successfully (no cancellation).");
            }
            catch (OperationCanceledException ex)
            {
                // Handle the cancellation.
                Console.WriteLine($"Save operation was cancelled: {ex.Message}");
            }
        }
    }
}
