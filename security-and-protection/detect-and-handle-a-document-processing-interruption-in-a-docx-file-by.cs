using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

namespace AsposeWordsCancellationDemo
{
    // Callback that checks a CancellationToken during document saving.
    public class CancellationSavingCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public CancellationSavingCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        // Called repeatedly while the document is being saved.
        public void Notify(DocumentSavingArgs args)
        {
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException(
                    $"Saving cancelled. EstimatedProgress = {args.EstimatedProgress}");
        }
    }

    // Callback that checks a CancellationToken during document loading.
    public class CancellationLoadingCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public CancellationLoadingCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        // Called repeatedly while the document is being loaded.
        public void Notify(DocumentLoadingArgs args)
        {
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException(
                    $"Loading cancelled. EstimatedProgress = {args.EstimatedProgress}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            const string inputPath = @"C:\Docs\Input.docx";
            // Path where the processed DOCX will be saved.
            const string outputPath = @"C:\Docs\Output.docx";

            // Create a CancellationTokenSource that will request cancellation after a short delay.
            using var cts = new CancellationTokenSource();
            // For demonstration, cancel after 200 milliseconds.
            cts.CancelAfter(200);
            CancellationToken token = cts.Token;

            try
            {
                // Load the document with a loading progress callback that respects the token.
                LoadOptions loadOptions = new LoadOptions
                {
                    ProgressCallback = new CancellationLoadingCallback(token)
                };
                Document doc = new Document(inputPath, loadOptions);

                // Perform any processing on the document here.
                // (No custom processing required for this demo.)

                // Prepare save options with a saving progress callback that respects the token.
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
                {
                    ProgressCallback = new CancellationSavingCallback(token)
                };

                // Save the document; the callbacks will throw if cancellation is requested.
                doc.Save(outputPath, saveOptions);
                Console.WriteLine("Document saved successfully.");
            }
            catch (OperationCanceledException ex)
            {
                // Handle the interruption gracefully.
                Console.WriteLine($"Operation was cancelled: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle other possible exceptions (e.g., file not found, corrupted file).
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
