using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace DocumentProcessingInterruptionDemo
{
    // Callback that checks a CancellationToken during document loading.
    public class CancellationTokenLoadingCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public CancellationTokenLoadingCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        // Called periodically while the document is being loaded.
        public void Notify(DocumentLoadingArgs args)
        {
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException(
                    $"Loading canceled. EstimatedProgress = {args.EstimatedProgress}");
        }
    }

    // Callback that checks a CancellationToken during document saving.
    public class CancellationTokenSavingCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public CancellationTokenSavingCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        // Called periodically while the document is being saved.
        public void Notify(DocumentSavingArgs args)
        {
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException(
                    $"Saving canceled. EstimatedProgress = {args.EstimatedProgress}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            const string inputPath = @"C:\Docs\BigDocument.docx";
            // Path where the processed document will be saved.
            const string outputPath = @"C:\Docs\ProcessedDocument.pdf";

            // Create a CancellationTokenSource that will request cancellation after 2 seconds.
            using var cts = new CancellationTokenSource();
            cts.CancelAfter(TimeSpan.FromSeconds(2));

            // Set up loading options with the custom loading callback.
            var loadOptions = new LoadOptions
            {
                ProgressCallback = new CancellationTokenLoadingCallback(cts.Token)
            };

            try
            {
                // Load the document using the load options.
                var doc = new Document(inputPath, loadOptions);

                // Optionally perform additional processing on the document here.

                // Set up saving options with the custom saving callback.
                var saveOptions = new PdfSaveOptions
                {
                    ProgressCallback = new CancellationTokenSavingCallback(cts.Token)
                };

                // Save the document; the callback will monitor cancellation.
                doc.Save(outputPath, saveOptions);
            }
            catch (OperationCanceledException ex)
            {
                // Handle the cancellation gracefully.
                Console.WriteLine($"Operation was canceled: {ex.Message}");
            }
            catch (FileCorruptedException ex)
            {
                // Handle a corrupted source file.
                Console.WriteLine($"The source file is corrupted: {ex.Message}");
            }
            catch (UnsupportedFileFormatException ex)
            {
                // Handle an unsupported file format.
                Console.WriteLine($"Unsupported file format: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors.
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
