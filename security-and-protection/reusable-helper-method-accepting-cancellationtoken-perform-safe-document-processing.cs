using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsHelper
{
    /// <summary>
    /// Provides helper methods for safe document loading and saving with cancellation support.
    /// </summary>
    public static class DocumentHelper
    {
        /// <summary>
        /// Loads a document, optionally processes it, and saves it while respecting the supplied <see cref="CancellationToken"/>.
        /// </summary>
        /// <param name="inputPath">Path to the source document.</param>
        /// <param name="outputPath">Path where the processed document will be saved.</param>
        /// <param name="saveFormat">Desired format for the saved document.</param>
        /// <param name="cancellationToken">Token used to request cancellation.</param>
        /// <returns>True if the operation completed successfully; false if it was cancelled.</returns>
        public static bool ProcessDocument(string inputPath, string outputPath, SaveFormat saveFormat, CancellationToken cancellationToken)
        {
            // Prepare loading options with a progress callback that checks the cancellation token.
            var loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingCancellationCallback(cancellationToken)
            };

            Document doc;
            try
            {
                // Load the document using the provided load options.
                doc = new Document(inputPath, loadOptions);
            }
            catch (OperationCanceledException)
            {
                // Loading was cancelled.
                return false;
            }

            // Placeholder for any document manipulation logic.
            // For example: doc.Range.Replace("foo", "bar");

            // Create a concrete SaveOptions instance appropriate for the requested format.
            var saveOptions = SaveOptions.CreateSaveOptions(saveFormat);
            saveOptions.ProgressCallback = new SavingCancellationCallback(cancellationToken);

            try
            {
                // Save the document using the concrete SaveOptions instance.
                doc.Save(outputPath, saveOptions);
            }
            catch (OperationCanceledException)
            {
                // Saving was cancelled.
                return false;
            }

            // Operation completed without cancellation.
            return true;
        }

        /// <summary>
        /// Implements <see cref="IDocumentLoadingCallback"/> to monitor loading progress and abort if cancellation is requested.
        /// </summary>
        private class LoadingCancellationCallback : IDocumentLoadingCallback
        {
            private readonly CancellationToken _cancellationToken;

            public LoadingCancellationCallback(CancellationToken cancellationToken)
            {
                _cancellationToken = cancellationToken;
            }

            public void Notify(DocumentLoadingArgs args)
            {
                if (_cancellationToken.IsCancellationRequested)
                {
                    // Throwing an OperationCanceledException aborts the loading process.
                    throw new OperationCanceledException($"Loading cancelled at progress {args.EstimatedProgress}%.");
                }
            }
        }

        /// <summary>
        /// Implements <see cref="IDocumentSavingCallback"/> to monitor saving progress and abort if cancellation is requested.
        /// </summary>
        private class SavingCancellationCallback : IDocumentSavingCallback
        {
            private readonly CancellationToken _cancellationToken;

            public SavingCancellationCallback(CancellationToken cancellationToken)
            {
                _cancellationToken = cancellationToken;
            }

            public void Notify(DocumentSavingArgs args)
            {
                if (_cancellationToken.IsCancellationRequested)
                {
                    // Throwing an OperationCanceledException aborts the saving process.
                    throw new OperationCanceledException($"Saving cancelled at progress {args.EstimatedProgress}%.");
                }
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage of the helper method.
            string inputPath = "input.docx";   // replace with an existing file path
            string outputPath = "output.pdf"; // replace with desired output path
            var tokenSource = new CancellationTokenSource();

            bool success = DocumentHelper.ProcessDocument(
                inputPath,
                outputPath,
                SaveFormat.Pdf,
                tokenSource.Token);

            Console.WriteLine(success ? "Document processed successfully." : "Document processing was cancelled.");
        }
    }
}
