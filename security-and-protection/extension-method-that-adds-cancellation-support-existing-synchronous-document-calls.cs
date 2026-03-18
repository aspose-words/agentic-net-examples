using System;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

namespace Aspose.Words.Extensions
{
    /// <summary>
    /// Provides extension methods for <see cref="Document"/> that add cancellation support to synchronous save operations.
    /// </summary>
    public static class DocumentExtensions
    {
        /// <summary>
        /// Saves the document to a file using the specified <see cref="SaveOptions"/> and observes the supplied <see cref="CancellationToken"/>.
        /// If cancellation is requested, an <see cref="OperationCanceledException"/> is thrown from the saving callback.
        /// </summary>
        /// <param name="document">The document to save.</param>
        /// <param name="fileName">The target file name.</param>
        /// <param name="saveOptions">The save options that control how the document is saved.</param>
        /// <param name="cancellationToken">The token to monitor for cancellation requests.</param>
        public static void Save(this Document document, string fileName, SaveOptions saveOptions, CancellationToken cancellationToken)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (fileName == null) throw new ArgumentNullException(nameof(fileName));
            if (saveOptions == null) throw new ArgumentNullException(nameof(saveOptions));

            // Attach a progress callback that checks the cancellation token.
            saveOptions.ProgressCallback = new CancellationSavingCallback(cancellationToken);

            // Perform the regular synchronous save. The callback will abort if cancellation occurs.
            document.Save(fileName, saveOptions);
        }

        /// <summary>
        /// Saves the document to a file using default save options for the file extension and observes the supplied <see cref="CancellationToken"/>.
        /// </summary>
        /// <param name="document">The document to save.</param>
        /// <param name="fileName">The target file name.</param>
        /// <param name="cancellationToken">The token to monitor for cancellation requests.</param>
        public static void Save(this Document document, string fileName, CancellationToken cancellationToken)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (fileName == null) throw new ArgumentNullException(nameof(fileName));

            // Create appropriate SaveOptions based on the file extension.
            SaveOptions options = SaveOptions.CreateSaveOptions(fileName);
            document.Save(fileName, options, cancellationToken);
        }

        /// <summary>
        /// Implementation of <see cref="IDocumentSavingCallback"/> that throws <see cref="OperationCanceledException"/>
        /// when the associated <see cref="CancellationToken"/> is signaled.
        /// </summary>
        private sealed class CancellationSavingCallback : IDocumentSavingCallback
        {
            private readonly CancellationToken _cancellationToken;

            public CancellationSavingCallback(CancellationToken cancellationToken)
            {
                _cancellationToken = cancellationToken;
            }

            public void Notify(DocumentSavingArgs args)
            {
                // The saving process periodically invokes this method.
                // If cancellation has been requested, abort the operation.
                if (_cancellationToken.IsCancellationRequested)
                {
                    // Include progress information for diagnostic purposes.
                    throw new OperationCanceledException(
                        $"Saving was canceled. EstimatedProgress = {args.EstimatedProgress}");
                }
            }
        }
    }

    // ---------------------------------------------------------------------
    // Minimal console entry point required for a .NET executable project.
    // ---------------------------------------------------------------------
    internal class Program
    {
        static void Main(string[] args)
        {
            // Example usage of the cancellation‑aware Save extension.
            // In a real scenario the token could be linked to UI cancellation, timeout, etc.
            var cancellationSource = new CancellationTokenSource();
            var token = cancellationSource.Token;

            // Create a simple document.
            Document doc = new Document();
            doc.RemoveAllChildren(); // keep it empty for the demo.

            // Optionally trigger cancellation after a short delay to demonstrate behaviour.
            // Here we do not cancel, but the code path is ready.
            // cancellationSource.CancelAfter(TimeSpan.FromMilliseconds(10));

            try
            {
                // Save using the extension that observes the token.
                doc.Save("output.docx", token);
                Console.WriteLine("Document saved successfully.");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Save operation was canceled: {ex.Message}");
            }
        }
    }
}
