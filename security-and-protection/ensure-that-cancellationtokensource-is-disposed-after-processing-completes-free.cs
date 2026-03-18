using System;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Create a CancellationTokenSource and ensure it is disposed after use.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Configure load options with a progress callback that can abort loading.
            LoadOptions loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingProgressCallback(cts.Token)
            };

            Document doc = null;
            try
            {
                // Load the document. The callback may throw OperationCanceledException.
                doc = new Document("input.docx", loadOptions);
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Loading canceled: {ex.Message}");
                return; // Exit if loading was aborted.
            }

            // ----- Document processing can be performed here -----
            // For example: doc.UpdateFields();

            // Save the processed document.
            doc.Save("output.docx");
        } // The CancellationTokenSource is disposed here, freeing system resources.
    }

    // Implements IDocumentLoadingCallback and respects a CancellationToken.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _token;
        private readonly DateTime _startTime = DateTime.Now;

        public LoadingProgressCallback(CancellationToken token)
        {
            _token = token;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            // Cancel if the external token has been signaled.
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException("Cancellation requested via token.");

            // Optional: cancel after a time limit (e.g., 30 seconds).
            if ((DateTime.Now - _startTime).TotalSeconds > 30)
                throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}");
        }
    }
}
