using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Configure load options with a progress callback.
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback()
        };

        try
        {
            // Load the document using the configured options.
            Document doc = new Document("Big document.docx", loadOptions);
            Console.WriteLine("Document loaded successfully.");
        }
        catch (OperationCanceledException ex)
        {
            // Handle cancellation triggered by the callback.
            Console.WriteLine($"Loading canceled: {ex.Message}");
        }
    }

    // User‑defined callback that aborts loading after a time limit.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _loadingStartedAt;
        private const double MaxDurationSeconds = 0.5; // Maximum allowed loading time.

        public LoadingProgressCallback()
        {
            _loadingStartedAt = DateTime.Now;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            double elapsed = (DateTime.Now - _loadingStartedAt).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }
}
