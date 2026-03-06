using System;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the large DOCX file.
        string filePath = "BigDocument.docx";

        // Cancel the operation after 2 seconds (adjust as needed).
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));

        try
        {
            // Load the document with cancellation support.
            Document doc = LoadDocumentWithCancellation(filePath, cts.Token);
            Console.WriteLine("Document loaded successfully.");
            // Continue processing the document here.
        }
        catch (OperationCanceledException ex)
        {
            // Handle the cancellation gracefully.
            Console.WriteLine($"Loading canceled: {ex.Message}");
        }
    }

    // Loads a document using LoadOptions and a custom progress callback that checks the token.
    static Document LoadDocumentWithCancellation(string path, CancellationToken token)
    {
        var loadOptions = new LoadOptions
        {
            // Assign the callback that will be invoked periodically during loading.
            ProgressCallback = new LoadingProgressCallback(token)
        };

        // The Document constructor respects the LoadOptions and will abort if the callback throws.
        return new Document(path, loadOptions);
    }

    // Implements IDocumentLoadingCallback to abort loading when the token is signaled.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _token;
        private readonly DateTime _startTime = DateTime.Now;
        private const double MaxDurationSeconds = 5.0; // Optional hard limit.

        public LoadingProgressCallback(CancellationToken token)
        {
            _token = token;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            // Abort if the external CancellationToken has been triggered.
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException($"Loading canceled by token at {DateTime.Now}");

            // Optional: also abort if loading exceeds a predefined duration.
            double elapsed = (DateTime.Now - _startTime).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException($"Loading exceeded {MaxDurationSeconds}s; EstimatedProgress={args.EstimatedProgress}");
        }
    }
}
