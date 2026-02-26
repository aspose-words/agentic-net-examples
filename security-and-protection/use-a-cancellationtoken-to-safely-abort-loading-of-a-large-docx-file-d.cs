using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the large DOCX file.
        string filePath = "BigDocument.docx";

        // Create a cancellation token source that will request cancellation after a timeout.
        var cts = new CancellationTokenSource();
        cts.CancelAfter(TimeSpan.FromMilliseconds(500)); // Adjust as needed.

        // Configure load options to use a custom progress callback that checks the token.
        var loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback(cts.Token)
        };

        try
        {
            // Load the document with the specified load options.
            Document doc = new Document(filePath, loadOptions);

            // Document loaded successfully – continue processing.
            Console.WriteLine("Document loaded successfully.");
        }
        catch (OperationCanceledException ex)
        {
            // Loading was aborted via the cancellation token.
            Console.WriteLine($"Loading canceled: {ex.Message}");
        }
    }

    // Implements IDocumentLoadingCallback to monitor loading progress.
    // Throws OperationCanceledException when the supplied token signals cancellation.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _token;
        private readonly DateTime _startTime;

        public LoadingProgressCallback(CancellationToken token)
        {
            _token = token;
            _startTime = DateTime.Now;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            // Abort loading if cancellation was requested.
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException(
                    $"Loading aborted at {DateTime.Now}, progress {args.EstimatedProgress}%.");

            // Optional: additional time‑based abort logic can be added here.
            // double elapsedSeconds = (DateTime.Now - _startTime).TotalSeconds;
            // if (elapsedSeconds > 10)
            //     throw new OperationCanceledException($"Loading timed out after {elapsedSeconds}s.");
        }
    }
}
