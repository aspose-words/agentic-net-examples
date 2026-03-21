using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Ensure a document exists for the demo.
        string docPath = Path.Combine(Path.GetTempPath(), "Big document.docx");
        if (!File.Exists(docPath))
        {
            var tempDoc = new Document();
            var builder = new DocumentBuilder(tempDoc);
            builder.Writeln("This is a sample document created for the cancellation token demo.");
            tempDoc.Save(docPath);
        }

        // Create a cancellation token source that will trigger cancellation after 500 ms.
        using var cancellationSource = new CancellationTokenSource();
        cancellationSource.CancelAfter(500);

        // Configure load options to use a custom loading callback that checks the token.
        var loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback(cancellationSource.Token)
        };

        try
        {
            // Load the document; the callback will abort loading if cancellation is requested.
            var doc = new Document(docPath, loadOptions);
            Console.WriteLine("Document loaded successfully.");
        }
        catch (OperationCanceledException ex)
        {
            // Loading was interrupted.
            Console.WriteLine($"Loading canceled: {ex.Message}");
        }
    }

    // Implements IDocumentLoadingCallback and aborts loading when the token signals cancellation.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public LoadingProgressCallback(CancellationToken cancellationToken) => _cancellationToken = cancellationToken;

        public void Notify(DocumentLoadingArgs args)
        {
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}");
        }
    }
}
