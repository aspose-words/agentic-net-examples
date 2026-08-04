using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a simple document and save it locally.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document used to demonstrate a timeout.");
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        doc.Save(filePath);

        // Set up a cancellation token that will be triggered after 1 second.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(TimeSpan.FromSeconds(1));

        // LoadOptions with a progress callback that checks the cancellation token.
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback(cts.Token)
        };

        try
        {
            // Attempt to load the document. The callback will abort the load when the token is cancelled.
            Document loadedDoc = new Document(filePath, loadOptions);
            Console.WriteLine("Document loaded successfully (no timeout).");
        }
        catch (OperationCanceledException ex)
        {
            // Expected when the timeout occurs.
            Console.WriteLine($"Loading cancelled due to timeout: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Any other unexpected errors.
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Implements the Aspose.Words loading callback and aborts loading when cancellation is requested.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _token;

        public LoadingProgressCallback(CancellationToken token)
        {
            _token = token;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}");
        }
    }
}
