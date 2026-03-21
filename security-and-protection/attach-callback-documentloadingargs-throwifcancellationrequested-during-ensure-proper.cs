using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

public class LoadingProgressCallback : IDocumentLoadingCallback
{
    private readonly CancellationToken _cancellationToken;
    private readonly DateTime _startTime;

    public LoadingProgressCallback(CancellationToken cancellationToken)
    {
        _cancellationToken = cancellationToken;
        _startTime = DateTime.Now;
    }

    // This method is called repeatedly while the document is being loaded.
    public void Notify(DocumentLoadingArgs args)
    {
        // If the external token signals cancellation, abort the loading by throwing.
        if (_cancellationToken.IsCancellationRequested)
            throw new OperationCanceledException($"Loading cancelled by token. EstimatedProgress = {args.EstimatedProgress}");
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare a temporary document to load.
        string tempPath = Path.Combine(Path.GetTempPath(), "MyDocument.docx");
        var tempDoc = new Document();
        var builder = new DocumentBuilder(tempDoc);
        builder.Writeln("Sample content for testing cancellation.");
        tempDoc.Save(tempPath);

        // Cancel the load after 500 milliseconds.
        using var cts = new CancellationTokenSource(500);

        // Attach the custom progress callback to the load options.
        var loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback(cts.Token)
        };

        try
        {
            // Load the document with the specified options.
            var doc = new Document(tempPath, loadOptions);
            Console.WriteLine("Document loaded successfully.");
        }
        catch (OperationCanceledException ex)
        {
            // Handle the cancellation gracefully.
            Console.WriteLine($"Loading was cancelled: {ex.Message}");
        }
        finally
        {
            // Clean up the temporary file.
            if (File.Exists(tempPath))
                File.Delete(tempPath);
        }
    }
}
