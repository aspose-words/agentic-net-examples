using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

public class LoadingProgressCallback : IDocumentLoadingCallback
{
    private readonly CancellationToken _cancellationToken;
    private readonly DateTime _loadingStartedAt;

    public LoadingProgressCallback(CancellationToken cancellationToken)
    {
        _cancellationToken = cancellationToken;
        _loadingStartedAt = DateTime.Now;
    }

    // Called periodically during document loading.
    public void Notify(DocumentLoadingArgs args)
    {
        // Abort if the external token requests cancellation.
        if (_cancellationToken.IsCancellationRequested)
            throw new OperationCanceledException($"Loading canceled by token. EstimatedProgress = {args.EstimatedProgress}");

        // Optional: also cancel after a time limit.
        double elapsedSeconds = (DateTime.Now - _loadingStartedAt).TotalSeconds;
        const double maxDuration = 5.0; // seconds
        if (elapsedSeconds > maxDuration)
            throw new OperationCanceledException($"Loading exceeded max duration ({maxDuration}s). EstimatedProgress = {args.EstimatedProgress}");
    }
}

public class DocumentProcessor
{
    public void LoadDocumentWithCancellation(string filePath, CancellationToken cancellationToken)
    {
        // Configure load options to use the progress callback.
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback(cancellationToken)
        };

        try
        {
            // Load the document; the callback will monitor progress and abort if needed.
            Document doc = new Document(filePath, loadOptions);

            // Document loaded successfully – proceed with processing.
            Console.WriteLine($"Document loaded. Page count: {doc.PageCount}");
        }
        catch (OperationCanceledException ex)
        {
            // Handle cancellation (e.g., log, clean up, inform user).
            Console.WriteLine($"Document loading was canceled: {ex.Message}");
        }
        catch (FileCorruptedException ex)
        {
            // Handle corrupted file scenario.
            Console.WriteLine($"File is corrupted: {ex.Message}");
        }
        catch (UnsupportedFileFormatException ex)
        {
            // Handle unsupported format scenario.
            Console.WriteLine($"Unsupported format: {ex.Message}");
        }
    }
}

// Example usage.
public class Program
{
    public static void Main()
    {
        string docPath = @"C:\Docs\LargeDocument.docx";

        // Create a cancellation token source that cancels after 3 seconds.
        using (CancellationTokenSource cts = new CancellationTokenSource(TimeSpan.FromSeconds(3)))
        {
            DocumentProcessor processor = new DocumentProcessor();
            processor.LoadDocumentWithCancellation(docPath, cts.Token);
        }
    }
}
