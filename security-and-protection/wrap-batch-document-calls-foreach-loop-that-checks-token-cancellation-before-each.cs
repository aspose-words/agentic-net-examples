using System;
using System.Collections.Generic;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

public class TokenCancellationCallback : IDocumentLoadingCallback
{
    private readonly CancellationToken _cancellationToken;

    public TokenCancellationCallback(CancellationToken cancellationToken)
    {
        _cancellationToken = cancellationToken;
    }

    // This method is called periodically during document loading.
    // If the token has been cancelled we abort the load by throwing.
    public void Notify(DocumentLoadingArgs args)
    {
        if (_cancellationToken.IsCancellationRequested)
            throw new OperationCanceledException("Document loading was cancelled.");
    }
}

public static class DocumentBatchLoader
{
    // Loads a collection of documents, checking the cancellation token before each load.
    // Returns the successfully loaded documents.
    public static List<Document> LoadDocuments(IEnumerable<string> filePaths, CancellationToken cancellationToken)
    {
        var loadedDocs = new List<Document>();

        foreach (var path in filePaths)
        {
            // Check cancellation before starting the load.
            if (cancellationToken.IsCancellationRequested)
                break; // Exit the loop if cancellation was requested.

            // Configure load options with a progress callback that also respects the token.
            var loadOptions = new LoadOptions
            {
                ProgressCallback = new TokenCancellationCallback(cancellationToken)
            };

            try
            {
                // Load the document using the constructor that accepts a filename and LoadOptions.
                var doc = new Document(path, loadOptions);
                loadedDocs.Add(doc);
            }
            catch (OperationCanceledException)
            {
                // Loading was cancelled; stop processing further files.
                break;
            }
            catch (Exception ex)
            {
                // Handle other loading errors as needed (log, rethrow, etc.).
                Console.WriteLine($"Failed to load '{path}': {ex.Message}");
            }
        }

        return loadedDocs;
    }
}

public static class Program
{
    // Entry point required for a console application.
    public static void Main(string[] args)
    {
        // Example file list – replace with actual paths.
        var files = new[] { "doc1.docx", "doc2.docx", "doc3.docx" };

        using var cts = new CancellationTokenSource();
        // Optionally cancel after a timeout or based on user input.
        // cts.CancelAfter(TimeSpan.FromSeconds(10));

        List<Document> docs = DocumentBatchLoader.LoadDocuments(files, cts.Token);
        Console.WriteLine($"Successfully loaded {docs.Count} document(s).");
    }
}
