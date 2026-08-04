using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    // Callback that checks a CancellationToken and aborts loading when requested.
    private class LoadingCancellationCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _token;

        public LoadingCancellationCallback(CancellationToken token)
        {
            _token = token;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException("Document loading was cancelled.");
        }
    }

    public static void Main()
    {
        // Prepare a temporary folder for the sample files.
        string artifactsDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(artifactsDir);
        string filePath = Path.Combine(artifactsDir, "Sample.docx");

        // Create a simple document and save it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document used to demonstrate loading cancellation.");
        doc.Save(filePath);

        // Create a CancellationTokenSource and cancel it immediately.
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        // Set up LoadOptions with a progress callback that respects the token.
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingCancellationCallback(cts.Token)
        };

        // Attempt to load the document; expect an OperationCanceledException.
        try
        {
            Document loadedDoc = new Document(filePath, loadOptions);
            // If loading succeeds (unlikely), indicate success.
            Console.WriteLine("Document loaded successfully (cancellation not triggered).");
        }
        catch (OperationCanceledException ex)
        {
            // Expected path when cancellation occurs.
            Console.WriteLine($"Loading was interrupted: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Any other unexpected exceptions.
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
