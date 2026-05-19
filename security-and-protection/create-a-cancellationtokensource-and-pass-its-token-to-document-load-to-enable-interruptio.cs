using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    // Callback that checks a CancellationToken and aborts loading when requested.
    private class CancellationLoadingCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _token;

        public CancellationLoadingCallback(CancellationToken token)
        {
            _token = token;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException("Document loading was cancelled.");
        }
    }

    public static async Task Main(string[] args)
    {
        // Prepare a temporary folder for the demo files.
        string artifactsDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(artifactsDir);

        // Create a large document to make the load operation take noticeable time.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 2000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        string sourcePath = Path.Combine(artifactsDir, "LargeDocument.docx");
        doc.Save(sourcePath);
        if (!File.Exists(sourcePath))
            throw new InvalidOperationException("Failed to save the source document.");

        // Set up a CancellationTokenSource that will be cancelled shortly after loading starts.
        using var cts = new CancellationTokenSource();

        // Start the load operation in a separate task.
        Task loadTask = Task.Run(() =>
        {
            try
            {
                // Configure LoadOptions with a progress callback that can cancel the load.
                LoadOptions loadOptions = new LoadOptions
                {
                    ProgressCallback = new CancellationLoadingCallback(cts.Token)
                };

                // Load the document; the callback will monitor the token.
                Document loadedDoc = new Document(sourcePath, loadOptions);

                // If loading completes, save the loaded document to verify successful load.
                string loadedPath = Path.Combine(artifactsDir, "LoadedDocument.docx");
                loadedDoc.Save(loadedPath);
                Console.WriteLine($"Document loaded successfully. Pages: {loadedDoc.PageCount}");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Document loading was cancelled.");
            }
        }, cts.Token);

        // Cancel the operation after a short delay (e.g., 10 milliseconds).
        await Task.Delay(10);
        cts.Cancel();

        // Await the load task to ensure the program finishes after handling cancellation.
        await loadTask;
    }
}
