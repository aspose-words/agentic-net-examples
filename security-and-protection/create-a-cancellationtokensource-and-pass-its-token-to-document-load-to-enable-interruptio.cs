using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Loading; // IDocumentLoadingCallback is in this namespace

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the demo files.
        string artifactsDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document and save it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document used to demonstrate cancellation during loading.");
        string filePath = Path.Combine(artifactsDir, "Sample.docx");
        doc.Save(filePath);

        // Set up a CancellationTokenSource that will cancel after a short delay.
        using var cts = new CancellationTokenSource();
        Task.Run(async () =>
        {
            await Task.Delay(100); // Adjust the delay as needed.
            cts.Cancel();
        });

        // Attempt to load the document with the cancellation token.
        try
        {
            // Use LoadOptions with a ProgressCallback that checks the token.
            LoadOptions loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingCancellationCallback(cts.Token)
            };

            // Load the document using the constructor that accepts LoadOptions.
            Document loadedDoc = new Document(filePath, loadOptions);

            // If loading completes before cancellation, output some information.
            Console.WriteLine("Document loaded successfully. Text length: " + loadedDoc.GetText().Length);
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document loading was canceled via CancellationToken.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }

    // Callback that aborts loading when the cancellation token is set.
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
                throw new OperationCanceledException("Loading was canceled.");
        }
    }
}
