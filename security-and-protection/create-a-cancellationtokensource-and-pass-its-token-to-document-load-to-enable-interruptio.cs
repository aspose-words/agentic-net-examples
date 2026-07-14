using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    // Callback that checks the cancellation token and aborts loading when requested.
    private class CancelLoadingCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _token;

        public CancelLoadingCallback(CancellationToken token)
        {
            _token = token;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException();
        }
    }

    public static void Main()
    {
        // Prepare a folder for temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create and save a simple document.
        string samplePath = Path.Combine(artifactsDir, "Sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document used to demonstrate cancellation during loading.");
        doc.Save(samplePath);

        // Set up a CancellationTokenSource that will cancel shortly after loading starts.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel after 10 milliseconds.
            Task.Delay(10).ContinueWith(_ => cts.Cancel());

            try
            {
                // Configure LoadOptions with a progress callback that respects the token.
                LoadOptions loadOptions = new LoadOptions
                {
                    ProgressCallback = new CancelLoadingCallback(cts.Token)
                };

                // Load the document; the callback will throw if cancellation occurs.
                Document loadedDoc = new Document(samplePath, loadOptions);

                // If loading completes, output a simple verification.
                Console.WriteLine("Document loaded successfully. Text length: " + loadedDoc.GetText().Length);
            }
            catch (OperationCanceledException)
            {
                // Expected when the token triggers cancellation.
                Console.WriteLine("Document loading was canceled.");
            }
        }
    }
}
