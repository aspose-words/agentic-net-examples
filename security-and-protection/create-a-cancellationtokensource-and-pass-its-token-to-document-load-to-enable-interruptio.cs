using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    // Callback that checks the cancellation token during document loading.
    private class LoadingCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _token;

        public LoadingCallback(CancellationToken token)
        {
            _token = token;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException("Document loading was canceled.");
        }
    }

    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple document and save it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        string filePath = Path.Combine(outputDir, "sample.docx");
        doc.Save(filePath);

        // Set up cancellation.
        using var cts = new CancellationTokenSource();
        // Cancel immediately to demonstrate interruption.
        cts.Cancel();

        // Load the document with a progress callback that respects the token.
        try
        {
            LoadOptions loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingCallback(cts.Token)
            };

            Document loadedDoc = new Document(filePath, loadOptions);
            // If loading succeeds (unlikely in this example), write a message.
            Console.WriteLine("Document loaded successfully.");
        }
        catch (OperationCanceledException ex)
        {
            // Expected path when the token triggers cancellation.
            Console.WriteLine(ex.Message);
        }
    }
}
