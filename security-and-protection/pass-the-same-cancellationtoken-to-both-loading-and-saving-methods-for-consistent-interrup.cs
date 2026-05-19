using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a cancellation token source that could be used to cancel operations.
        using var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Prepare a temporary folder and file path for the sample document.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(tempFolder);
        string filePath = Path.Combine(tempFolder, "Sample.docx");

        // -----------------------------------------------------------------
        // Create a simple document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words with CancellationToken!");

        // -----------------------------------------------------------------
        // Save the document while using the same CancellationToken for progress monitoring.
        // -----------------------------------------------------------------
        try
        {
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                // The callback will abort the save operation if the token is cancelled.
                ProgressCallback = new SavingProgressCallback(token)
            };

            doc.Save(filePath, saveOptions);
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document saving was canceled.");
            return;
        }

        // -----------------------------------------------------------------
        // Load the document using the same CancellationToken for progress monitoring.
        // -----------------------------------------------------------------
        Document loadedDoc;
        try
        {
            LoadOptions loadOptions = new LoadOptions
            {
                // The callback will abort the load operation if the token is cancelled.
                ProgressCallback = new LoadingProgressCallback(token)
            };

            loadedDoc = new Document(filePath, loadOptions);
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document loading was canceled.");
            return;
        }

        // -----------------------------------------------------------------
        // Verify that the document was loaded correctly.
        // -----------------------------------------------------------------
        string text = loadedDoc.GetText().Trim();
        Console.WriteLine($"Loaded document text: {text}");
    }

    // Callback for saving that checks the cancellation token.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _token;

        public SavingProgressCallback(CancellationToken token) => _token = token;

        public void Notify(DocumentSavingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException("Saving was canceled via token.");
        }
    }

    // Callback for loading that checks the cancellation token.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _token;

        public LoadingProgressCallback(CancellationToken token) => _token = token;

        public void Notify(DocumentLoadingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException("Loading was canceled via token.");
        }
    }
}
