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
        // Prepare file paths.
        const string sourcePath = "source.docx";
        const string resultPath = "result.docx";

        // Create a simple document and save it.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words with cancellation support!");
        doc.Save(sourcePath);

        // Create a single CancellationTokenSource to be shared.
        using var cts = new CancellationTokenSource();

        // Load the document using the same token via a loading progress callback.
        var loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback(cts.Token)
        };
        var loadedDoc = new Document(sourcePath, loadOptions);

        // Save the document using the same token via a saving progress callback.
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback(cts.Token)
        };
        loadedDoc.Save(resultPath, saveOptions);

        // Validate that the result file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        Console.WriteLine("Document loaded and saved successfully with shared CancellationToken.");
    }

    // Loading callback that checks the shared token.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _token;

        public LoadingProgressCallback(CancellationToken token) => _token = token;

        public void Notify(DocumentLoadingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException("Loading was cancelled.", _token);
        }
    }

    // Saving callback that checks the shared token.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _token;

        public SavingProgressCallback(CancellationToken token) => _token = token;

        public void Notify(DocumentSavingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException("Saving was cancelled.", _token);
        }
    }
}
