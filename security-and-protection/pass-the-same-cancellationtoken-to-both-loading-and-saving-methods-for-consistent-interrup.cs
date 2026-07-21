using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;      // Needed for LoadOptions
using Aspose.Words.Saving;      // Needed for OoxmlSaveOptions and IDocumentSavingCallback

public class Program
{
    public static void Main()
    {
        // Create a single CancellationTokenSource to be used for both load and save operations.
        using var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Define file paths in the current working directory.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        string destinationPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple document and save it (initial creation, no token needed).
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words with a shared CancellationToken.");
        doc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document while respecting the same CancellationToken.
        // -----------------------------------------------------------------
        Document loadedDoc = LoadDocumentWithCancellation(sourcePath, token);

        // -----------------------------------------------------------------
        // 3. Save the loaded document while also respecting the same token.
        // -----------------------------------------------------------------
        SaveDocumentWithCancellation(loadedDoc, destinationPath, token);

        // -----------------------------------------------------------------
        // 4. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(destinationPath))
            throw new InvalidOperationException("The output document was not saved.");

        // Indicate successful completion.
        Console.WriteLine("Document processed successfully.");
    }

    private static Document LoadDocumentWithCancellation(string path, CancellationToken token)
    {
        // Abort early if cancellation was requested before loading.
        token.ThrowIfCancellationRequested();

        // LoadOptions can be used for passwords or other settings; not needed here.
        var loadOptions = new LoadOptions();

        // Load the document from file using the same token for consistency.
        return new Document(path, loadOptions);
    }

    private static void SaveDocumentWithCancellation(Document doc, string path, CancellationToken token)
    {
        // Abort early if cancellation was requested before saving.
        token.ThrowIfCancellationRequested();

        // Configure save options; using OoxmlSaveOptions to allow a progress callback.
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            // The callback checks the same token during the save operation.
            ProgressCallback = new CancelSavingCallback(token)
        };

        // Save the document using the configured options.
        doc.Save(path, saveOptions);
    }

    // Implementation of IDocumentSavingCallback that throws if the token signals cancellation.
    private class CancelSavingCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public CancelSavingCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        public void Notify(DocumentSavingArgs args)
        {
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException("Document saving was cancelled via CancellationToken.");
        }
    }
}
