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
        // Create a cancellation token source and obtain the token.
        var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Prepare output directory and file paths.
        string outputDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(outputDir);
        string filePath = Path.Combine(outputDir, "Sample.docx");
        string copyPath = Path.Combine(outputDir, "SampleCopy.docx");

        // Create a simple document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words with a shared CancellationToken.");

        // Configure save options with a progress callback that respects the token.
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.ProgressCallback = new TokenProgressCallback(token);

        // Save the document using the same token.
        doc.Save(filePath, saveOptions);

        // Before loading, check the token for cancellation.
        if (token.IsCancellationRequested)
            throw new OperationCanceledException(token);

        // Load the document (no direct token support, but we can abort beforehand).
        var loadOptions = new LoadOptions(); // No password needed for this example.
        var loadedDoc = new Document(filePath, loadOptions);

        // Optional verification of content.
        string loadedText = loadedDoc.GetText().Trim();

        // Save the loaded document again using the same token and save options.
        loadedDoc.Save(copyPath, saveOptions);
    }

    // Progress callback that throws if the shared CancellationToken is cancelled.
    private class TokenProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _token;
        public TokenProgressCallback(CancellationToken token) => _token = token;
        public void Notify(DocumentSavingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException(_token);
        }
    }
}
