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
        // Create a cancellation token source that will not be cancelled.
        using var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Prepare a temporary folder for the demo files.
        string artifactsDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(artifactsDir);

        // Path of the document to be saved and later loaded.
        string docPath = Path.Combine(artifactsDir, "SampleDocument.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words with CancellationToken!");

        // -----------------------------------------------------------------
        // 2. Save the document using the same CancellationToken.
        // -----------------------------------------------------------------
        // Since Aspose.Words does not provide a Save overload that accepts a
        // CancellationToken, we manually check the token before invoking Save.
        if (token.IsCancellationRequested)
            throw new OperationCanceledException("Save operation was cancelled.", token);

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        doc.Save(docPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(docPath))
            throw new FileNotFoundException("The document was not saved correctly.", docPath);

        // -----------------------------------------------------------------
        // 3. Load the document using the same CancellationToken.
        // -----------------------------------------------------------------
        // Similarly, we check the token before loading.
        if (token.IsCancellationRequested)
            throw new OperationCanceledException("Load operation was cancelled.", token);

        LoadOptions loadOptions = new LoadOptions();
        Document loadedDoc = new Document(docPath, loadOptions);

        // Simple validation: the loaded text should match the original.
        string originalText = doc.GetText().Trim();
        string loadedText = loadedDoc.GetText().Trim();

        if (!originalText.Equals(loadedText, StringComparison.Ordinal))
            throw new InvalidOperationException("The loaded document content does not match the original.");

        // Cleanup: delete the temporary files (optional).
        File.Delete(docPath);
        Directory.Delete(artifactsDir, true);
    }
}
