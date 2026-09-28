using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    // Helper that saves a document while respecting a CancellationToken.
    private static void SaveDocument(Document doc, string path, SaveOptions options, CancellationToken token)
    {
        // Throw if cancellation was requested before the operation.
        token.ThrowIfCancellationRequested();

        // Perform the synchronous save.
        doc.Save(path, options);

        // Throw if cancellation was requested after the operation.
        token.ThrowIfCancellationRequested();
    }

    // Helper that loads a document while respecting a CancellationToken.
    private static Document LoadDocument(string path, LoadOptions options, CancellationToken token)
    {
        // Throw if cancellation was requested before the operation.
        token.ThrowIfCancellationRequested();

        // Perform the synchronous load.
        var doc = new Document(path, options);

        // Throw if cancellation was requested after the operation.
        token.ThrowIfCancellationRequested();

        return doc;
    }

    public static void Main()
    {
        // Create a cancellation token source and obtain the token.
        using var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Create a simple document with one paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Path for the temporary file.
        string filePath = "sample.docx";

        // Save the document using the helper that respects the same token.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        SaveDocument(doc, filePath, saveOptions, token);

        // Load the document using the helper that respects the same token.
        LoadOptions loadOptions = new LoadOptions();
        Document loadedDoc = LoadDocument(filePath, loadOptions, token);

        // Validate that the loaded document contains at least one paragraph.
        if (loadedDoc.GetChildNodes(NodeType.Paragraph, true).Count == 0)
        {
            throw new InvalidOperationException("The loaded document does not contain any paragraphs.");
        }

        // Clean up the temporary file.
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }
}
