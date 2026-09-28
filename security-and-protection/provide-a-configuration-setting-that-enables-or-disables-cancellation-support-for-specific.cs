using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    // Configuration setting: enable or disable cancellation support for processing stages.
    private static readonly bool EnableCancellation = true;

    public static void Main()
    {
        // Create a simple source document.
        const string sourcePath = "source.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        doc.Save(sourcePath);

        // Prepare a cancellation token based on the configuration.
        CancellationToken token = EnableCancellation ? new CancellationTokenSource().Token : CancellationToken.None;

        // Load the document. If cancellation is enabled and a cancellation is requested,
        // the operation will be aborted before the load.
        token.ThrowIfCancellationRequested();
        LoadOptions loadOptions = new LoadOptions(); // No CancellationToken property in this version.
        Document loadedDoc = new Document(sourcePath, loadOptions);

        // Apply a simple protection to demonstrate a processing stage.
        loadedDoc.Protect(ProtectionType.ReadOnly, "pwd");

        // Save the protected document. Again, respect the cancellation token.
        token.ThrowIfCancellationRequested();
        const string outputPath = "protected.docx";
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        loadedDoc.Save(outputPath, saveOptions);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        // Clean up temporary files (optional).
        File.Delete(sourcePath);
        File.Delete(outputPath);
    }
}
