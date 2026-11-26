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
        // Configuration: enable or disable cancellation support.
        bool enableCancellation = true;

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World");

        // Save the original document.
        string originalPath = "Original.docx";
        doc.Save(originalPath);
        if (!File.Exists(originalPath))
            throw new InvalidOperationException("Failed to create the original document.");

        // Prepare a cancellation token source.
        CancellationTokenSource cts = new CancellationTokenSource();

        // Example: cancel the operation before loading if cancellation is enabled.
        if (enableCancellation && cts.Token.IsCancellationRequested)
            throw new OperationCanceledException("Loading was cancelled.");

        // Load the document (no native CancellationToken property on LoadOptions).
        LoadOptions loadOptions = new LoadOptions();
        Document loadedDoc = new Document(originalPath, loadOptions);

        // Example: cancel the operation before saving if cancellation is enabled.
        if (enableCancellation && cts.Token.IsCancellationRequested)
            throw new OperationCanceledException("Saving was cancelled.");

        // Save the document (no native CancellationToken property on OoxmlSaveOptions).
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        string outputPath = "Processed.docx";
        loadedDoc.Save(outputPath, saveOptions);
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to save the processed document.");

        // Indicate completion.
        Console.WriteLine("Document processing completed successfully.");
    }
}
