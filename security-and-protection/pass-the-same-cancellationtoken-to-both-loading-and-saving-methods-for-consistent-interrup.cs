using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static async Task Main(string[] args)
    {
        // Create a cancellation token source and obtain the token.
        using var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Define temporary file paths.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(tempFolder);
        string sourcePath = Path.Combine(tempFolder, "sample.docx");
        string outputPath = Path.Combine(tempFolder, "output.docx");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words! This document will be loaded and saved using the same CancellationToken.");

        // Save the source document synchronously, wrapped in a Task that observes the token.
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        await Task.Run(() =>
        {
            token.ThrowIfCancellationRequested();
            doc.Save(sourcePath, saveOptions);
        }, token);

        // Load the document synchronously, wrapped in a Task that observes the token.
        var loadOptions = new LoadOptions();
        Document loadedDoc = await Task.Run(() =>
        {
            token.ThrowIfCancellationRequested();
            return new Document(sourcePath, loadOptions);
        }, token);

        // Save the loaded document to a new file, again wrapped in a Task that observes the token.
        await Task.Run(() =>
        {
            token.ThrowIfCancellationRequested();
            loadedDoc.Save(outputPath, saveOptions);
        }, token);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The output file was not created at '{outputPath}'.");
        }

        // Optional cleanup (commented out to keep files for inspection).
        // File.Delete(sourcePath);
        // File.Delete(outputPath);
        // Directory.Delete(tempFolder);
    }
}
