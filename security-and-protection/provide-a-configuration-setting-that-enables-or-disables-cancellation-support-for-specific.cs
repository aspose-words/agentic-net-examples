using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    // Configuration that controls whether cancellation support is enabled.
    public class CancellationConfig
    {
        public bool EnableCancellationSupport { get; set; }
    }

    public static void Main()
    {
        // Prepare configuration.
        var config = new CancellationConfig { EnableCancellationSupport = true };

        // Create a simple document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Define file paths.
        string originalPath = Path.Combine(Path.GetTempPath(), "original.docx");
        string processedPath = Path.Combine(Path.GetTempPath(), "processed.docx");

        // Save the original document.
        doc.Save(originalPath);
        ValidateFileExists(originalPath, "original document");

        // Prepare cancellation token if enabled (currently not used by Aspose.Words APIs).
        CancellationToken token = CancellationToken.None;
        CancellationTokenSource cts = null;
        if (config.EnableCancellationSupport)
        {
            cts = new CancellationTokenSource();
            token = cts.Token;
        }

        // Load the document. Aspose.Words does not expose a CancellationToken on LoadOptions,
        // so we simply load the document normally.
        var loadOptions = new LoadOptions();
        var loadedDoc = new Document(originalPath, loadOptions);

        // Save the document. Aspose.Words does not expose a CancellationToken on OoxmlSaveOptions,
        // so we save without attaching the token.
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        loadedDoc.Save(processedPath, saveOptions);
        ValidateFileExists(processedPath, "processed document");

        // Clean up temporary files.
        File.Delete(originalPath);
        File.Delete(processedPath);
    }

    private static void ValidateFileExists(string path, string description)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"The {description} was not created at '{path}'.");
        }
    }
}
