using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a simple source document and save it locally.
        // -----------------------------------------------------------------
        const string sourcePath = "source.docx";

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");
        doc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Prepare a CancellationTokenSource.
        //    (The current Aspose.Words API does not expose a Load method
        //     that accepts a CancellationToken, so we create the token
        //     but load the document using the standard constructor.)
        // -----------------------------------------------------------------
        using var cts = new CancellationTokenSource();

        // -----------------------------------------------------------------
        // 3. Load the document with LoadOptions.
        // -----------------------------------------------------------------
        var loadOptions = new LoadOptions(); // No special options needed.
        // The token is not used because the API version does not support it.
        // If a future version adds overload Document.Load(string, LoadOptions, CancellationToken),
        // the token can be passed there.
        Document loadedDoc = new Document(sourcePath, loadOptions);

        // -----------------------------------------------------------------
        // 4. Save the loaded document to verify that loading succeeded.
        // -----------------------------------------------------------------
        const string outputPath = "loaded.docx";
        loadedDoc.Save(outputPath);

        // -----------------------------------------------------------------
        // 5. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved as expected.");

        Console.WriteLine("Document loaded and saved successfully.");
    }
}
