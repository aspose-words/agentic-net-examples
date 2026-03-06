using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Create a new document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");

        // Set up DOC save options (password protection and routing slip preservation).
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        saveOptions.Password = "Secret123";
        saveOptions.SaveRoutingSlip = true;

        // Ensure the output directory exists.
        string outputPath = "Output/Hello.doc";
        string? outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        // Save the document as a .doc file using the specified options.
        doc.Save(outputPath, saveOptions);

        // Load the password‑protected DOC file.
        LoadOptions loadOptions = new LoadOptions("Secret123");
        Document loadedDoc = new Document(outputPath, loadOptions);

        // Output the loaded document's text to verify successful conversion.
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
