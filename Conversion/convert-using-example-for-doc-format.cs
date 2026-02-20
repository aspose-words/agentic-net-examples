using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading; // For LoadOptions if needed

class ConvertToDocExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");
        builder.Writeln("This document will be saved in the legacy DOC format.");

        // Configure save options for the DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc)
        {
            // Example: embed PostScript fonts if needed.
            AllowEmbeddingPostScriptFonts = false,
            // Example: compress all metafiles.
            AlwaysCompressMetafiles = true,
            // Example: set a password for opening the saved DOC file.
            Password = "SecretPassword"
        };

        // Define the output path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.doc");

        // Save the document using the specified options.
        doc.Save(outputPath, saveOptions);

        // -----------------------------------------------------------------
        // Load the saved DOC file back into a Document object to verify.
        // The LoadFormat will be detected automatically.
        Document loadedDoc = new Document(outputPath, new LoadOptions("SecretPassword"));

        // Output the text to the console to confirm successful load.
        Console.WriteLine("Loaded document text:");
        Console.WriteLine(loadedDoc.GetText());

        // Optionally, re‑save the document to another format (e.g., PDF) to demonstrate further conversion.
        string pdfPath = Path.Combine(Environment.CurrentDirectory, "Result.pdf");
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);
    }
}
