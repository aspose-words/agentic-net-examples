using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("Hello world!");

        // Initialize save options for the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Optional: set a password or other options here.
        // saveOptions.Password = "MyPassword";

        // Save the document as a .doc file using the specified options.
        string outputPath = "ConvertedDocument.doc";
        doc.Save(outputPath, saveOptions);

        Console.WriteLine($"Document saved to {outputPath}");
    }
}
