using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class ConvertToDoc
{
    static void Main()
    {
        // Create a new blank document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Initialize save options for the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        // Example: protect the saved file with a password (RC4 encryption).
        saveOptions.Password = "MyPassword";
        // Example: preserve routing slip information if present.
        saveOptions.SaveRoutingSlip = true;

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ConvertedDocument.doc");

        // Save the document using the specified DOC save options.
        doc.Save(outputPath, saveOptions);

        // Load the password‑protected DOC file to verify it was saved correctly.
        LoadOptions loadOptions = new LoadOptions { Password = "MyPassword" };
        Document loadedDoc = new Document(outputPath, loadOptions);
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
