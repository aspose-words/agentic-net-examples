using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class ConvertToDocWithOptions
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

        // Example: protect the document with a password (RC4 encryption).
        saveOptions.Password = "MyPassword";

        // Example: preserve routing slip data if present.
        saveOptions.SaveRoutingSlip = true;

        // Save the document using the specified options.
        string outputPath = "ConvertedDocument.doc";
        doc.Save(outputPath, saveOptions);

        // To demonstrate loading the password‑protected DOC, use LoadOptions.
        LoadOptions loadOptions = new LoadOptions("MyPassword");
        Document loadedDoc = new Document(outputPath, loadOptions);

        // Verify that the text was preserved.
        Console.WriteLine(loadedDoc.GetText().Trim()); // Outputs: Hello world!
    }
}
