using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Create a new blank document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Configure save options for the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        saveOptions.Password = "MyPassword";          // Protect the file with a password.
        saveOptions.SaveRoutingSlip = true;           // Preserve routing slip if present.

        // Save the document using the specified options.
        string outputPath = "DocSaveOptions.SaveAsDoc.doc";
        doc.Save(outputPath, saveOptions);

        // Load the saved document using the password.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "MyPassword";
        Document loadedDoc = new Document(outputPath, loadOptions);

        // Output the document text to verify successful load.
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
