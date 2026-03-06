using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");

        // Save the document as a legacy .doc file with a password.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        saveOptions.Password = "MyPassword";
        doc.Save("Hello.doc", saveOptions);

        // Load the password‑protected document using LoadOptions.
        LoadOptions loadOptions = new LoadOptions("MyPassword");
        Document loadedDoc = new Document("Hello.doc", loadOptions);

        // Output the text to verify that the document was loaded correctly.
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
