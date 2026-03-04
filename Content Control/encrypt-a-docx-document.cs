using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class EncryptDocxExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Configure save options to encrypt the DOCX using a password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "MyPassword";

        // Save the encrypted document.
        string encryptedPath = "EncryptedDocument.docx";
        doc.Save(encryptedPath, saveOptions);

        // -----------------------------------------------------------------
        // Demonstrate loading the encrypted document with the correct password.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("MyPassword");
        Document loadedDoc = new Document(encryptedPath, loadOptions);

        // Verify that the document was loaded successfully.
        Console.WriteLine(loadedDoc.GetText().Trim()); // Outputs: Hello world!
    }
}
