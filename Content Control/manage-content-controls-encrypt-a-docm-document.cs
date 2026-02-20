using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class EncryptDocmExample
{
    static void Main()
    {
        // Folder where the document will be saved.
        string artifactsDir = @"C:\Temp\";

        // Create a new blank document.
        Document doc = new Document();

        // Add some content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this DOCM file is encrypted with a password.");

        // Configure save options for DOCM format and set the encryption password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);
        saveOptions.Password = "MyPassword";

        // Save the encrypted DOCM file.
        string encryptedPath = artifactsDir + "EncryptedDocument.docm";
        doc.Save(encryptedPath, saveOptions);

        // Load the encrypted document using the correct password.
        LoadOptions loadOptions = new LoadOptions("MyPassword");
        Document loadedDoc = new Document(encryptedPath, loadOptions);

        // Verify that the document was loaded successfully.
        Console.WriteLine("Document text after loading:");
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
