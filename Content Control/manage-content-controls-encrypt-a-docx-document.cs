using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class EncryptDocxExample
{
    static void Main()
    {
        // Path where the encrypted document will be saved.
        string outputPath = "EncryptedDocument.docx";

        // Create a new blank document.
        Document doc = new Document();

        // Add some content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this document is encrypted with a password.");

        // Configure save options to encrypt the DOCX using the ECMA376 algorithm.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "MySecretPassword";

        // Save the document with encryption.
        doc.Save(outputPath, saveOptions);

        // -----------------------------------------------------------------
        // Demonstrate loading the encrypted document using the correct password.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "MySecretPassword";
        Document loadedDoc = new Document(outputPath, loadOptions);

        // Verify that the document was loaded successfully.
        Console.WriteLine("Loaded text: " + loadedDoc.GetText().Trim());
    }
}
