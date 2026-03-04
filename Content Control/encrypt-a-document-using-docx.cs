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
        builder.Writeln("Hello world! This document is encrypted.");

        // Configure save options to encrypt the DOCX with a password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "MySecretPassword";

        // Save the encrypted document.
        string outputPath = "EncryptedDocument.docx";
        doc.Save(outputPath, saveOptions);

        // Attempt to load the encrypted document without a password – this will throw.
        try
        {
            Document loadFail = new Document(outputPath);
        }
        catch (IncorrectPasswordException)
        {
            Console.WriteLine("Failed to open without password (as expected).");
        }

        // Load the encrypted document using the correct password.
        LoadOptions loadOptions = new LoadOptions("MySecretPassword");
        Document loadedDoc = new Document(outputPath, loadOptions);

        // Verify the content.
        Console.WriteLine("Loaded text: " + loadedDoc.GetText().Trim());
    }
}
