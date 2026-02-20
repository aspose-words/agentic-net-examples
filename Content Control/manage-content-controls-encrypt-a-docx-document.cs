using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;   // LoadOptions resides in this namespace
using Aspose.Words;           // For IncorrectPasswordException

class EncryptDocxExample
{
    static void Main()
    {
        // Path where the document will be saved.
        string filePath = @"C:\Temp\EncryptedDocument.docx";

        // Create a new blank document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello encrypted world!");

        // Configure save options to encrypt the DOCX with a password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "MySecretPassword";

        // Save the document using the encryption options.
        doc.Save(filePath, saveOptions);

        // -----------------------------------------------------------------
        // Load the encrypted document with the correct password.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("MySecretPassword");
        Document loadedDoc = new Document(filePath, loadOptions);

        // Verify that the document was loaded correctly.
        Console.WriteLine("Loaded text: " + loadedDoc.GetText().Trim());

        // -----------------------------------------------------------------
        // Demonstrate handling of an incorrect password.
        // -----------------------------------------------------------------
        try
        {
            // Attempt to load without providing the password (or with a wrong one).
            Document wrongLoad = new Document(filePath);
        }
        catch (IncorrectPasswordException)
        {
            Console.WriteLine("Failed to load document: incorrect or missing password.");
        }
    }
}
