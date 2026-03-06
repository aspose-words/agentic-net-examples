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
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello encrypted DOCM!");

        // Configure save options to encrypt the document with a password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "Secret123";

        // Save the document as a macro‑enabled DOCM file with encryption.
        string encryptedPath = "EncryptedDocument.docm";
        doc.Save(encryptedPath, saveOptions);

        // Trying to load the encrypted file without a password throws an exception.
        try
        {
            Document _ = new Document(encryptedPath);
        }
        catch (IncorrectPasswordException)
        {
            // Expected: the document is encrypted.
        }

        // Load the encrypted document using the correct password.
        LoadOptions loadOptions = new LoadOptions("Secret123");
        Document loadedDoc = new Document(encryptedPath, loadOptions);

        // Output the document text to verify successful decryption.
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
