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
        builder.Writeln("Hello encrypted world!");

        // Configure save options to encrypt the DOCX with a password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "Secret123";

        // Save the encrypted document.
        string filePath = "EncryptedDocument.docx";
        doc.Save(filePath, saveOptions);

        // Attempt to open the encrypted file without a password – should throw.
        try
        {
            Document _ = new Document(filePath);
        }
        catch (IncorrectPasswordException)
        {
            // Expected exception.
        }

        // Load the encrypted document using the correct password.
        LoadOptions loadOptions = new LoadOptions("Secret123");
        Document loadedDoc = new Document(filePath, loadOptions);

        // Output the document text to verify successful decryption.
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
