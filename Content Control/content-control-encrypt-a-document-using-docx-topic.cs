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
        builder.Writeln("Hello encrypted world!");

        // Configure save options to encrypt the DOCX with a password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Password = "Secret123";

        // Save the encrypted document.
        string filePath = "EncryptedDocument.docx";
        doc.Save(filePath, saveOptions);

        // Attempt to open the document without providing a password.
        // This should throw IncorrectPasswordException.
        try
        {
            Document wrongLoad = new Document(filePath);
        }
        catch (IncorrectPasswordException)
        {
            Console.WriteLine("Failed to open without password (as expected).");
        }

        // Load the document with the correct password.
        LoadOptions loadOptions = new LoadOptions("Secret123");
        Document loadedDoc = new Document(filePath, loadOptions);

        // Verify that the document was loaded successfully.
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
