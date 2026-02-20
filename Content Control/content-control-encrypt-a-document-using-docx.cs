using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class EncryptDocxExample
{
    static void Main()
    {
        // Folder where the files will be saved.
        string artifactsDir = @"C:\Temp\";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This document will be encrypted.");

        // Configure save options to encrypt the DOCX with a password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "MyPassword"; // ECMA376 Standard encryption.

        // Save the encrypted document.
        string encryptedPath = artifactsDir + "EncryptedDocument.docx";
        doc.Save(encryptedPath, saveOptions);

        // Attempt to load the document without a password – will throw IncorrectPasswordException.
        try
        {
            Document loadFail = new Document(encryptedPath);
        }
        catch (IncorrectPasswordException)
        {
            Console.WriteLine("Failed to open without password (as expected).");
        }

        // Load the encrypted document using the correct password.
        LoadOptions loadOptions = new LoadOptions("MyPassword");
        Document loadedDoc = new Document(encryptedPath, loadOptions);

        // Verify that the content was loaded correctly.
        Console.WriteLine("Document text after loading with password:");
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
