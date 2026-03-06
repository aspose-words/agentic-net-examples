using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class EncryptDocxExample
{
    static void Main()
    {
        // Folder where the document will be saved.
        string artifactsDir = @"C:\Temp\Artifacts\";
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Configure save options to encrypt the DOCX with a password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "MyPassword";

        // Save the encrypted document.
        string encryptedPath = Path.Combine(artifactsDir, "Encrypted.docx");
        doc.Save(encryptedPath, saveOptions);

        // -----------------------------------------------------------------
        // Demonstrate that loading without a password throws an exception.
        try
        {
            Document wrong = new Document(encryptedPath);
        }
        catch (IncorrectPasswordException)
        {
            // Expected: the document is encrypted and a password is required.
        }

        // Load the encrypted document using the correct password.
        LoadOptions loadOptions = new LoadOptions("MyPassword");
        Document loadedDoc = new Document(encryptedPath, loadOptions);

        // Verify that the content was loaded correctly.
        string text = loadedDoc.GetText().Trim(); // Should be "Hello world!"
        Console.WriteLine($"Loaded text: {text}");
    }
}
