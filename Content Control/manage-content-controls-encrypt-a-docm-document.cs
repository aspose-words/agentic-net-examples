using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading; // Added namespace for LoadOptions

class EncryptDocmExample
{
    static void Main()
    {
        // Folder for input and output files.
        string dataDir = @"C:\Temp\"; // <-- adjust as needed

        // Paths for the original and encrypted DOCM files.
        string originalPath = Path.Combine(dataDir, "Original.docm");
        string encryptedPath = Path.Combine(dataDir, "Encrypted.docm");

        // -------------------------------------------------
        // 1. Create a new DOCM document and add some text.
        // -------------------------------------------------
        Document doc = new Document();                     // create
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, encrypted DOCM!");

        // Save the document as DOCM (no encryption yet).
        doc.Save(originalPath, SaveFormat.Docm);           // save

        // -------------------------------------------------
        // 2. Encrypt the document while saving.
        // -------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); // create save options
        saveOptions.Password = "MySecretPassword";               // set encryption password

        // Save the same document with encryption applied.
        doc.Save(encryptedPath, saveOptions);               // save with options

        // -------------------------------------------------
        // 3. Load the encrypted document using the password.
        // -------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("MySecretPassword"); // create load options with password
        Document encryptedDoc = new Document(encryptedPath, loadOptions); // load

        // Verify that the content is accessible after decryption.
        Console.WriteLine(encryptedDoc.GetText().Trim());
    }
}
