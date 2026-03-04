using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class EncryptDocmExample
{
    static void Main()
    {
        // Folder where the files will be written.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a blank DOCM document and add some content.
        // -----------------------------------------------------------------
        Document doc = new Document();                     // create
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This document will be encrypted.");

        // -----------------------------------------------------------------
        // 2. Set a password using OoxmlSaveOptions (ECMA376 encryption).
        // -----------------------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Password = "MySecretPassword"                 // encrypt on save
        };

        string encryptedPath = Path.Combine(artifactsDir, "EncryptedDocument.docm");
        doc.Save(encryptedPath, saveOptions);             // save with encryption

        // -----------------------------------------------------------------
        // 3. Attempt to load without a password – should throw.
        // -----------------------------------------------------------------
        try
        {
            Document loadFail = new Document(encryptedPath); // load (no password)
        }
        catch (IncorrectPasswordException)
        {
            Console.WriteLine("Failed to open without password (as expected).");
        }

        // -----------------------------------------------------------------
        // 4. Load the encrypted document using the correct password.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("MySecretPassword"); // load with password
        Document loadedDoc = new Document(encryptedPath, loadOptions); // load

        // Verify that the text was loaded correctly.
        Console.WriteLine("Loaded text: " + loadedDoc.GetText().Trim());

        // -----------------------------------------------------------------
        // 5. Optionally, re‑save without encryption (clear the password).
        // -----------------------------------------------------------------
        OoxmlSaveOptions unencryptedOptions = new OoxmlSaveOptions
        {
            Password = null // or string.Empty
        };
        string unencryptedPath = Path.Combine(artifactsDir, "DecryptedDocument.docm");
        loadedDoc.Save(unencryptedPath, unencryptedOptions); // save without encryption

        Console.WriteLine("Document encrypted and then decrypted successfully.");
    }
}
