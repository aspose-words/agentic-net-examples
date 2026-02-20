using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class ProtectOrEncryptDocument
{
    static void Main()
    {
        // Paths to the source and destination files.
        string sourcePath = @"C:\Docs\input.docx";
        string protectedPath = @"C:\Docs\protected.docx";

        // Password that will be used for write‑protection and/or encryption.
        string password = "MyPassword";

        // -----------------------------------------------------------------
        // Load the original (unprotected) document.
        // No password is required for loading because the source file is not encrypted.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions();               // default options
        Document doc = new Document(sourcePath, loadOptions);

        // -----------------------------------------------------------------
        // Apply write‑protection (recommended read‑only + password).
        // This does NOT encrypt the file contents; it only sets a password
        // required to modify the document.
        // -----------------------------------------------------------------
        doc.WriteProtection.SetPassword(password);
        doc.WriteProtection.ReadOnlyRecommended = true;

        // -----------------------------------------------------------------
        // Encrypt the document when saving.
        // OoxmlSaveOptions.Password encrypts the file using the ECMA376
        // standard encryption algorithm.
        // -----------------------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Password = password               // encrypt the file with the same password
        };

        // Save the protected and encrypted document.
        doc.Save(protectedPath, saveOptions);

        // -----------------------------------------------------------------
        // Demonstrate loading the encrypted document using the password.
        // -----------------------------------------------------------------
        LoadOptions loadProtected = new LoadOptions(password);
        Document loadedProtected = new Document(protectedPath, loadProtected);

        // Verify that the document was loaded successfully.
        Console.WriteLine("Document loaded. Text length: " + loadedProtected.GetText().Length);
    }
}
