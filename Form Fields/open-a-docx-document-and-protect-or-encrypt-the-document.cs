using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class ProtectOrEncryptDocument
{
    static void Main()
    {
        // Paths to the source and output files.
        string inputPath = @"C:\Docs\Sample.docx";
        string protectedPath = @"C:\Docs\Sample.Protected.docx";
        string encryptedPath = @"C:\Docs\Sample.Encrypted.docx";

        // Load the existing DOCX document.
        Document doc = new Document(inputPath);

        // -------------------------------------------------
        // 1. Protect the document (read‑only) with a password.
        // -------------------------------------------------
        // This adds document protection that limits editing in Microsoft Word.
        // The password is required only when the user tries to modify the document.
        doc.Protect(ProtectionType.ReadOnly, "ProtectPassword123");
        // Save the protected document.
        doc.Save(protectedPath);

        // -------------------------------------------------
        // 2. Encrypt the document with a password.
        // -------------------------------------------------
        // Encryption is applied when saving using OoxmlSaveOptions.
        // The document can be opened only after providing the correct password.
        OoxmlSaveOptions encryptOptions = new OoxmlSaveOptions
        {
            Password = "EncryptPassword456"
        };
        doc.Save(encryptedPath, encryptOptions);

        // -------------------------------------------------
        // 3. Load the encrypted document (demonstration).
        // -------------------------------------------------
        // When opening an encrypted file, supply the password via LoadOptions.
        LoadOptions loadOpts = new LoadOptions("EncryptPassword456");
        Document encryptedDoc = new Document(encryptedPath, loadOpts);

        // Verify that the document was loaded successfully.
        Console.WriteLine("Encrypted document text: " + encryptedDoc.GetText().Trim());
    }
}
