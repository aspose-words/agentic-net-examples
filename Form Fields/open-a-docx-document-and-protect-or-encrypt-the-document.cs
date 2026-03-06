using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ProtectOrEncryptDocument
{
    static void Main()
    {
        // Paths to the source and result files.
        const string inputPath = @"C:\Docs\Source.docx";
        const string protectedPath = @"C:\Docs\Protected.docx";
        const string encryptedPath = @"C:\Docs\Encrypted.docx";

        // Load the existing DOCX document.
        Document doc = new Document(inputPath);

        // -------------------------------------------------
        // 1. Protect the document (document protection).
        // -------------------------------------------------
        // Apply read‑only protection and set a password that Word will ask for.
        doc.Protect(ProtectionType.ReadOnly, "ProtectPassword");

        // Save the protected document (no encryption applied here).
        doc.Save(protectedPath);

        // -------------------------------------------------
        // 2. Encrypt the document (password protection on save).
        // -------------------------------------------------
        // Create save options for OOXML (DOCX) and assign a password.
        OoxmlSaveOptions encryptOptions = new OoxmlSaveOptions
        {
            Password = "EncryptPassword"
        };

        // Save the same document using the encryption options.
        doc.Save(encryptedPath, encryptOptions);

        // -------------------------------------------------
        // Optional: Demonstrate write protection (different from document protection).
        // -------------------------------------------------
        // Set a write‑protection password and recommend read‑only opening.
        doc.WriteProtection.SetPassword("WriteProtectPwd");
        doc.WriteProtection.ReadOnlyRecommended = true;

        // Save the document with write protection (still using the same encryption password if desired).
        // Here we reuse the same encrypted file path for illustration.
        doc.Save(encryptedPath, encryptOptions);
    }
}
