using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Words.Settings;

class ProtectAndEncryptDocument
{
    static void Main()
    {
        // Paths for the source and output files.
        string sourcePath = "input.docx";
        string protectedPath = "protected.docx";
        string encryptedPath = "encrypted.docx";

        // Load an existing DOCX document.
        Document doc = new Document(sourcePath);

        // -------------------------------------------------
        // 1. Protect the document (document protection, not encryption).
        // -------------------------------------------------
        // Apply read‑only protection with a password.
        doc.Protect(ProtectionType.ReadOnly, "DocProtectionPwd");

        // Save the protected document (no encryption applied here).
        doc.Save(protectedPath);

        // -------------------------------------------------
        // 2. Encrypt the document (password‑protect the file).
        // -------------------------------------------------
        // Create save options for OOXML (DOCX) format and set a password.
        OoxmlSaveOptions encryptOptions = new OoxmlSaveOptions();
        encryptOptions.Password = "FileEncryptionPwd";

        // Save the same document with encryption.
        doc.Save(encryptedPath, encryptOptions);

        // -------------------------------------------------
        // 3. Load the encrypted document using the correct password.
        // -------------------------------------------------
        LoadOptions loadOpts = new LoadOptions("FileEncryptionPwd");
        Document encryptedDoc = new Document(encryptedPath, loadOpts);

        // Verify that the document was loaded successfully.
        Console.WriteLine("Encrypted document text:");
        Console.WriteLine(encryptedDoc.GetText().Trim());

        // -------------------------------------------------
        // 4. Remove document protection (optional demonstration).
        // -------------------------------------------------
        // Unprotect using the protection password.
        bool unprotected = encryptedDoc.Unprotect("DocProtectionPwd");
        Console.WriteLine($"Document unprotected: {unprotected}");

        // Save the unprotected version (still encrypted with the file password).
        encryptedDoc.Save("unprotected_" + encryptedPath, encryptOptions);
    }
}
