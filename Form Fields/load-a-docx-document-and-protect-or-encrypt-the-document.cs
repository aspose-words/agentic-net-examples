using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Words.Settings;

class ProtectOrEncryptDocument
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\Sample.docx";

        // Load the document. No password is needed for an unencrypted file.
        Document doc = new Document(inputPath);

        // -------------------------------------------------
        // 1. Apply document protection (read‑only) with a password.
        // -------------------------------------------------
        // This restricts editing in Microsoft Word; the document can still be edited programmatically.
        doc.Protect(ProtectionType.ReadOnly, "DocProtectPwd");

        // -------------------------------------------------
        // 2. (Optional) Apply write‑protection with a separate password.
        // -------------------------------------------------
        // Write protection adds a read‑only recommendation and requires a password to modify the file.
        doc.WriteProtection.SetPassword("WriteProtectPwd");
        doc.WriteProtection.ReadOnlyRecommended = true;

        // -------------------------------------------------
        // 3. Save the document with encryption (ECMA376 standard).
        // -------------------------------------------------
        // The OoxmlSaveOptions.Password property encrypts the file.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Password = "EncryptPwd"
        };

        string outputPath = @"C:\Docs\Sample_Protected_Encrypted.docx";
        doc.Save(outputPath, saveOptions);

        // -------------------------------------------------
        // 4. Demonstrate loading the encrypted document back.
        // -------------------------------------------------
        // The password must be supplied via LoadOptions.
        LoadOptions loadOptions = new LoadOptions("EncryptPwd");
        Document loadedDoc = new Document(outputPath, loadOptions);

        // Verify that the document is still protected.
        Console.WriteLine($"Protection type after reload: {loadedDoc.ProtectionType}");
        Console.WriteLine($"Write protection enabled: {loadedDoc.WriteProtection.IsWriteProtected}");
    }
}
