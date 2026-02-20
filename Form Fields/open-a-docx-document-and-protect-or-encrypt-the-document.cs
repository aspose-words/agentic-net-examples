using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class ProtectOrEncryptDocument
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Path where the protected/encrypted file will be saved.
        string outputPath = @"C:\Docs\output.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // -------------------------------------------------
        // 1. Write‑protect the document (requires a password to modify).
        // -------------------------------------------------
        doc.WriteProtection.SetPassword("WriteProtectPwd");
        doc.WriteProtection.ReadOnlyRecommended = true; // Suggest opening as read‑only.

        // -------------------------------------------------
        // 2. Encrypt the document when saving (requires a password to open).
        // -------------------------------------------------
        // OoxmlSaveOptions is used for DOCX/DOCM/DOTX/DOTM/FlatOpc formats.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Password = "OpenEncryptPwd"; // Password required to open the file.

        // Save the document with both write‑protection and encryption.
        doc.Save(outputPath, saveOptions);

        // -------------------------------------------------
        // 3. Demonstrate loading the encrypted document.
        // -------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("OpenEncryptPwd");
        Document loadedDoc = new Document(outputPath, loadOptions);

        // Verify that the write‑protection password is still valid.
        bool isProtected = loadedDoc.WriteProtection.IsWriteProtected;
        bool passwordValid = loadedDoc.WriteProtection.ValidatePassword("WriteProtectPwd");

        Console.WriteLine($"Write protected: {isProtected}");
        Console.WriteLine($"Write‑protection password valid: {passwordValid}");
    }
}
