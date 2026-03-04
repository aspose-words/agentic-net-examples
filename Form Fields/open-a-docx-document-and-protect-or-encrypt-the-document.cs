using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class ProtectOrEncryptDocument
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // -------------------------------------------------
        // 1. Protect the document (document protection).
        // -------------------------------------------------
        // Apply read‑only protection with a password.
        // This restricts editing in Microsoft Word but does not encrypt the file.
        doc.Protect(ProtectionType.ReadOnly, "DocProtectionPwd");

        // -------------------------------------------------
        // 2. Encrypt the document (save with a password).
        // -------------------------------------------------
        // Create save options for OOXML (DOCX) format and set the encryption password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Password = "EncryptionPwd"
        };

        // Path for the protected and encrypted output file.
        string outputPath = @"C:\Docs\ProtectedEncryptedDocument.docx";

        // Save the document using the specified options.
        doc.Save(outputPath, saveOptions);

        // -------------------------------------------------
        // 3. Demonstrate loading the encrypted document.
        // -------------------------------------------------
        // LoadOptions with the correct password are required to open the encrypted file.
        LoadOptions loadOptions = new LoadOptions("EncryptionPwd");
        Document loadedDoc = new Document(outputPath, loadOptions);

        // Verify that the document was loaded successfully.
        Console.WriteLine("Document text after loading encrypted file:");
        Console.WriteLine(loadedDoc.GetText());
    }
}
