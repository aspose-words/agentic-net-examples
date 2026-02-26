using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        string inputPath = "Input.docx";
        Document doc = new Document(inputPath);

        // Apply read‑only protection with a password.
        doc.Protect(ProtectionType.ReadOnly, "ProtectPwd");

        // Set up save options to encrypt the document with a password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "EncryptPwd";

        // Save the protected and encrypted document.
        string outputPath = "ProtectedEncrypted.docx";
        doc.Save(outputPath, saveOptions);
    }
}
