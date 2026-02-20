using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Words.Vba;
using Aspose.Words.Properties;

class VbaDocumentProtectionDemo
{
    static void Main()
    {
        // Paths to the source documents.
        string sourcePath1 = @"C:\Docs\Document1.docx";
        string sourcePath2 = @"C:\Docs\Document2.docx";

        // -----------------------------------------------------------------
        // Load the first document (may contain VBA macros).
        // -----------------------------------------------------------------
        Document doc1 = new Document(sourcePath1);

        // -----------------------------------------------------------------
        // Check if the document contains VBA macros.
        // -----------------------------------------------------------------
        if (doc1.HasMacros)
        {
            Console.WriteLine("Document1 contains VBA macros.");
            // Example: list macro module names.
            foreach (VbaModule module in doc1.VbaProject.Modules)
                Console.WriteLine($" - Module: {module.Name}");
        }

        // -----------------------------------------------------------------
        // Protect the document for editing (read‑only) and set a write‑protection password.
        // -----------------------------------------------------------------
        doc1.Protect(ProtectionType.ReadOnly, "EditPassword");
        doc1.WriteProtection.SetPassword("WritePassword");
        doc1.WriteProtection.ReadOnlyRecommended = true;

        // -----------------------------------------------------------------
        // Encrypt the document with a password when saving.
        // -----------------------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = "EncryptionPassword"
        };

        string protectedPath = @"C:\Docs\Document1_Protected.docx";
        doc1.Save(protectedPath, saveOptions);

        // -----------------------------------------------------------------
        // Verify that the saved file is encrypted.
        // -----------------------------------------------------------------
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(protectedPath);
        Console.WriteLine($"IsEncrypted: {formatInfo.IsEncrypted}");
        Console.WriteLine($"LoadFormat: {formatInfo.LoadFormat}");

        // -----------------------------------------------------------------
        // Load the encrypted document using the correct password.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("EncryptionPassword");
        Document loadedProtectedDoc = new Document(protectedPath, loadOptions);

        // -----------------------------------------------------------------
        // Load the second document for comparison.
        // -----------------------------------------------------------------
        Document doc2 = new Document(sourcePath2);

        // -----------------------------------------------------------------
        // Simple comparison: compare the plain text of both documents.
        // -----------------------------------------------------------------
        bool textsAreEqual = string.Equals(loadedProtectedDoc.GetText(), doc2.GetText(), StringComparison.Ordinal);
        Console.WriteLine($"Documents have identical text: {textsAreEqual}");

        // -----------------------------------------------------------------
        // Additional comparison: compare the number of sections.
        // -----------------------------------------------------------------
        bool sectionsAreEqual = loadedProtectedDoc.Sections.Count == doc2.Sections.Count;
        Console.WriteLine($"Documents have the same number of sections: {sectionsAreEqual}");
    }
}
