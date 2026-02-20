using System;
using Aspose.Words;

class ProtectDocumentExample
{
    static void Main()
    {
        // Path to the existing DOCX file to be loaded.
        string inputPath = @"C:\Docs\input.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Apply write protection with a password.
        doc.WriteProtection.SetPassword("MyPassword");

        // Recommend that the document be opened as read‑only.
        doc.WriteProtection.ReadOnlyRecommended = true;

        // Path where the protected document will be saved.
        string outputPath = @"C:\Docs\output.docx";

        // Save the protected document.
        doc.Save(outputPath);
    }
}
