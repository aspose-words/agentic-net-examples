using System;
using Aspose.Words;

class ProtectDocumentExample
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Set a write‑protection password (max 15 characters).
        doc.WriteProtection.SetPassword("MyPassword");

        // Recommend that the document be opened as read‑only.
        doc.WriteProtection.ReadOnlyRecommended = true;

        // Save the protected document.
        string outputPath = @"C:\Docs\output_protected.docx";
        doc.Save(outputPath);
    }
}
