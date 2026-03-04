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

        // Apply protection to the whole document.
        // Here we use ReadOnly protection, which prevents any changes.
        doc.Protect(ProtectionType.ReadOnly);

        // Path where the protected document will be saved.
        string outputPath = @"C:\Docs\protected.docx";

        // Save the protected document.
        doc.Save(outputPath);
    }
}
