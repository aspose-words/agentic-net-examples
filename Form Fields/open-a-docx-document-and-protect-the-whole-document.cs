using System;
using Aspose.Words;

class ProtectDocumentExample
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = "input.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Apply protection to the entire document.
        // Here we use ReadOnly protection; other types are available in the ProtectionType enum.
        doc.Protect(ProtectionType.ReadOnly);

        // Path where the protected document will be saved.
        string outputPath = "protected.docx";

        // Save the protected document.
        doc.Save(outputPath);
    }
}
