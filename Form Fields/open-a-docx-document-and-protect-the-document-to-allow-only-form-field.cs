using System;
using Aspose.Words;

class ProtectDocumentExample
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Path where the protected document will be saved.
        string outputPath = @"C:\Docs\output_protected.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        doc.Save(outputPath);
    }
}
