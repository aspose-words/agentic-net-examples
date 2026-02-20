using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ProtectDocumentExample
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Apply protection that allows only form fields to be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Path where the protected document will be saved.
        string outputPath = @"C:\Docs\output_protected.docx";

        // Save the protected document.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
