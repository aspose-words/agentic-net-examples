using System;
using Aspose.Words;

class ProtectDocument
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = "input.docx";

        // Path where the protected document will be saved.
        string outputPath = "output.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        doc.Save(outputPath);
    }
}
