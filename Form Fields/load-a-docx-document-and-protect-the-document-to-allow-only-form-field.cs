using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = "input.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
