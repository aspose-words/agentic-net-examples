using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SampleDocument.docx";

        // Path where the resulting PDF will be saved.
        string pdfPath = @"C:\Docs\SampleDocument.pdf";

        // Load the DOCX document from the file system.
        Document doc = new Document(sourcePath);

        // Save the loaded document as PDF.
        // The Save method automatically determines the format from the file extension.
        doc.Save(pdfPath);
    }
}
