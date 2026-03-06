using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting PDF will be saved.
        string pdfPath = @"C:\Docs\ConvertedDocument.pdf";

        // Load the DOCX document from the file system.
        Document doc = new Document(sourcePath);

        // Save the document as PDF.
        // The file extension ".pdf" tells Aspose.Words to use the PDF format automatically.
        doc.Save(pdfPath);
    }
}
