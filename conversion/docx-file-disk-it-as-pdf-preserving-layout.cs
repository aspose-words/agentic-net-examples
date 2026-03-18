using System;
using Aspose.Words;

class ConvertDocxToPdf
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path to the destination PDF file.
        string outputPath = @"C:\Docs\ConvertedDocument.pdf";

        // Load the existing DOCX document from disk.
        Document doc = new Document(inputPath);

        // Save the document as PDF. The file extension determines the format.
        doc.Save(outputPath);
    }
}
