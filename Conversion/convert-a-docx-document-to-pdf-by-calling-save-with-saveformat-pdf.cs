using System;
using Aspose.Words;

class DocxToPdfConverter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\ConvertedDocument.pdf";

        // Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Save the document as PDF using the SaveFormat.Pdf overload.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
