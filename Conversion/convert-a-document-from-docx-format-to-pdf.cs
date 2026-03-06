using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path to the destination PDF file.
        string outputPath = @"C:\Docs\ConvertedDocument.pdf";

        // Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Save the document as PDF. The format is inferred from the file extension.
        doc.Save(outputPath);
    }
}
