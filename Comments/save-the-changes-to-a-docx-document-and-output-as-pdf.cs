using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the PDF will be saved.
        string outputPath = @"C:\Docs\ConvertedDocument.pdf";

        // Load the existing DOCX document.
        Document doc = new Document(inputPath);

        // Save the document as PDF using the overload that specifies the format.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
