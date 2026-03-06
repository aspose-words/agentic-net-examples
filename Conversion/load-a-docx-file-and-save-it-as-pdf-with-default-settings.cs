using System;
using Aspose.Words;

class ConvertDocxToPdf
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\Sample.docx";

        // Path where the PDF will be saved.
        string outputFile = @"C:\Docs\Sample.pdf";

        // Load the existing DOCX document.
        Document doc = new Document(inputFile);

        // Save the document as PDF. The format is inferred from the .pdf extension.
        doc.Save(outputFile);
    }
}
