using System;
using Aspose.Words;

class ConvertDocxToPdf
{
    static void Main()
    {
        // Path to the source DOCX file
        string inputFile = "input.docx";

        // Path to the destination PDF file
        string outputFile = "output.pdf";

        // Load the DOCX document from the file system
        Document doc = new Document(inputFile);

        // Save the document as PDF; the format is inferred from the .pdf extension
        doc.Save(outputFile);
    }
}
