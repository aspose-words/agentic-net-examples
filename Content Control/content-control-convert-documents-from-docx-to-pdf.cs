using System;
using Aspose.Words;

class ConvertDocxToPdf
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the PDF will be saved.
        string outputPath = "output.pdf";

        // Load the DOCX document using the Document(string) constructor.
        Document doc = new Document(inputPath);

        // Save the document as PDF. The format is inferred from the .pdf extension.
        doc.Save(outputPath);
    }
}
