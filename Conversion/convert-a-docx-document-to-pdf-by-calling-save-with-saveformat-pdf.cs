using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = "output.pdf";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Convert and save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
