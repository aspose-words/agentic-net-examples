using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document.
        string inputPath = "input.docx";

        // Path where the PDF will be saved.
        string outputPath = "output.pdf";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Set the desired page size (e.g., A4).
        doc.FirstSection.PageSetup.PaperSize = PaperSize.A4;

        // Save the document as a PDF file.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
