using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = "input.docx";

        // Path where the resulting PDF will be saved.
        string outputFile = "output.pdf";

        // Load the DOCX document into an Aspose.Words Document object.
        Document doc = new Document(inputFile);

        // Create PDF save options (optional – customize as needed).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF using the specified options.
        doc.Save(outputFile, pdfOptions);
    }
}
