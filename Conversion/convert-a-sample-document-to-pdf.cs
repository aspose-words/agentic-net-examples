using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input Word document (replace with your actual file path)
        string inputFile = "SampleDocument.docx";

        // Output PDF file path
        string outputFile = "SampleDocument.pdf";

        // Load the document from the input file
        Document doc = new Document(inputFile);

        // Create PDF save options (default settings)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Example: set additional options if needed
        // pdfOptions.ColorMode = ColorMode.Grayscale; // render images in grayscale

        // Save the document as a PDF using the specified options
        doc.Save(outputFile, pdfOptions);
    }
}
