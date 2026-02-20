using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input DOCX file path
        string inputPath = "input.docx";

        // Output PDF file path
        string outputPath = "output.pdf";

        // Load the DOCX document
        Document doc = new Document(inputPath);

        // Configure PDF save options (customize as needed)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Example: embed all fonts in the PDF
        pdfOptions.EmbedFullFonts = true;
        // Example: set PDF/A compliance (uncomment if required)
        // pdfOptions.Compliance = PdfCompliance.PdfA2b;

        // Save the document as PDF using the specified options
        doc.Save(outputPath, pdfOptions);
    }
}
