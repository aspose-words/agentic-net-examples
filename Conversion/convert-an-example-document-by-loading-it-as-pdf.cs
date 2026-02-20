using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class PdfLoadExample
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = Path.Combine("Data", "Example.pdf");

        // Configure PDF load options (optional settings can be adjusted here).
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Example: do not skip images while loading the PDF.
            SkipPdfImages = false
        };

        // Load the PDF document into an Aspose.Words Document object.
        Document doc = new Document(pdfPath, loadOptions);

        // Save the loaded document in another format (e.g., DOCX).
        string outputPath = Path.Combine("Data", "Example.docx");
        doc.Save(outputPath);
    }
}
