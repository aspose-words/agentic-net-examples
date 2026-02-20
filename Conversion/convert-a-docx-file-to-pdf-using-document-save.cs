using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("input.docx");

        // Optionally configure PDF save options (e.g., high quality rendering).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            UseHighQualityRendering = true,
            // Additional options can be set here if needed.
        };

        // Save the document as a PDF file using the specified options.
        doc.Save("output.pdf", pdfOptions);
    }
}
