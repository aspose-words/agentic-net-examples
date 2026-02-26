using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class PdfLoadAndConvert
{
    static void Main()
    {
        // Input PDF file path
        string pdfPath = @"C:\Input\sample.pdf";

        // Output DOCX file path (conversion target)
        string docxPath = @"C:\Output\sample_converted.docx";

        // Create PdfLoadOptions – default constructor
        PdfLoadOptions loadOptions = new PdfLoadOptions();

        // Optionally configure load options (e.g., skip images, set page range)
        // loadOptions.SkipPdfImages = false;
        // loadOptions.PageIndex = 0;
        // loadOptions.PageCount = int.MaxValue;

        // Load the PDF document using the PdfLoadOptions
        Document doc = new Document(pdfPath, loadOptions);

        // Save the loaded document in DOCX format (conversion)
        doc.Save(docxPath);
    }
}
