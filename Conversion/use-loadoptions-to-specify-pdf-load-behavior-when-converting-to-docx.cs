using System;
using Aspose.Words;
using Aspose.Words.Loading;

class PdfToDocxConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\sample.pdf";

        // Path where the resulting DOCX file will be saved.
        string docxPath = @"C:\Output\sample.docx";

        // Create a PdfLoadOptions object to control how the PDF is loaded.
        PdfLoadOptions loadOptions = new PdfLoadOptions();

        // Example options:
        // Skip loading images from the PDF.
        loadOptions.SkipPdfImages = true;

        // Load only the first page of the PDF (zero‑based index).
        loadOptions.PageIndex = 0;
        loadOptions.PageCount = 1;

        // Load the PDF document using the specified load options.
        Document document = new Document(pdfPath, loadOptions);

        // Save the loaded document as DOCX.
        document.Save(docxPath);
    }
}
