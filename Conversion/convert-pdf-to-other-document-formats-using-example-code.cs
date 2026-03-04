using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfConversionExample
{
    static void Main()
    {
        // Path to the source PDF file.
        string inputPdfPath = @"C:\Docs\Input.pdf";

        // Load the PDF document. The constructor automatically detects the format.
        Document pdfDoc = new Document(inputPdfPath);

        // Convert the whole PDF to a DOCX file.
        pdfDoc.Save(@"C:\Docs\Output.docx", SaveFormat.Docx);

        // Convert the PDF to an HTML file.
        pdfDoc.Save(@"C:\Docs\Output.html", SaveFormat.Html);

        // Convert the PDF to plain text.
        pdfDoc.Save(@"C:\Docs\Output.txt", SaveFormat.Text);

        // Convert the PDF to an OpenDocument Text file.
        pdfDoc.Save(@"C:\Docs\Output.odt", SaveFormat.Odt);

        // Convert the PDF to an EPUB file.
        pdfDoc.Save(@"C:\Docs\Output.epub", SaveFormat.Epub);

        // --------------------------------------------------------------------
        // Example: Load only the first page of the PDF and convert that page.
        // --------------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            PageIndex = 0,   // Zero‑based index of the first page to read.
            PageCount = 1    // Number of pages to read.
        };

        // Load the PDF with the specified load options.
        Document firstPageDoc = new Document(inputPdfPath, loadOptions);

        // Save the extracted page as a DOCX document.
        firstPageDoc.Save(@"C:\Docs\FirstPage.docx", SaveFormat.Docx);
    }
}
