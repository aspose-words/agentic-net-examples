using System;
using Aspose.Words;

class PdfConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string inputPdfPath = @"C:\Docs\sample.pdf";

        // Load the PDF document. The constructor automatically detects the format.
        Document pdfDocument = new Document(inputPdfPath);

        // Convert PDF to DOCX.
        pdfDocument.Save(@"C:\Docs\sample_converted.docx", SaveFormat.Docx);

        // Convert PDF to HTML.
        pdfDocument.Save(@"C:\Docs\sample_converted.html", SaveFormat.Html);

        // Convert PDF to plain text.
        pdfDocument.Save(@"C:\Docs\sample_converted.txt", SaveFormat.Text);

        // Convert PDF to XPS.
        pdfDocument.Save(@"C:\Docs\sample_converted.xps", SaveFormat.Xps);

        // Convert PDF to EPUB.
        pdfDocument.Save(@"C:\Docs\sample_converted.epub", SaveFormat.Epub);
    }
}
