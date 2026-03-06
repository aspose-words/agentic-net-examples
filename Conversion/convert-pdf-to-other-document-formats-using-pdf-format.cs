using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source PDF file.
            string pdfPath = @"C:\Docs\Sample.pdf";

            // Load the PDF document using the Document constructor that accepts a file name.
            Document pdfDoc = new Document(pdfPath);

            // Convert PDF to DOCX.
            pdfDoc.Save(@"C:\Docs\Sample_Converted.docx", SaveFormat.Docx);

            // Convert PDF to HTML.
            pdfDoc.Save(@"C:\Docs\Sample_Converted.html", SaveFormat.Html);

            // Convert PDF to plain text.
            pdfDoc.Save(@"C:\Docs\Sample_Converted.txt", SaveFormat.Text);

            // Convert PDF to EPUB.
            pdfDoc.Save(@"C:\Docs\Sample_Converted.epub", SaveFormat.Epub);

            // Convert PDF to ODT.
            pdfDoc.Save(@"C:\Docs\Sample_Converted.odt", SaveFormat.Odt);

            // Example using XpsSaveOptions for PDF to XPS conversion.
            XpsSaveOptions xpsOptions = new XpsSaveOptions();
            pdfDoc.Save(@"C:\Docs\Sample_Converted.xps", xpsOptions);
        }
    }
}
