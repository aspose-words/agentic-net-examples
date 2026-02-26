using System;
using Aspose.Words;

namespace PdfConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source PDF file.
            string pdfPath = @"C:\Input\sample.pdf";

            // Load the PDF document. The Document constructor automatically detects the format.
            Document pdfDocument = new Document(pdfPath);

            // Convert and save the document as DOCX.
            string docxPath = @"C:\Output\sample_converted.docx";
            pdfDocument.Save(docxPath, SaveFormat.Docx);

            // Additional conversions can be performed by changing the target extension
            // or by specifying a different SaveFormat value.

            // Example: Save as HTML.
            string htmlPath = @"C:\Output\sample_converted.html";
            pdfDocument.Save(htmlPath, SaveFormat.Html);

            // Example: Save as plain text.
            string txtPath = @"C:\Output\sample_converted.txt";
            pdfDocument.Save(txtPath, SaveFormat.Text);

            // Example: Save as EPUB.
            string epubPath = @"C:\Output\sample_converted.epub";
            pdfDocument.Save(epubPath, SaveFormat.Epub);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
