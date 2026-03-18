using System;
using Aspose.Words;

namespace PdfToDocxConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source PDF file.
            string pdfPath = @"C:\Input\sample.pdf";

            // Path where the converted DOCX file will be saved.
            string docxPath = @"C:\Output\sample.docx";

            // Load the PDF document. The Document constructor automatically detects the format.
            Document pdfDocument = new Document(pdfPath);

            // Save the loaded document as DOCX. The Save method determines the format from the file extension.
            pdfDocument.Save(docxPath);
        }
    }
}
