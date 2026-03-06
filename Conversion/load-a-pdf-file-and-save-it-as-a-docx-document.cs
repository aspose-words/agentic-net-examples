using System;
using Aspose.Words;

namespace PdfToDocxConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source PDF file.
            string pdfFilePath = @"C:\Path\To\Your\Input.pdf";

            // Path where the converted DOCX file will be saved.
            string docxFilePath = @"C:\Path\To\Your\Output.docx";

            // Load the PDF document. The Document constructor automatically detects the format.
            Document pdfDocument = new Document(pdfFilePath);

            // Save the loaded document as DOCX. The Save method determines the format from the file extension.
            pdfDocument.Save(docxFilePath);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
