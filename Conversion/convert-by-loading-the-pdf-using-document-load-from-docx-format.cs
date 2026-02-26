using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace PdfToDocxConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source PDF file.
            string pdfPath = @"C:\Input\sample.pdf";

            // Path where the resulting DOCX file will be saved.
            string docxPath = @"C:\Output\sample.docx";

            // Create load options specific for PDF files.
            // PdfLoadOptions inherits from LoadOptions and allows PDF‑specific settings.
            PdfLoadOptions loadOptions = new PdfLoadOptions();

            // Load the PDF document into an Aspose.Words Document object.
            // The constructor (string fileName, LoadOptions loadOptions) is used as per the provided rules.
            Document document = new Document(pdfPath, loadOptions);

            // Save the loaded document as DOCX.
            // The Save method with (string fileName, SaveFormat saveFormat) follows the documented API.
            document.Save(docxPath, SaveFormat.Docx);

            Console.WriteLine("PDF successfully converted to DOCX.");
        }
    }
}
