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

            // Load the PDF document using the Document constructor.
            Document pdfDocument = new Document(pdfPath);

            // Convert PDF to DOCX by saving with a .docx extension.
            string docxPath = @"C:\Docs\SampleConverted.docx";
            pdfDocument.Save(docxPath); // SaveFormat is inferred from the extension.

            // Convert PDF to HTML.
            string htmlPath = @"C:\Docs\SampleConverted.html";
            pdfDocument.Save(htmlPath, SaveFormat.Html);

            // Convert PDF to plain text.
            string txtPath = @"C:\Docs\SampleConverted.txt";
            pdfDocument.Save(txtPath, SaveFormat.Text);

            // Convert PDF to ODT (OpenDocument Text).
            string odtPath = @"C:\Docs\SampleConverted.odt";
            pdfDocument.Save(odtPath, SaveFormat.Odt);

            // Optional: Use SaveOptions for advanced scenarios (e.g., password protection).
            // OoxmlSaveOptions docxOptions = new OoxmlSaveOptions { Password = "Secret" };
            // pdfDocument.Save(@"C:\Docs\Protected.docx", docxOptions);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
