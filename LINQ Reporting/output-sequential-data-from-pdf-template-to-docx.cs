using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfToDocxConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source PDF template.
            string pdfTemplatePath = @"C:\Input\Template.pdf";

            // Path where the resulting DOCX will be saved.
            string outputDocxPath = @"C:\Output\Result.docx";

            // Load the PDF document using the Document constructor that accepts a file name.
            Document pdfDocument = new Document(pdfTemplatePath);

            // Save the loaded document as DOCX using the Save method overload that specifies the format.
            pdfDocument.Save(outputDocxPath, SaveFormat.Docx);

            // Optional: inform the user that the conversion has completed.
            Console.WriteLine("PDF template has been successfully converted to DOCX.");
        }
    }
}
