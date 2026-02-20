using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocxToPdfConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input DOCX file path
            string inputPath = @"C:\Docs\input.docx";

            // Output PDF file path
            string outputPath = @"C:\Docs\output.pdf";

            // Load the DOCX document using the Document constructor (creates a Document from a file)
            Document doc = new Document(inputPath);

            // Create PDF save options (inherits from FixedPageSaveOptions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the document as PDF using the provided save method
            doc.Save(outputPath, pdfOptions);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
