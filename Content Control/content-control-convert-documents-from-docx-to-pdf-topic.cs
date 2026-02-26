using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\SampleDocument.docx";

            // Path where the resulting PDF will be saved.
            string outputPath = @"C:\Docs\SampleDocument.pdf";

            // Load the existing DOCX document.
            Document doc = new Document(inputPath);

            // Optionally configure PDF save options (e.g., compliance level).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: set PDF/A-1b compliance.
                // Compliance = PdfCompliance.PdfA1b
            };

            // Save the document as PDF. The overload with SaveOptions is used
            // to demonstrate the lifecycle rule for saving with options.
            doc.Save(outputPath, pdfOptions);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
