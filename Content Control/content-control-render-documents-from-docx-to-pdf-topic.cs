using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Define input and output paths.
            string inputPath = @"C:\Docs\SampleDocument.docx";   // DOCX containing content controls.
            string outputPath = @"C:\Docs\SampleDocument.pdf"; // Desired PDF output.

            // Load the DOCX document from the file system.
            Document doc = new Document(inputPath);

            // Optional: configure PDF save options (e.g., high‑quality rendering).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                UseHighQualityRendering = true,
                // Additional options can be set here if needed.
            };

            // Save the document as PDF using the specified options.
            doc.Save(outputPath, pdfOptions);

            Console.WriteLine("Document successfully rendered to PDF.");
        }
    }
}
