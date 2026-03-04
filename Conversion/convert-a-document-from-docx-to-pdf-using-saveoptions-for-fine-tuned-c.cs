using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input DOCX file path.
            string inputPath = @"C:\Docs\SampleDocument.docx";

            // Output PDF file path.
            string outputPath = @"C:\Docs\SampleDocument.pdf";

            // Load the DOCX document.
            Document doc = new Document(inputPath);

            // Create PDF save options for fine‑tuned control.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: enable high‑quality rendering (slower but better output).
                UseHighQualityRendering = true,

                // Example: reduce memory usage for large documents.
                MemoryOptimization = true,

                // Example: set the number of outline levels to include in the PDF bookmark pane.
                OutlineOptions = { HeadingsOutlineLevels = 3 },

                // Example: set the PDF compliance level (optional).
                // Compliance = PdfCompliance.PdfA2b
            };

            // Save the document as PDF using the specified options.
            doc.Save(outputPath, pdfOptions);
        }
    }
}
