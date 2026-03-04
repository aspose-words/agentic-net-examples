using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPdfConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeWordsPdfConversion <inputPath> <outputPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            var converter = new PdfConverter();
            converter.ConvertToPdf(inputPath, outputPath);
            Console.WriteLine($"PDF saved to {outputPath}");
        }
    }

    public class PdfConverter
    {
        /// <summary>
        /// Converts a Word document to PDF using custom PdfSaveOptions.
        /// </summary>
        /// <param name="inputPath">Full path to the source .doc/.docx file.</param>
        /// <param name="outputPath">Full path where the resulting PDF will be saved.</param>
        public void ConvertToPdf(string inputPath, string outputPath)
        {
            // Load the source document.
            Document doc = new Document(inputPath);

            // Create a PdfSaveOptions instance via the factory method.
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

            // Cast to PdfSaveOptions to access PDF‑specific properties.
            PdfSaveOptions pdfOptions = (PdfSaveOptions)saveOptions;

            // ---- Custom PDF output options ----
            pdfOptions.Compliance = PdfCompliance.PdfA1b;                     // PDF/A‑1b compliance
            pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard; // Export custom properties as standard entries
            pdfOptions.ExportDocumentStructure = true;                       // Preserve document structure (tags)
            pdfOptions.MemoryOptimization = true;                           // Optimize memory usage for large docs
            pdfOptions.PageMode = PdfPageMode.UseOutlines;                  // Show outlines/bookmarks on open
            pdfOptions.ZoomFactor = 150;                                    // Open at 150 % zoom

            // Save the document to PDF using the configured options.
            doc.Save(outputPath, pdfOptions);
        }
    }
}
