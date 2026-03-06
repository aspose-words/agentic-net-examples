using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            const string inputPath = @"C:\Docs\Sample.docx";

            // Path where the resulting PDF will be saved.
            const string outputPath = @"C:\Docs\Sample_Converted.pdf";

            // Load the DOCX document using the Aspose.Words Document constructor.
            Document doc = new Document(inputPath);

            // Create a PdfSaveOptions object to customize PDF output.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set the PDF compliance level (e.g., PDF/A-1b for archival).
                Compliance = PdfCompliance.PdfA1b,

                // Control the quality of JPEG images embedded in the PDF.
                // Value range is 0 (lowest quality) to 100 (highest quality).
                JpegQuality = 80,

                // Choose the image compression type for all images.
                // Here we explicitly request JPEG compression.
                ImageCompression = PdfImageCompression.Jpeg,

                // Optionally, render all colors in grayscale.
                // ColorMode = ColorMode.Grayscale
            };

            // Save the document as PDF using the configured options.
            doc.Save(outputPath, pdfOptions);
        }
    }
}
