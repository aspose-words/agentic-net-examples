using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

namespace ImageToPdfBatch
{
    public class Converter
    {
        /// <summary>
        /// Converts a collection of PNG images to a single PDF document.
        /// Each image is placed on its own page.
        /// </summary>
        /// <param name="pngFilePaths">Full paths to the PNG files.</param>
        /// <param name="outputPdfPath">Full path where the resulting PDF will be saved.</param>
        public static void ConvertPngsToPdf(IEnumerable<string> pngFilePaths, string outputPdfPath)
        {
            // Create a new empty Word document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            bool firstImage = true;
            foreach (string pngPath in pngFilePaths)
            {
                // Insert the PNG image. The image is inserted inline at 100% scale.
                builder.InsertImage(pngPath);

                // After each image (except the last) insert a page break so the next image starts on a new page.
                // For the first image we do not need a preceding break.
                if (!firstImage)
                {
                    // Insert a page break before the current image to ensure each image starts on a new page.
                    // This approach works even if the image itself occupies the whole page.
                    builder.InsertBreak(BreakType.PageBreak);
                }

                // Mark that we have processed at least one image.
                firstImage = false;
            }

            // Prepare PDF save options (optional – can be omitted if defaults are sufficient).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: keep the original image quality.
                ImageCompression = PdfImageCompression.Auto,
                JpegQuality = 100
            };

            // Save the assembled document as a PDF file.
            doc.Save(outputPdfPath, pdfOptions);
        }

        // Example usage.
        public static void Main()
        {
            // Define the PNG files to be merged.
            List<string> pngFiles = new List<string>
            {
                @"C:\Images\Page1.png",
                @"C:\Images\Page2.png",
                @"C:\Images\Page3.png"
            };

            // Destination PDF file.
            string outputPdf = @"C:\Output\MergedDocument.pdf";

            // Perform the conversion.
            ConvertPngsToPdf(pngFiles, outputPdf);

            Console.WriteLine("PDF created successfully at: " + outputPdf);
        }
    }
}
