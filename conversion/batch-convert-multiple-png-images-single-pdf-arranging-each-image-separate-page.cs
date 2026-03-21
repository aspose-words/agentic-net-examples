using System;
using System.Collections.Generic;
using System.IO;
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
                if (!firstImage)
                {
                    // Insert a page break before the next image.
                    builder.InsertBreak(BreakType.PageBreak);
                }

                // Insert the PNG image. The image is inserted inline at 100% scale.
                builder.InsertImage(pngPath);

                firstImage = false;
            }

            // Prepare PDF save options (optional – can be omitted if defaults are sufficient).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ImageCompression = PdfImageCompression.Auto,
                JpegQuality = 100
            };

            // Save the assembled document as a PDF file.
            doc.Save(outputPdfPath, pdfOptions);
        }

        // Example usage.
        public static void Main()
        {
            // Create a temporary directory for the demo images and output PDF.
            string tempDir = Path.Combine(Path.GetTempPath(), "ImageToPdfDemo");
            Directory.CreateDirectory(tempDir);

            // Generate three simple PNG images (1x1 white pixel) from a base64 string.
            List<string> pngFiles = new List<string>();
            byte[] pngData = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lK6XAAAAAElFTkSuQmCC");

            for (int i = 1; i <= 3; i++)
            {
                string filePath = Path.Combine(tempDir, $"Page{i}.png");
                File.WriteAllBytes(filePath, pngData);
                pngFiles.Add(filePath);
            }

            // Destination PDF file.
            string outputPdf = Path.Combine(tempDir, "MergedDocument.pdf");

            // Perform the conversion.
            ConvertPngsToPdf(pngFiles, outputPdf);

            Console.WriteLine($"PDF created successfully at: {outputPdf}");
        }
    }
}
