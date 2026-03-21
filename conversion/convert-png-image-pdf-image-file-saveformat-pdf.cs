using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ImageToPdfConversion
{
    public static class Converter
    {
        /// <summary>
        /// Converts a PNG image file to a PDF document.
        /// </summary>
        /// <param name="pngFilePath">Full path to the source PNG image.</param>
        /// <param name="pdfFilePath">Full path where the resulting PDF will be saved.</param>
        public static void ConvertPngToPdf(string pngFilePath, string pdfFilePath)
        {
            // Create a new empty document.
            var doc = new Document();

            // Use DocumentBuilder to insert the PNG image into the document.
            var builder = new DocumentBuilder(doc);
            builder.InsertImage(pngFilePath);

            // Ensure the output directory exists.
            var pdfDir = Path.GetDirectoryName(pdfFilePath);
            if (!string.IsNullOrEmpty(pdfDir) && !Directory.Exists(pdfDir))
                Directory.CreateDirectory(pdfDir);

            // Save the document as PDF.
            doc.Save(pdfFilePath, SaveFormat.Pdf);
        }

        // Example usage.
        public static void Main()
        {
            // Create a temporary PNG file (1x1 pixel, red).
            string tempPngPath = Path.Combine(Path.GetTempPath(), "sample.png");
            const string base64Png = 
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7+X6cAAAAASUVORK5CYII=";
            File.WriteAllBytes(tempPngPath, Convert.FromBase64String(base64Png));

            // Define the output PDF path.
            string pdfPath = Path.Combine(Path.GetTempPath(), "sample.pdf");

            ConvertPngToPdf(tempPngPath, pdfPath);

            Console.WriteLine($"Conversion completed. PDF saved to: {pdfPath}");
        }
    }
}
