using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source PDF document.
        string inputPdfPath = @"C:\Docs\source.pdf";

        // Folder where the JPEG images will be saved.
        string outputFolder = @"C:\Docs\JpgPages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the PDF document.
        Document doc = new Document(inputPdfPath);

        // Create ImageSaveOptions for JPEG format.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Set high quality (100 = best quality, minimum compression).
            JpegQuality = 100,

            // Optional: increase resolution for sharper images.
            Resolution = 300,

            // Optional: enable high‑quality rendering algorithms.
            UseHighQualityRendering = true,
            UseAntiAliasing = true
        };

        // Iterate through each page of the PDF and save it as a separate JPEG file.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Select the current page for rendering.
            jpegOptions.PageSet = new PageSet(pageIndex);

            // Build the output file name (e.g., Page_1.jpg, Page_2.jpg, ...).
            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.jpg");

            // Save the selected page as a JPEG image.
            doc.Save(outputPath, jpegOptions);
        }
    }
}
