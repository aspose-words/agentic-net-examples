using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source PDF document.
        Document pdfDoc = new Document(@"C:\Input\source.pdf");

        // Configure image save options for high‑resolution JPEG output.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Set a high DPI resolution (e.g., 300) for better detail.
            Resolution = 300f,

            // Use the highest JPEG quality (low compression) for maximum image quality.
            // In Aspose.Words the quality is controlled by the JpegQuality property (0‑100).
            // 100 corresponds to the lowest compression (best quality).
            JpegQuality = 100,

            // Enable high‑quality rendering algorithms (optional but improves result).
            UseHighQualityRendering = true
        };

        // Save each page as a separate JPEG file. The "page_#.jpg" pattern creates one file per page.
        pdfDoc.Save(@"C:\Output\page_#.jpg", jpegOptions);
    }
}
