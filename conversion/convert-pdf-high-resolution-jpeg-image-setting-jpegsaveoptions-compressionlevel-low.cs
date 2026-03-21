using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Prepare input and output paths relative to the current directory.
        string currentDir = Directory.GetCurrentDirectory();
        string inputPath = Path.Combine(currentDir, "source.pdf");
        string outputDir = Path.Combine(currentDir, "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple document and save it as PDF to serve as the source PDF.
        Document tempDoc = new Document();
        tempDoc.FirstSection.Body.AppendChild(new Paragraph(tempDoc));
        tempDoc.FirstSection.Body.FirstParagraph.AppendChild(new Run(tempDoc, "Sample PDF content for conversion."));
        tempDoc.Save(inputPath, SaveFormat.Pdf);

        // Load the source PDF document.
        Document pdfDoc = new Document(inputPath);

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
        string outputPattern = Path.Combine(outputDir, "page_#.jpg");
        pdfDoc.Save(outputPattern, jpegOptions);
    }
}
