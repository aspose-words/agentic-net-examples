using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddImageWatermark
{
    static void Main()
    {
        // Path to the source WORDML document.
        string inputPath = @"C:\Docs\SourceDocument.xml";

        // Path where the watermarked document will be saved.
        string outputPath = @"C:\Docs\WatermarkedDocument.docx";

        // Path to the image that will be used as a watermark.
        string imagePath = @"C:\Images\WatermarkLogo.png";

        // Load the WORDML document.
        Document doc = new Document(inputPath);

        // Configure image watermark options (optional).
        ImageWatermarkOptions imageWatermarkOptions = new ImageWatermarkOptions
        {
            Scale = 5,          // Scale factor of the watermark image.
            IsWashout = false   // Disable washout effect if desired.
        };

        // Add the image watermark to the document.
        doc.Watermark.SetImage(imagePath, imageWatermarkOptions);

        // Save the document with the watermark applied.
        doc.Save(outputPath);
    }
}
