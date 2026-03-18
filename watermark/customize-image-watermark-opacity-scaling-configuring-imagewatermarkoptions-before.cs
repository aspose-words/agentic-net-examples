using System;
using Aspose.Words;                // Core Aspose.Words namespace
using Aspose.Words.Drawing;        // For Watermark related classes

class ImageWatermarkExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Configure watermark options.
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions();
        watermarkOptions.Scale = 5;                 // Scale factor (e.g., 5 times the original size)
        watermarkOptions.IsWashout = false;        // Disable washout to make the watermark fully opaque
        // Optional: set transparency (0 = opaque, 1 = fully transparent)
        // watermarkOptions.Transparency = 0.3;    // 30% transparent (70% opaque)

        // Path to the image that will be used as a watermark.
        string imagePath = @"C:\Images\Logo.jpg";

        // Apply the image watermark with the configured options.
        // Use the overload that accepts a file path – no System.Drawing.Image required.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Save the document with the watermark applied.
        doc.Save(@"C:\Output\DocumentWithImageWatermark.docx");
    }
}
