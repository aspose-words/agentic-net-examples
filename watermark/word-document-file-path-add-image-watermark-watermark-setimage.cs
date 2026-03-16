using System;
using Aspose.Words;                       // Core Aspose.Words namespace
using Aspose.Words.Drawing;               // For ImageWatermarkOptions

class Program
{
    static void Main()
    {
        // Path to the existing Word document.
        string inputDocPath = @"C:\Docs\SourceDocument.docx";

        // Path to the image that will be used as a watermark.
        string watermarkImagePath = @"C:\Images\Watermark.png";

        // Path where the watermarked document will be saved.
        string outputDocPath = @"C:\Docs\WatermarkedDocument.docx";

        // Load the document from the file system.
        Document doc = new Document(inputDocPath);

        // Configure watermark appearance (optional).
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 5,          // Scale factor of the image.
            IsWashout = false   // Disable washout effect for a solid appearance.
        };

        // Add the image watermark to the document using the overload that accepts a file path.
        // This avoids the need for System.Drawing.Image which is not available in .NET 6+ without extra packages.
        doc.Watermark.SetImage(watermarkImagePath, watermarkOptions);

        // Save the modified document.
        doc.Save(outputDocPath);
    }
}
