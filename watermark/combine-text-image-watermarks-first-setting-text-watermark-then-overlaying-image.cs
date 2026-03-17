using System;
using System.Drawing;                     // For Color and Image handling
using Aspose.Words;                       // Core document classes
using Aspose.Words.Drawing;               // Watermark layout enums
using Aspose.Words.Saving;                // (optional) for save options if needed

class CombineWatermarks
{
    static void Main()
    {
        // Paths – adjust these to your environment
        string imagePath = @"C:\Images\Logo.png";
        string outputPath = @"C:\Output\CombinedWatermark.docx";

        // Create a new blank document
        Document doc = new Document();

        // ---------- Text watermark ----------
        TextWatermarkOptions textOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };
        doc.Watermark.SetText("Confidential", textOptions);

        // ---------- Image watermark ----------
        ImageWatermarkOptions imageOptions = new ImageWatermarkOptions
        {
            Scale = 5,          // Scale factor (5 = 500%)
            IsWashout = false   // Do not apply washout effect
        };
        // Add the image watermark on top of the existing text watermark
        doc.Watermark.SetImage(imagePath, imageOptions);

        // Save the document with both watermarks applied
        doc.Save(outputPath);
    }
}
