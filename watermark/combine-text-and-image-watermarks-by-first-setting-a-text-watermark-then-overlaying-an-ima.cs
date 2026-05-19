using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // ---------- Text watermark ----------
        // Configure text watermark options.
        TextWatermarkOptions textOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = System.Drawing.Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };
        // Apply the text watermark.
        doc.Watermark.SetText("CONFIDENTIAL", textOptions);

        // ---------- Image watermark ----------
        // Prepare a small sample PNG image (1x1 pixel red dot) as a local file.
        const string imagePath = "sample.png";
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Configure image watermark options.
        ImageWatermarkOptions imageOptions = new ImageWatermarkOptions
        {
            Scale = 5,          // Larger than default.
            IsWashout = false   // Do not apply washout effect.
        };
        // Apply the image watermark on top of the existing text watermark.
        doc.Watermark.SetImage(imagePath, imageOptions);

        // Save the resulting document.
        const string outputPath = "CombinedWatermark.docx";
        doc.Save(outputPath);

        // Clean up the temporary image file.
        if (File.Exists(imagePath))
            File.Delete(imagePath);
    }
}
