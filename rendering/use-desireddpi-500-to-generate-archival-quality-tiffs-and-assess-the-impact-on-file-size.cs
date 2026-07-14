using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample multi‑page document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is page {i} of the sample document.");
            // Insert a placeholder image to increase file size.
            builder.InsertImage(ImageFromPlaceholder());
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save with default DPI (uses default resolution of 96 DPI).
        ImageSaveOptions defaultOptions = new ImageSaveOptions(SaveFormat.Tiff);
        defaultOptions.PageLayout = MultiPageLayout.TiffFrames(); // Multi‑page TIFF.
        string defaultPath = Path.Combine(artifactsDir, "DefaultDpi.tiff");
        doc.Save(defaultPath, defaultOptions);

        // Save with archival‑quality DPI (500).
        ImageSaveOptions highDpiOptions = new ImageSaveOptions(SaveFormat.Tiff);
        highDpiOptions.PageLayout = MultiPageLayout.TiffFrames(); // Multi‑page TIFF.
        highDpiOptions.Resolution = 500; // Set both horizontal and vertical DPI.
        string highDpiPath = Path.Combine(artifactsDir, "HighDpi.tiff");
        doc.Save(highDpiPath, highDpiOptions);

        // Compare file sizes.
        long defaultSize = new FileInfo(defaultPath).Length;
        long highDpiSize = new FileInfo(highDpiPath).Length;

        Console.WriteLine($"Default DPI TIFF size: {defaultSize} bytes");
        Console.WriteLine($"High DPI (500) TIFF size: {highDpiSize} bytes");
    }

    // Generates a simple in‑memory bitmap and returns it as a byte array.
    private static byte[] ImageFromPlaceholder()
    {
        // Use Aspose.Drawing to create a small PNG image without System.Drawing.
        using (var ms = new MemoryStream())
        {
            var bitmap = new Aspose.Drawing.Bitmap(100, 100);
            using (var graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightGray);
                graphics.DrawString(
                    "Img",
                    new Aspose.Drawing.Font("Arial", 12),
                    Aspose.Drawing.Brushes.Black,
                    new Aspose.Drawing.PointF(10, 40));
            }
            bitmap.Save(ms, Aspose.Drawing.Imaging.ImageFormat.Png);
            return ms.ToArray();
        }
    }
}
