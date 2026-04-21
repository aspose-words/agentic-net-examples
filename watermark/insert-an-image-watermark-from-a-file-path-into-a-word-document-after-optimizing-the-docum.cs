using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Define file names for the sample image and the output document.
        const string imagePath = "sample.png";
        const string outputPath = "WatermarkedDocument.docx";

        // Create a simple 1x1 pixel PNG image from a Base64 string.
        // This avoids using System.Drawing or other image generation APIs.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lK5XAAAAAElFTkSuQmCC");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new blank Word document.
        Document doc = new Document();

        // Optimize the document for a specific Word version (optional but fulfills the "after optimizing" requirement).
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

        // Configure image watermark options.
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            // Scale = 0 means auto-scaling to fit the page margins.
            Scale = 0,
            // Disable washout effect so the image appears with original colors.
            IsWashout = false
        };

        // Apply the image watermark using the file path.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Save the resulting document.
        doc.Save(outputPath);
    }
}
