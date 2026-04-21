using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class WatermarkDemo
{
    public static void Main()
    {
        // Define output document path.
        const string outputPath = "Watermarked.docx";

        // Define a temporary image file path for the image watermark.
        const string imagePath = "sample.png";

        // Create a simple 1x1 pixel PNG image from a base64 string.
        // This avoids using System.Drawing APIs.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XbZcAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a blank document and add some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document to demonstrate watermarks.");

        // Simulate user selection of watermark type.
        // Change this value to WatermarkType.Image to apply an image watermark.
        WatermarkType selectedWatermark = WatermarkType.Text;

        // Apply the appropriate watermark based on the selected type.
        switch (selectedWatermark)
        {
            case WatermarkType.Text:
                doc.Watermark.SetText("Confidential");
                break;

            case WatermarkType.Image:
                // Configure image watermark options if needed.
                ImageWatermarkOptions imgOptions = new ImageWatermarkOptions
                {
                    Scale = 1,          // No scaling.
                    IsWashout = false   // Keep original colors.
                };
                doc.Watermark.SetImage(imagePath, imgOptions);
                break;

            default:
                // No watermark applied.
                break;
        }

        // Save the resulting document.
        doc.Save(outputPath);

        // Clean up the temporary image file.
        if (File.Exists(imagePath))
        {
            File.Delete(imagePath);
        }

        // Optional: Output result information.
        Console.WriteLine($"Document saved to '{outputPath}' with watermark type: {selectedWatermark}");
    }
}
