using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "Document.docx");
        string imagePath = Path.Combine(outputDir, "watermark.png");
        string resultPath = Path.Combine(outputDir, "Watermarked.docx");

        // Create a minimal PNG image (1x1 pixel) from a Base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, imageBytes);

        // Create a blank document and add a simple paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with an image watermark.");

        // Configure image watermark options (optional).
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 5,          // Scale factor for the watermark.
            IsWashout = false   // Disable washout effect.
        };

        // Apply the image watermark using the file path.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Save the watermarked document as DOCX.
        doc.Save(resultPath, SaveFormat.Docx);
    }
}
