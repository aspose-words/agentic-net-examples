using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Create a simple 1x1 PNG image (transparent) from a Base64 string.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lKXcAAAAAElFTkSuQmCC";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.Combine(outputFolder, "watermark.png");
        File.WriteAllBytes(imagePath, imageBytes);

        // Create a new blank document and add some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains an image watermark.");

        // Configure image watermark options (optional).
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 0.5,          // Scale the image to 50% of its original size.
            IsWashout = false    // Do not apply washout effect.
        };

        // Apply the image watermark using the file path.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Save the watermarked document as DOCX.
        string outputDocPath = Path.Combine(outputFolder, "Watermarked.docx");
        doc.Save(outputDocPath, SaveFormat.Docx);
    }
}
