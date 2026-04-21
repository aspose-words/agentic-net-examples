using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class CombineWatermarks
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a simple source document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document demonstrates combined text and image watermarks.");

        // 2. Create a local sample image (1x1 pixel PNG) from a base64 string
        string imagePath = Path.Combine(outputDir, "sample.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // 3. Apply a text watermark
        TextWatermarkOptions textOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Red,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };
        doc.Watermark.SetText("CONFIDENTIAL", textOptions);

        // 4. Overlay an image watermark on top of the text watermark
        ImageWatermarkOptions imageOptions = new ImageWatermarkOptions
        {
            Scale = 5,          // enlarge the image watermark
            IsWashout = false   // keep original colors
        };
        doc.Watermark.SetImage(imagePath, imageOptions);

        // 5. Save the resulting document
        string resultPath = Path.Combine(outputDir, "CombinedWatermark.docx");
        doc.Save(resultPath);

        // Simple verification that the file was created
        Console.WriteLine(File.Exists(resultPath)
            ? $"Document saved successfully: {resultPath}"
            : "Failed to save the document.");
    }
}
