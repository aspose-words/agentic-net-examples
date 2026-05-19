using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the sample document, watermark image, and the final result.
        string docPath = Path.Combine(outputDir, "Sample.docx");
        string imgPath = Path.Combine(outputDir, "watermark.png");
        string resultPath = Path.Combine(outputDir, "Watermarked.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple Word document with some content.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("The image watermark will appear behind this text.");

        // Save the initial document (optional, just to have a file on disk).
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Create a small PNG image to use as a watermark.
        //    The image is a 1x1 transparent pixel (base64 encoded).
        // -----------------------------------------------------------------
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK5cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imgPath, imageBytes);

        // -----------------------------------------------------------------
        // 3. Apply the image watermark using the Document.Watermark API.
        // -----------------------------------------------------------------
        // Use the overload that accepts a file path and ImageWatermarkOptions.
        ImageWatermarkOptions options = new ImageWatermarkOptions(); // default options
        doc.Watermark.SetImage(imgPath, options);

        // -----------------------------------------------------------------
        // 4. Save the watermarked document as DOCX.
        // -----------------------------------------------------------------
        doc.Save(resultPath, SaveFormat.Docx);

        // Simple validation that the output file exists.
        if (File.Exists(resultPath))
        {
            Console.WriteLine($"Watermarked document saved to: {resultPath}");
        }
        else
        {
            Console.WriteLine("Failed to create the watermarked document.");
        }
    }
}
