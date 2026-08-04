using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string docPath = Path.Combine(baseDir, "sample.docx");
        string imagePath = Path.Combine(baseDir, "watermark.png");
        string outputPath = Path.Combine(baseDir, "sample_with_watermark.docx");

        // Create a simple Word document if it does not already exist.
        if (!File.Exists(docPath))
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document.");
            doc.Save(docPath);
        }

        // Create a tiny PNG image to be used as the watermark.
        // The image is a 1x1 pixel transparent PNG encoded in Base64.
        if (!File.Exists(imagePath))
        {
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X3ZcAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, imageBytes);
        }

        // Load the existing document.
        Document loadedDoc = new Document(docPath);

        // Apply the image watermark using the overload that accepts a file path and options.
        ImageWatermarkOptions options = new ImageWatermarkOptions
        {
            // Use default options; you can customize Scale or IsWashout here if needed.
        };
        loadedDoc.Watermark.SetImage(imagePath, options);

        // Save the document with the watermark applied.
        loadedDoc.Save(outputPath);

        // Simple validation – confirm that the output file exists.
        Console.WriteLine(File.Exists(outputPath)
            ? $"Watermark applied successfully. Output saved to: {outputPath}"
            : "Failed to create the output document.");
    }
}
