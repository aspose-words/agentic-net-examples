using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string baseDir = Directory.GetCurrentDirectory();
        string docPath = Path.Combine(baseDir, "sample.docx");
        string imagePath = Path.Combine(baseDir, "watermark.png");
        string outputPath = Path.Combine(baseDir, "output.docx");

        // -----------------------------------------------------------------
        // Create a simple Word document if it does not already exist.
        // -----------------------------------------------------------------
        if (!File.Exists(docPath))
        {
            Document blankDoc = new Document();
            var builder = new DocumentBuilder(blankDoc);
            builder.Writeln("This is a sample document.");
            blankDoc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // Create a tiny PNG image to use as a watermark.
        // The image is a 1x1 transparent pixel encoded in base64.
        // -----------------------------------------------------------------
        if (!File.Exists(imagePath))
        {
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, imageBytes);
        }

        // -----------------------------------------------------------------
        // Load the existing document.
        // -----------------------------------------------------------------
        Document doc = new Document(docPath);

        // -----------------------------------------------------------------
        // Add the image watermark using the Document.Watermark API.
        // -----------------------------------------------------------------
        var imageWatermarkOptions = new ImageWatermarkOptions(); // default options
        doc.Watermark.SetImage(imagePath, imageWatermarkOptions);

        // -----------------------------------------------------------------
        // Save the document with the watermark applied.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
