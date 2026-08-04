using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a paragraph so the document has visible content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with combined watermarks.");

        // Prepare a simple PNG image to use as the image watermark.
        // The image is a 1x1 red pixel encoded in Base64.
        string imagePath = "watermark.png";
        if (!File.Exists(imagePath))
        {
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGP8/5+hHgAHggJ/6VYV5wAAAABJRU5ErkJggg==");
            File.WriteAllBytes(imagePath, pngBytes);
        }

        // Apply a text watermark first.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Overlay an image watermark on top of the text watermark.
        // Use the overload that accepts a file path and ImageWatermarkOptions.
        ImageWatermarkOptions imgOptions = new ImageWatermarkOptions();
        doc.Watermark.SetImage(imagePath, imgOptions);

        // Save the resulting document.
        string outputPath = "WatermarkedDocument.docx";
        doc.Save(outputPath);
    }
}
