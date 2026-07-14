using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a temporary PNG image file (1x1 pixel) that will be inserted into the document.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        CreateSampleImage(imagePath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image as a floating shape with custom size, position and wrap type.
        // Parameters: file name, horizontal position reference, left offset,
        // vertical position reference, top offset, width, height, wrap type.
        Shape imageShape = builder.InsertImage(
            imagePath,
            RelativeHorizontalPosition.Margin, 50,   // 50 points from the left margin
            RelativeVerticalPosition.Margin, 100,    // 100 points from the top margin
            200,                                      // width in points
            150,                                      // height in points
            WrapType.Square);                         // text will wrap around the image's bounding box

        // Ensure the image appears in front of the text.
        imageShape.BehindText = false;

        // Save the document to disk.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ImageShape.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to save the output document.");

        // Clean up the temporary image file.
        File.Delete(imagePath);
    }

    // Writes a minimal PNG (1x1 pixel) to the specified path.
    private static void CreateSampleImage(string path)
    {
        // Base64-encoded PNG data for a 1x1 pixel transparent image.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);
    }
}
