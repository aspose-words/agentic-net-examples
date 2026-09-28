using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary image file (a tiny PNG).
        string imagePath = Path.Combine(Path.GetTempPath(), "sample_image.png");
        CreateSampleImage(imagePath);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image with explicit size (in points).
        // 1 point = 1/72 inch. Convert pixels to points assuming 96 DPI.
        double widthInPoints = 200 * 72.0 / 96.0;
        double heightInPoints = 100 * 72.0 / 96.0;
        Shape imageShape = builder.InsertImage(imagePath, widthInPoints, heightInPoints);

        // Configure wrapping, positioning and relative references.
        imageShape.WrapType = WrapType.Square;
        imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        imageShape.Left = ConvertUtil.MillimeterToPoint(20);   // 20 mm from the left of the page
        imageShape.Top = ConvertUtil.MillimeterToPoint(30);    // 30 mm from the top of the page

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ImageShapeExample.docx");
        doc.Save(outputPath);

        // Validate that the document was saved.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Validate that the inserted shape has the expected wrap type.
        if (imageShape.WrapType != WrapType.Square)
            throw new InvalidOperationException("Wrap type was not set correctly.");

        // Clean up temporary image file.
        if (File.Exists(imagePath))
            File.Delete(imagePath);
    }

    // Writes a minimal PNG image (1x1 pixel) to the specified path.
    private static void CreateSampleImage(string path)
    {
        // Base64-encoded PNG (transparent 1x1 pixel).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);
    }
}
