using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for sample files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeImageShapeExample");
        Directory.CreateDirectory(tempFolder);

        // Create a simple PNG image file from an embedded base‑64 string.
        string imagePath = Path.Combine(tempFolder, "sample.png");
        // This is a 16×16 pixel transparent PNG.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAAF0lEQVR42mNk+M9Qz0AEYBxVSFIAAQAA//8DAF6cB+UAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image with explicit size (points) and obtain the Shape object.
        double widthPoints = 150;   // Width in points.
        double heightPoints = 150;  // Height in points.
        Shape imageShape = builder.InsertImage(imagePath, widthPoints, heightPoints);

        // Configure wrapping, positioning, and size.
        imageShape.WrapType = WrapType.Square;
        imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
        imageShape.Left = ConvertUtil.MillimeterToPoint(20); // 20 mm from the left margin.
        imageShape.Top = ConvertUtil.MillimeterToPoint(30);  // 30 mm from the top margin.

        // Save the document.
        string outputPath = Path.Combine(tempFolder, "ImageShapeOutput.docx");
        doc.Save(outputPath);

        // Validate that the document and image shape were created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to save the output document.");

        if (imageShape.ShapeType != ShapeType.Image)
            throw new Exception("Inserted shape is not an image.");

        // Optional cleanup of the temporary image file.
        // File.Delete(imagePath);
    }
}
