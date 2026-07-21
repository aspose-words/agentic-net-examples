using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary sample image.
        string imagePath = "sample.png";
        CreateSampleImage(imagePath);

        // Create a new document and insert the image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape pictureShape = builder.InsertImage(imagePath);

        // Preserve the original size.
        double originalWidth = pictureShape.Width;
        double originalHeight = pictureShape.Height;

        // Create a rectangle AutoShape with the same size.
        Shape rectangleShape = new Shape(doc, ShapeType.Rectangle)
        {
            Width = originalWidth,
            Height = originalHeight,
            WrapType = WrapType.Inline
        };

        // Insert the rectangle after the picture and remove the picture.
        var parentParagraph = pictureShape.ParentParagraph;
        if (parentParagraph == null)
            throw new Exception("Picture shape is not inside a paragraph.");

        parentParagraph.InsertAfter(rectangleShape, pictureShape);
        pictureShape.Remove();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify the output file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Clean up the temporary image file.
        if (File.Exists(imagePath))
            File.Delete(imagePath);
    }

    private static void CreateSampleImage(string path)
    {
        // Minimal 1x1 pixel PNG (transparent)
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);
    }
}
