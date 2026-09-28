using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a tiny PNG image (1x1 pixel) from a Base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        // Insert the image into a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape pictureShape;
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            pictureShape = builder.InsertImage(ms);
        }

        // Preserve the original size.
        double originalWidth = pictureShape.Width;
        double originalHeight = pictureShape.Height;

        // Replace the picture shape with a rectangle AutoShape while keeping its size.
        Paragraph parentParagraph = pictureShape.ParentParagraph;
        if (parentParagraph == null)
            throw new Exception("The inserted image shape does not have a parent paragraph.");

        // Create a new rectangle shape.
        Shape rectangleShape = new Shape(doc, ShapeType.Rectangle)
        {
            Width = originalWidth,
            Height = originalHeight,
            // Preserve the same wrapping as the original image.
            WrapType = pictureShape.WrapType
        };

        // Insert the rectangle after the original picture shape.
        parentParagraph.InsertAfter(rectangleShape, pictureShape);

        // Remove the original picture shape.
        pictureShape.Remove();

        // Save the document.
        string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Validation: ensure the file exists and the shape type was changed.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        Shape finalShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
        if (finalShape.ShapeType != ShapeType.Rectangle)
            throw new Exception("The shape type was not changed to Rectangle.");
    }
}
