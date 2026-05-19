using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a simple 1x1 red PNG image as a byte array (base64 encoded).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Create a new document and insert the image shape.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape pictureShape = builder.InsertImage(imageBytes);

        // Preserve the original size of the picture shape.
        double originalWidth = pictureShape.Width;
        double originalHeight = pictureShape.Height;

        // Create a new rectangle AutoShape with the same size.
        Shape rectangleShape = new Shape(doc, ShapeType.Rectangle);
        rectangleShape.Width = originalWidth;
        rectangleShape.Height = originalHeight;

        // Insert the new shape after the picture shape and then remove the picture shape.
        pictureShape.ParentNode.InsertAfter(rectangleShape, pictureShape);
        pictureShape.Remove();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);

        // Validate that the file was created and that the shape is now a rectangle.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved successfully.");

        // Additional validation: ensure the document contains a rectangle shape.
        Shape finalShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
        if (finalShape.ShapeType != ShapeType.Rectangle)
            throw new Exception("The shape was not converted to a rectangle.");
    }
}
