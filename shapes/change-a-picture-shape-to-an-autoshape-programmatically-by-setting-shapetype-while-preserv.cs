using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // A tiny red PNG image (1x1 pixel) encoded in base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8z/C/HwAFgwJ/lKXcVwAAAABJRU5ErkJggg==";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image – this creates a picture shape (ShapeType.Image).
        Shape pictureShape = builder.InsertImage(imageBytes);

        // Preserve the original size of the picture shape.
        double originalWidth = pictureShape.Width;
        double originalHeight = pictureShape.Height;

        // Create a new rectangle AutoShape with the same size.
        Shape rectangleShape = new Shape(doc, ShapeType.Rectangle);
        rectangleShape.Width = originalWidth;
        rectangleShape.Height = originalHeight;

        // Replace the picture shape with the new rectangle shape in the document tree.
        pictureShape.ParentNode.InsertAfter(rectangleShape, pictureShape);
        pictureShape.Remove();

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
