using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary PNG image (1x1 pixel, light blue) without using System.Drawing.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        CreateSampleImage(imagePath);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image; this creates a picture shape (ShapeType.Image).
        Shape pictureShape = builder.InsertImage(imagePath);

        // Preserve the original dimensions.
        double originalWidth = pictureShape.Width;
        double originalHeight = pictureShape.Height;

        // Create a new rectangle AutoShape with the same size.
        Shape rectangleShape = new Shape(doc, ShapeType.Rectangle)
        {
            Width = originalWidth,
            Height = originalHeight
        };

        // Insert the new shape after the picture shape and then remove the picture shape.
        pictureShape.ParentNode.InsertAfter(rectangleShape, pictureShape);
        pictureShape.Remove();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved correctly.");

        // Optional cleanup of the temporary image.
        // File.Delete(imagePath);
    }

    // Writes a minimal PNG image (1x1 pixel, light blue) to the specified path.
    private static void CreateSampleImage(string path)
    {
        // Base64-encoded PNG (1x1 pixel, RGB 173,216,230 - LightBlue)
        const string base64Png = 
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, imageBytes);
    }
}
