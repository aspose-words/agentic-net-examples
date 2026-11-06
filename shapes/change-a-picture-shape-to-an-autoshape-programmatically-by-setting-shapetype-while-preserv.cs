using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple PNG image (1x1 pixel, light blue) and save it to disk
        // PNG data for a 1x1 light blue pixel
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
        string imagePath = Path.Combine(outputDir, "sample.png");
        File.WriteAllBytes(imagePath, pngData);

        // Create a new document and insert the image as a picture shape
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape pictureShape = builder.InsertImage(imagePath);

        // Preserve original size
        double originalWidth = pictureShape.Width;
        double originalHeight = pictureShape.Height;

        // Replace the picture shape with an AutoShape (Rectangle) while preserving size
        // Remove the picture shape from its parent paragraph
        Paragraph parentParagraph = (Paragraph)pictureShape.ParentNode;
        pictureShape.Remove();

        // Create a rectangle AutoShape with the same dimensions
        Shape rectangleShape = new Shape(doc, ShapeType.Rectangle);
        rectangleShape.Width = originalWidth;
        rectangleShape.Height = originalHeight;

        // Append the rectangle shape to the same paragraph
        parentParagraph.AppendChild(rectangleShape);

        // Save the document
        string docPath = Path.Combine(outputDir, "output.docx");
        doc.Save(docPath);

        // Validate that the document was saved
        if (!File.Exists(docPath))
            throw new Exception("Failed to save the output document.");

        // Indicate success (non-interactive)
        Console.WriteLine("Document saved successfully to: " + docPath);
    }
}
