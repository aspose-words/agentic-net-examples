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

        // Create a simple red PNG image (1x1 pixel) from a base64 string
        string imagePath = Path.Combine(outputDir, "sample.png");
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8z/C/HwAFgwJ/lKXcVwAAAABJRU5ErkJggg==";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, imageBytes);

        // Create a new document and a builder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image as a floating shape with custom size, position and wrap type
        // Width = 200 points, Height = 100 points, positioned 100 points from left and top margins
        Shape imageShape = builder.InsertImage(
            imagePath,
            RelativeHorizontalPosition.Margin, 100,   // left
            RelativeVerticalPosition.Margin, 100,     // top
            200,                                      // width
            100,                                      // height
            WrapType.Square);                         // wrap type

        // Ensure the shape is not behind text
        imageShape.BehindText = false;

        // Save the document
        string docPath = Path.Combine(outputDir, "ImageShapeExample.docx");
        doc.Save(docPath);

        // Validation: ensure the output files exist
        if (!File.Exists(imagePath))
            throw new Exception("Sample image was not created.");

        if (!File.Exists(docPath))
            throw new Exception("Document was not saved.");

        // Validation: check that the shape has the expected properties
        if (imageShape.WrapType != WrapType.Square ||
            Math.Abs(imageShape.Width - 200) > 0.01 ||
            Math.Abs(imageShape.Height - 100) > 0.01 ||
            Math.Abs(imageShape.Left - 100) > 0.01 ||
            Math.Abs(imageShape.Top - 100) > 0.01)
        {
            throw new Exception("Inserted image shape does not have the expected configuration.");
        }
    }
}
