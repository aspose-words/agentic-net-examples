using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape with initial width and height.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 100);

        // Lock the shape's aspect ratio.
        shape.AspectRatioLocked = true;

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AspectRatioLocked.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
