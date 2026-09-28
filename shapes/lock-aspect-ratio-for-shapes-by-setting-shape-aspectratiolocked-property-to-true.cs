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

        // Insert a rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Lock the aspect ratio of the shape.
        shape.AspectRatioLocked = true;

        // Save the document.
        string outputPath = "AspectRatioLocked.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
