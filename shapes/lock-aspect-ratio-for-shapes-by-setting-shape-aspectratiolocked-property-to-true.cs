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
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 100);
        // Lock the aspect ratio.
        shape.AspectRatioLocked = true;

        // Save the document.
        string outputPath = "AspectRatioLocked.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
