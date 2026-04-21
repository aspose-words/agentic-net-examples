using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle AutoShape with width 100 points and height 50 points.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Set the fill color of the rectangle to blue.
        rectangle.FillColor = Color.Blue;

        // Define the line (stroke) weight of the rectangle.
        rectangle.StrokeWeight = 2.0;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "RectangleShape.docx");
        doc.Save(outputPath);

        // Verify that the file was saved successfully.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
