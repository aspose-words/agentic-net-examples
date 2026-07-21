using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle AutoShape with specified width and height (in points).
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 100);

        // Set the fill color of the shape to blue.
        rectangle.FillColor = Color.Blue;

        // Define the line (stroke) weight of the shape.
        rectangle.StrokeWeight = 5.0; // Weight in points.

        // Optionally, set the stroke color to make the line visible.
        rectangle.StrokeColor = Color.Black;

        // Save the document to a file.
        string outputPath = "RectangleShape.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Inform that the operation completed successfully.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
