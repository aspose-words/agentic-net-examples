using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ShapeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a rectangle AutoShape with a size of 200x100 points.
            Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 100);

            // Set the fill color of the shape to blue.
            rectangle.FillColor = Color.Blue;

            // Define the line (stroke) weight of the shape.
            rectangle.StrokeWeight = 5.0; // Weight in points.

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RectangleShape.docx");

            // Save the document.
            doc.Save(outputPath);

            // Validate that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");

            // Optionally, inform that the operation succeeded (no console input required).
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
    }
}
